"""Conservative, page-local removal of user-selected PDF painting operations.

No rectangle redaction, shared resource replacement, or image reconstruction.
Analysis owns immutable input bytes; preview and export apply the same edits.
"""

from dataclasses import dataclass
import hashlib
import os
import re
import uuid

import fitz

from ratools_pdf.pdf.hyperlink_styles import _content_operations
from ratools_pdf.services.file_safety import check_unlinked_path, recycle_owned_path


class WatermarkError(ValueError):
    pass


@dataclass(frozen=True)
class Candidate:
    id: str
    page: int  # zero based
    kind: str
    label: str
    start: int = 0
    end: int = 0
    replacement: bytes = b''
    annot_xrefs: tuple = ()


@dataclass(frozen=True)
class Analysis:
    path: str
    source: bytes
    digest: str
    page_count: int
    pages: tuple
    candidates: tuple
    notes: tuple


def parse_pages(expression, page_count):
    """Accept one-based inclusive ranges, rejecting typos instead of broadening."""
    if not expression.strip():
        return tuple(range(page_count))
    selected = set()
    for part in expression.replace('，', ',').split(','):
        match = re.fullmatch(r'\s*([0-9]+)(?:\s*-\s*([0-9]+))?\s*', part)
        if not match:
            raise WatermarkError('页码格式应为 1-3,5；留空表示全部页面。')
        first = int(match[1])
        last = int(match[2] or first)
        if not 1 <= first <= last <= page_count:
            raise WatermarkError(f'页码范围必须在 1–{page_count} 内，且起始页不大于结束页。')
        selected.update(range(first - 1, last))
    return tuple(sorted(selected))


def _check_cancel(cancel):
    if cancel and cancel():
        raise WatermarkError('操作已取消；未发布输出。')


def _contents(doc, page):
    return b'\n'.join(doc.xref_stream(xref) for xref in page.get_contents())


def _annotation_refs(doc, page):
    kind, value = doc.xref_get_key(page.xref, 'Annots')
    if kind == 'null':
        return []
    if kind == 'xref':
        value = doc.xref_object(int(value.split()[0]), compressed=True)
    if not re.fullmatch(r'\[\s*(?:\d+\s+\d+\s+R\s*)*\]', value):
        raise WatermarkError('批注数组含内嵌或复杂对象，已跳过批注候选。')
    return [(int(xref), int(generation)) for xref, generation in
            re.findall(r'(\d+)\s+(\d+)\s+R', value)]


_PAINT = {'S', 's', 'f', 'F', 'f*', 'B', 'B*', 'b', 'b*'}
_KNOWN = set('q Q cm w J j M d ri i gs m l c v y h re W W* n BT ET Tc Tw Tz TL Tf Tr Ts Td TD Tm T* Tj TJ SC SCN sc scn G g RG rg K k CS cs Do BMC BDC EMC MP DP'.split()) | _PAINT | {"'", '"'}


def _painting_candidates(doc, page):
    data = _contents(doc, page)
    text = data.decode('latin1')  # byte offsets remain exact
    operations = list(_content_operations(text))
    if operations:
        tail = re.sub(r'%[^\r\n]*', '', text[operations[-1][2]:]).strip()
        if tail:
            raise WatermarkError('内容流尾部存在未完成指令。')
    elif text.strip() and re.sub(r'%[^\r\n]*', '', text).strip():
        raise WatermarkError('内容流缺少完整指令。')

    resources = {}
    for item in page.get_images(full=True):
        if item[-1] == 0:
            resources[item[7]] = ('图片', item[0])
    for xref, name, invoker, _bbox in page.get_xobjects():
        if invoker == 0:
            resources[name] = ('绘图／组合对象', xref)

    candidates = []
    stack, marked = [], 0
    render_mode = 0
    active_text = None
    path_active = False
    notes = set()
    for operator, start, end in operations:
        if operator not in _KNOWN:
            raise WatermarkError('包含首版不支持的绘制指令，已跳过该页内容对象。')
        command = text[start:end]
        if operator == 'q':
            stack.append(render_mode)
        elif operator == 'Q':
            if not stack:
                raise WatermarkError('图形状态不平衡。')
            render_mode = stack.pop()
        elif operator in ('BMC', 'BDC'):
            marked += 1
        elif operator == 'EMC':
            marked -= 1
            if marked < 0:
                raise WatermarkError('标记内容不平衡。')
        elif operator == 'Tr':
            match = re.fullmatch(r'\s*([0-7])\s+Tr', command)
            if not match:
                raise WatermarkError('无法判定文字渲染模式。')
            render_mode = int(match[1])
        elif operator == 'BT':
            if active_text is not None:
                raise WatermarkError('文字对象嵌套。')
            active_text = []
        elif operator in ('Tj', 'TJ', "'", '"'):
            if active_text is None:
                raise WatermarkError('文字绘制缺少完整文字对象。')
            active_text.append((operator, start, end, render_mode))
        elif operator == 'ET':
            if active_text is None:
                raise WatermarkError('文字对象不平衡。')
            # Leave text/graphics state instructions intact. With just one show,
            # dropping Tj/TJ cannot shift subsequent text within this BT/ET.
            if len(active_text) == 1 and active_text[0][0] in ('Tj', 'TJ') and active_text[0][3] in (0, 1, 2):
                _, begin, finish, _ = active_text[0]
                candidates.append(('文字', '独立文字对象', begin, finish, b''))
            elif active_text:
                notes.add('部分文字包含多次绘制、隐藏文字或裁剪文字，未列为候选。')
            active_text = None
        elif operator == 'Do':
            if active_text is not None:
                raise WatermarkError('文字对象内嵌绘图，首版不支持。')
            match = re.fullmatch(r'\s*/([^\s/<>\[\]()]+)\s+Do', command)
            if not match or match[1] not in resources:
                raise WatermarkError('无法可靠定位绘图资源。')
            kind, xref = resources[match[1]]
            if kind == '图片':
                placements = page.get_image_rects(xref)
                if any((rect & page.rect).get_area() >= page.rect.get_area() * .9 for rect in placements):
                    notes.add('覆盖整页的图片未列为候选；无法从扫描图像中分离水印。')
                    continue
            candidates.append((kind, kind + '对象', start, end, b''))
        elif operator in ('m', 're'):
            path_active = True
        elif operator in _PAINT:
            if active_text is not None:
                raise WatermarkError('文字和路径混排，首版不支持。')
            if path_active:
                # n ends the path without painting, preserving graphics state
                # and any clipping operation that other content depends on.
                candidates.append(('绘图', '独立路径对象', start, end, b'n'))
            path_active = False
        elif operator == 'n':
            path_active = False
    if stack or marked or active_text is not None or path_active:
        raise WatermarkError('内容流存在未闭合状态，已跳过该页内容对象。')
    return candidates, sorted(notes)


def analyze(path, page_expression='', cancel=None):
    path = check_unlinked_path(path)
    source = path.read_bytes()
    candidates, notes = [], []
    with fitz.open(stream=source, filetype='pdf') as doc:
        if doc.needs_pass:
            raise WatermarkError('请先在安全中心解锁 PDF，再使用去水印工具。')
        if not doc.is_pdf or doc.page_count == 0:
            raise WatermarkError('需要包含页面的 PDF 文件。')
        if doc.is_repaired:
            raise WatermarkError('PDF 需要修复，首版不处理已修复读取的文件。')
        pages = parse_pages(page_expression, doc.page_count)
        for number in pages:
            _check_cancel(cancel)
            page = doc[number]
            try:
                refs = {xref for xref, _generation in _annotation_refs(doc, page)}
                for annot in page.annots() or ():
                    if annot.type[0] not in (0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 25):
                        continue
                    if annot.xref not in refs:
                        continue
                    related = tuple(x for x in (annot.xref, annot.popup_xref) if x in refs)
                    candidates.append(Candidate(f'{number}:a:{annot.xref}', number, '批注',
                                                f'{annot.type[1]} 批注', annot_xrefs=related))
            except (ValueError, RuntimeError):
                notes.append(f'第 {number + 1} 页：批注结构无法安全分离，已跳过。')
            try:
                paints, skipped = _painting_candidates(doc, page)
                for index, (kind, label, start, end, replacement) in enumerate(paints):
                    candidates.append(Candidate(f'{number}:c:{index}', number, kind,
                                                f'{label} {index + 1}', start, end, replacement))
                notes.extend(f'第 {number + 1} 页：{note}' for note in skipped)
            except (ValueError, RuntimeError):
                notes.append(f'第 {number + 1} 页：复杂或不支持的内容流，已跳过文字／图片／绘图候选。')
        return Analysis(str(path), source, hashlib.sha256(source).hexdigest(),
                        doc.page_count, pages, tuple(candidates), tuple(notes))


def _selection(analysis, selected):
    ids = set(selected)
    known = {candidate.id: candidate for candidate in analysis.candidates}
    if not ids <= known.keys():
        raise WatermarkError('选择已失效，请重新分析。')
    return [candidate for candidate in analysis.candidates if candidate.id in ids]


def _apply(doc, candidates, cancel=None):
    for number in sorted({candidate.page for candidate in candidates}):
        _check_cancel(cancel)
        page = doc[number]
        selected = [candidate for candidate in candidates if candidate.page == number]
        contents = [candidate for candidate in selected if not candidate.annot_xrefs]
        if contents:
            data = _contents(doc, page)
            for candidate in sorted(contents, key=lambda item: item.start, reverse=True):
                data = data[:candidate.start] + candidate.replacement + data[candidate.end:]
            # Always allocate a private stream, including for shared /Contents.
            xref = doc.get_new_xref()
            doc.update_object(xref, '<<>>')
            doc.update_stream(xref, data)
            page.set_contents(xref)
        removed = {xref for candidate in selected for xref in candidate.annot_xrefs}
        if removed:
            refs = _annotation_refs(doc, page)
            # Change this page's array, never the shared annotation object.
            doc.xref_set_key(page.xref, 'Annots', '[' + ' '.join(
                f'{xref} {generation} R' for xref, generation in refs if xref not in removed) + ']')


def _pixmap(page, size=1000):
    zoom = min(2.0, size / max(page.rect.width, page.rect.height))
    return page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=fitz.csRGB, alpha=False)


def preview(analysis, selected, page_number, cancel=None):
    if page_number not in analysis.pages:
        raise WatermarkError('预览页面不在所选页码范围内。')
    candidates = [item for item in _selection(analysis, selected) if item.page == page_number]
    _check_cancel(cancel)
    with fitz.open(stream=analysis.source, filetype='pdf') as doc:
        before = _pixmap(doc[page_number]).tobytes('png')
        _apply(doc, candidates, cancel)
        after = _pixmap(doc[page_number]).tobytes('png')
    return before, after


def export(analysis, selected, destination, cancel=None):
    candidates = _selection(analysis, selected)
    if not candidates:
        raise WatermarkError('请先选择需要去除的水印对象。')
    if os.name != 'nt':
        raise WatermarkError('首版安全另存目前仅支持 Windows。')
    source = check_unlinked_path(analysis.path)
    target = check_unlinked_path(destination)
    if target.suffix.lower() != '.pdf':
        raise WatermarkError('输出文件扩展名必须为 .pdf。')
    if source.resolve() == target.resolve() or target.exists():
        raise WatermarkError('请使用新的输出文件名；不能覆盖原件或已有文件。')
    if not target.parent.is_dir():
        raise WatermarkError('输出目录不存在。')
    if hashlib.sha256(source.read_bytes()).hexdigest() != analysis.digest:
        raise WatermarkError('原文件已改变，请重新分析和预览。')
    stage = target.parent / ('.ratools-watermark-' + uuid.uuid4().hex + '.pdf')
    stage_owned = False
    try:
        with fitz.open(stream=analysis.source, filetype='pdf') as doc:
            _apply(doc, candidates, cancel)
            _check_cancel(cancel)
            # The staging file is exclusively created; no existing file is replaced.
            with stage.open('xb') as stream:
                stage_owned = True
                stream.write(doc.tobytes(garbage=3, deflate=True))
            with fitz.open(stage) as reopened:
                if reopened.page_count != doc.page_count or reopened.is_repaired:
                    raise WatermarkError('输出结构复检失败。')
                for number in range(doc.page_count):
                    _check_cancel(cancel)
                    if reopened[number].rect != doc[number].rect:
                        raise WatermarkError('输出页面尺寸发生意外变化。')
                for number in sorted({item.page for item in candidates}):
                    _check_cancel(cancel)
                    if _pixmap(reopened[number], 700).samples != _pixmap(doc[number], 700).samples:
                        raise WatermarkError('保存后的页面与处理预览不一致。')
        _check_cancel(cancel)
        check_unlinked_path(target)
        check_unlinked_path(source)
        if hashlib.sha256(source.read_bytes()).hexdigest() != analysis.digest:
            raise WatermarkError('原文件已改变，请重新分析和预览。')
        _check_cancel(cancel)
        # Windows rename refuses an existing destination, including a race with
        # another writer. The original and all pre-existing outputs stay intact.
        os.rename(stage, target)
    finally:
        if stage_owned and stage.exists():
            try:
                recycle_owned_path(stage, target.parent)
            except Exception as error:
                raise WatermarkError(f'输出未发布，临时文件无法移入回收站，已保留在：{stage}') from error
    return str(target)
