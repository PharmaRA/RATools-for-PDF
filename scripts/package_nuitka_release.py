"""Stage the PyMuPDF runtime and produce verified Windows-compatible ZIPs."""

import argparse
from copy import copy
import hashlib
import json
import os
from pathlib import Path, PurePosixPath, PureWindowsPath
import shutil
import stat
import sys
import uuid
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from ratools_pdf.services.file_safety import check_unlinked_path, recycle_owned_path


def exclusion_reason(member):
    """Exclude only optional layout weights and PyMuPDF bytecode caches."""
    parts = PurePosixPath(member).parts
    if parts and parts[0] == 'pymupdf':
        relative = parts[1:]
    elif len(parts) > 1 and parts[1] == 'pymupdf':
        relative = parts[2:]
    else:
        return None
    if '__pycache__' in relative:
        return 'pymupdf-bytecode-cache'
    if (relative[:3] == ('layout', 'resources', 'onnx')
            and PurePosixPath(member).suffix.lower() == '.onnx'):
        return 'optional-layout-model'
    return None


def recycle_target(path, owned_root):
    target = check_unlinked_path(path)
    root = check_unlinked_path(owned_root).resolve()
    resolved = target.resolve()
    if resolved == root or not resolved.is_relative_to(root):
        raise ValueError('Recycle target must be a child of the owned directory')
    recycle_owned_path(target, root)


def _folder_inventory(source):
    source = check_unlinked_path(source).resolve()
    if not source.is_dir():
        raise ValueError('Package source must be a directory')
    kept = [(source, source.name + '/')]
    excluded = []
    for current, directories, files in os.walk(source, followlinks=False):
        current = check_unlinked_path(current)
        retained_directories = []
        for name in sorted(directories):
            path = check_unlinked_path(current / name)
            member = source.name + '/' + path.relative_to(source).as_posix() + '/'
            reason = exclusion_reason(member)
            if reason:
                excluded.append({'member': member, 'reason': reason, 'bytes': 0})
            else:
                retained_directories.append(name)
                kept.append((path, member))
        directories[:] = retained_directories
        for name in sorted(files):
            path = check_unlinked_path(current / name)
            if not path.is_file():
                raise ValueError('Package contains a non-regular file')
            member = source.name + '/' + path.relative_to(source).as_posix()
            reason = exclusion_reason(member)
            if reason:
                excluded.append({'member': member, 'reason': reason,
                                 'bytes': path.stat().st_size})
            else:
                kept.append((path, member))
    return source, sorted(kept, key=lambda item: item[1]), excluded


def stage_pymupdf(source, destination, owned_root):
    source = check_unlinked_path(source).resolve()
    destination = check_unlinked_path(destination)
    root = check_unlinked_path(owned_root).resolve()
    if source.name != 'pymupdf' or not (source / '__init__.py').is_file():
        raise ValueError('Expected the installed pymupdf package')
    if destination.name != 'pymupdf':
        raise ValueError('Runtime destination must be named pymupdf')
    if destination.resolve() == root or not destination.resolve().is_relative_to(root):
        raise ValueError('Runtime destination must be inside the build directory')
    if (destination.resolve().is_relative_to(source)
            or source.is_relative_to(destination.resolve())):
        raise ValueError('Runtime source and destination must not overlap')
    source, kept, excluded = _folder_inventory(source)
    # Complete the source/link preflight before changing an existing destination.
    if destination.exists():
        recycle_target(destination, root)
    destination.mkdir(parents=True, exist_ok=False)
    file_count = payload_bytes = 0
    for path, _member in kept:
        target = destination / path.relative_to(source)
        check_unlinked_path(target)
        if path.is_dir():
            target.mkdir(parents=True, exist_ok=True)
        else:
            shutil.copy2(path, target)
            file_count += 1
            payload_bytes += path.stat().st_size
    return {'status': 'passed', 'operation': 'stage-pymupdf',
            'files': file_count, 'payload_bytes': payload_bytes, 'excluded': excluded}


def _safe_archive_member(info):
    member = info.orig_filename.replace('\\', '/')
    if ('\x00' in member or member.startswith('/') or PureWindowsPath(member).drive
            or '..' in member.split('/')):
        raise ValueError('Unsafe ZIP member path')
    mode = info.external_attr >> 16
    if stat.S_IFMT(mode) not in (0, stat.S_IFREG, stat.S_IFDIR):
        raise ValueError('ZIP contains a link or non-regular member')
    if info.flag_bits & 1:
        raise ValueError('Encrypted ZIP members are unsupported')
    normalized = PurePosixPath(member).as_posix()
    if normalized == '.':
        raise ValueError('Empty ZIP member path')
    return normalized + '/' if member.endswith('/') else normalized


def _new_output(output, source):
    output = check_unlinked_path(output)
    source = check_unlinked_path(source).resolve()
    if output.exists():
        raise FileExistsError('ZIP output already exists')
    if output.resolve() == source or (source.is_dir() and output.resolve().is_relative_to(source)):
        raise ValueError('ZIP output must be outside its source')
    output.parent.mkdir(parents=True, exist_ok=True)
    return output


def _write_verified_zip(output, records, excluded, operation, source_bytes=None):
    temporary = check_unlinked_path(
        output.with_name(output.name + '.' + uuid.uuid4().hex + '.partial'))
    expected = {}
    with temporary.open('xb') as stream:
        with zipfile.ZipFile(stream, 'w', compression=zipfile.ZIP_DEFLATED,
                             compresslevel=9) as archive:
            for info, read_payload in records:
                payload = b'' if info.is_dir() else read_payload()
                compression = zipfile.ZIP_STORED if info.is_dir() else zipfile.ZIP_DEFLATED
                archive.writestr(info, payload, compress_type=compression, compresslevel=9)
                expected[info.filename] = hashlib.sha256(payload).hexdigest()
    with zipfile.ZipFile(temporary) as archive:
        if archive.namelist() != list(expected):
            raise RuntimeError('ZIP member inventory changed')
        for info in archive.infolist():
            if hashlib.sha256(archive.read(info)).hexdigest() != expected[info.filename]:
                raise RuntimeError('ZIP payload verification failed')
    check_unlinked_path(output)
    if output.exists():
        raise FileExistsError('ZIP output appeared during packaging')
    # Windows rename does not overwrite a concurrently created output. A failed
    # operation leaves its uniquely named partial archive available for diagnosis.
    temporary.rename(output)
    files = sum(not info.is_dir() for info, _read in records)
    with output.open('rb') as stream:
        archive_hash = hashlib.file_digest(stream, 'sha256').hexdigest()
    return {'status': 'passed', 'operation': operation, 'compression': 'deflate-9',
            'files': files, 'zip_bytes': output.stat().st_size,
            'source_zip_bytes': source_bytes,
            'sha256': archive_hash,
            'retained_payloads_verified': files, 'excluded': excluded}


def create_zip(source, output):
    output = _new_output(output, source)
    _source, kept, excluded = _folder_inventory(source)
    records = []
    for path, member in kept:
        info = zipfile.ZipInfo.from_file(path, member)
        records.append((info, lambda path=path: path.read_bytes()))
    return _write_verified_zip(output, records, excluded, 'zip')


def repack_zip(source, output):
    source = check_unlinked_path(source)
    output = _new_output(output, source)
    with zipfile.ZipFile(source) as original:
        records, excluded, seen = [], [], set()
        for info in original.infolist():
            member = _safe_archive_member(info)
            if member.casefold() in seen:
                raise ValueError('Duplicate ZIP member path')
            seen.add(member.casefold())
            reason = exclusion_reason(member)
            if reason:
                excluded.append({'member': member, 'reason': reason,
                                 'bytes': info.file_size, 'compressed_bytes': info.compress_size})
                continue
            target_info = copy(info)
            target_info.filename = member
            records.append((target_info, lambda info=info: original.read(info)))
        records.sort(key=lambda item: item[0].filename)
        return _write_verified_zip(output, records, excluded, 'repack', source.stat().st_size)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    stage = commands.add_parser('stage-pymupdf')
    stage.add_argument('--source', type=Path, required=True)
    stage.add_argument('--destination', type=Path, required=True)
    stage.add_argument('--owned-root', type=Path, required=True)
    recycle = commands.add_parser('recycle')
    recycle.add_argument('--path', type=Path, required=True)
    recycle.add_argument('--owned-root', type=Path, required=True)
    for command in ('zip', 'repack'):
        package = commands.add_parser(command)
        package.add_argument('--source', type=Path, required=True)
        package.add_argument('--output', type=Path, required=True)
        package.add_argument('--report', type=Path)
    args = parser.parse_args(argv)
    try:
        report_path = getattr(args, 'report', None)
        if report_path is not None:
            report_path = check_unlinked_path(report_path)
            if report_path.exists() or report_path.resolve() == args.output.resolve():
                raise FileExistsError('Packaging report must use a new separate path')
        if args.command == 'stage-pymupdf':
            result = stage_pymupdf(args.source, args.destination, args.owned_root)
        elif args.command == 'recycle':
            recycle_target(args.path, args.owned_root)
            result = {'status': 'passed', 'operation': 'recycle'}
        else:
            operation = create_zip if args.command == 'zip' else repack_zip
            result = operation(args.source, args.output)
        if report_path is not None:
            report_path.parent.mkdir(parents=True, exist_ok=True)
            with report_path.open('x', encoding='utf-8') as stream:
                json.dump(result, stream, ensure_ascii=False, indent=2)
        summary = {key: value for key, value in result.items() if key != 'excluded'}
        summary['excluded_entries'] = len(result.get('excluded', []))
        print(json.dumps(summary, ensure_ascii=False))
        return 0
    except Exception as exc:
        print('PACKAGING_FAILED ' + type(exc).__name__, file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
