"""Plan collision-free batch destinations before any engine writes a file."""

import hashlib
import glob
import os
import re

from ratools_pdf.services.file_safety import check_unlinked_path


def common_base_dir(file_paths, fallback=''):
    try:
        dirs = [os.path.dirname(os.path.abspath(path)) for path in file_paths]
        return os.path.commonpath(dirs)
    except ValueError:
        return fallback


def relative_source_directory(file_path, common_base):
    directory = os.path.dirname(os.path.abspath(file_path))
    if common_base:
        try:
            relative = os.path.relpath(directory, common_base)
        except ValueError:
            relative = '..'
        if relative == '.':
            return ''
        if relative != '..' and not relative.startswith('..' + os.sep) and not os.path.isabs(relative):
            return relative
    # Different drives / external sources retain the whole source directory and volume.
    drive, tail = os.path.splitdrive(directory)
    if len(drive) == 2 and drive[1] == ':':
        volume = 'drive-' + drive[0].upper()
    elif drive:
        volume = 'volume-' + hashlib.sha256(os.path.normcase(drive).encode('utf-8')).hexdigest()
    else:
        volume = 'root'
    return os.path.join('_external', volume, tail.lstrip('/\\'))


def path_key(path):
    return os.path.normcase(os.path.abspath(path))


def validate_output_plan(pairs, input_paths, *, allow_in_place=False, patterns=False):
    """Reject duplicate destinations and writes over any queued input.

    pairs are (source, destination); for split output, %d represents qpdf's
    page number or page range. Existing results may be replaced, queued inputs may not.
    """
    inputs = {path_key(path) for path in input_paths}
    destinations = set()
    for source, destination in pairs:
        check_unlinked_path(source)
        check_unlinked_path(destination)
        key = path_key(destination)
        if key in destinations:
            raise ValueError(f'多个任务将写入同一输出，已停止整批处理: {destination}')
        destinations.add(key)
        if patterns:
            if key.count('%d') != 1:
                raise ValueError(f'拆分输出路径必须仅含一个页码占位符: {destination}')
            expression = re.escape(key).replace('%d', r'\d+(?:-\d+)?')
            if any(re.fullmatch(expression, item) for item in inputs):
                raise ValueError(f'拆分结果可能覆盖待处理输入，已停止整批处理: {destination}')
            # A literal %d path does not audit already existing numbered outputs.
            for existing in glob.iglob(glob.escape(destination).replace('%d', '*')):
                if re.fullmatch(expression, path_key(existing)):
                    check_unlinked_path(existing)
        elif key in inputs and not (allow_in_place and key == path_key(source)):
            raise ValueError(f'输出将覆盖待处理输入，已停止整批处理: {destination}')
