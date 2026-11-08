"""Process-local recycle adapters inherited by Python build subprocesses."""

import os
from pathlib import Path
import shutil
import sys
import tempfile

from ratools_pdf.services.file_safety import check_unlinked_path, recycle_owned_path


def install_from_environment():
    root = Path(__file__).resolve().parents[1]
    work = check_unlinked_path(os.environ['RATOOLS_BUILD_WORK']).resolve()
    dist = check_unlinked_path(os.environ['RATOOLS_BUILD_DIST']).resolve()
    if (work.parent != (root/'build_pyinstaller').resolve()
            or dist.parent != (root/'dist').resolve()
            or work.name != dist.name
            or len(work.name) != 32
            or any(char not in '0123456789abcdef' for char in work.name)
            or not work.is_dir() or not dist.is_dir()):
        raise RuntimeError('Invalid owned build directories')

    def recycle(path, *_args, **_kwargs):
        candidate = check_unlinked_path(os.fsdecode(path))
        for owner in (work, dist):
            if candidate.resolve().is_relative_to(owner):
                if candidate.exists():
                    recycle_owned_path(candidate, owner)
                return
        raise RuntimeError('Build cleanup attempted outside the current owned directories')

    def audit(event, _args):
        if event in ('os.remove', 'os.rmdir', 'shutil.rmtree'):
            raise RuntimeError('Physical deletion is disabled during this build')

    sys.addaudithook(audit)
    for name in ('remove', 'unlink', 'rmdir'):
        setattr(os, name, recycle)
    shutil.rmtree = recycle
    tempfile.tempdir = str(work/'tmp')
    sys.dont_write_bytecode = True


def deletion_is_blocked():
    """No file operation: exercise the audit hook with a synthetic event."""
    import sys
    try:
        sys.audit('os.remove', 'SYNTHETIC_NO_FILE', -1)
    except RuntimeError:
        return True
    return False
