"""Refresh a Windows/Python 3.11 release lock from official wheel metadata.

Runtime versions come only from the four declared application distributions.
Build tool versions are explicit inputs; transitive closure is resolved via PyPI.
Review the newly generated lock before building. No installation is performed.
"""

import argparse
import importlib.metadata as metadata
import json
from pathlib import Path
import sys
import urllib.request

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from packaging.requirements import Requirement
from packaging.tags import sys_tags
from packaging.utils import canonicalize_name, parse_wheel_filename
from scripts.release_metadata import validate_dependency_closure
from ratools_pdf.services.file_safety import check_unlinked_path


def generate(output, pyinstaller_version, proxy=None):
    output = check_unlinked_path(output)
    if output.exists():
        raise FileExistsError('Lock refresh requires a new output file; review before replacing the active lock')
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({'https':proxy,'http':proxy} if proxy else {}))
    pending = [(name, metadata.version(name), 'runtime') for name in ('PySide6','PyMuPDF','Pillow','pywin32')]
    pending += [('PyInstaller',pyinstaller_version,'build')]
    packages = {}
    tags = list(sys_tags())
    while pending:
        name, version, role = pending.pop(0)
        key = canonicalize_name(name)
        if key in packages:
            continue
        endpoint = f'https://pypi.org/pypi/{name}/{version}/json' if version else f'https://pypi.org/pypi/{name}/json'
        with opener.open(endpoint,timeout=45) as response:
            data = json.load(response)
        candidates = []
        for wheel in data['urls']:
            if wheel['packagetype'] != 'bdist_wheel' or wheel.get('yanked'):
                continue
            wheel_tags = parse_wheel_filename(wheel['filename'])[3]
            rank = next((i for i, tag in enumerate(tags) if tag in wheel_tags), None)
            if rank is not None:
                candidates.append((rank,wheel))
        if not candidates:
            raise ValueError(f'No compatible wheel: {name}')
        wheel = sorted(candidates,key=lambda entry:entry[0])[0][1]
        package = {'name':name, 'version':data['info']['version'], 'role':role,
                   'wheel':wheel['filename'], 'url':wheel['url'], 'sha256':wheel['digests']['sha256'],
                   'requires_dist':data['info']['requires_dist'] or []}
        packages[key] = package
        for text in package['requires_dist']:
            req = Requirement(text)
            if req.marker and not req.marker.evaluate({'extra':''}):
                continue
            dep = canonicalize_name(req.name)
            if dep in packages or any(canonicalize_name(p[0]) == dep for p in pending):
                continue
            # Preserve an installed compatible version when available; otherwise
            # fetch the current official version and validate the complete closure.
            try:
                selected = metadata.version(req.name)
                if selected not in req.specifier:
                    selected = None
            except metadata.PackageNotFoundError:
                selected = None
            pending.append((req.name,selected,role))
    lock = {'schema_version':'1.0','target':'CPython-3.11-Windows-AMD64',
            'markers':'Current CPython 3.11 Windows AMD64, extra empty; no optional extras',
            'packages':sorted(packages.values(),key=lambda item:canonicalize_name(item['name']))}
    validate_dependency_closure(lock)
    output.write_text(json.dumps(lock,indent=2)+'\n',encoding='utf-8')
    print('LOCK_WRITTEN packages='+str(len(packages)))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('output',type=Path)
    parser.add_argument('--pyinstaller-version',required=True)
    parser.add_argument('--proxy')
    args = parser.parse_args()
    try:
        generate(args.output,args.pyinstaller_version,args.proxy)
    except Exception as exc:
        raise SystemExit('Lock generation failed: '+type(exc).__name__) from None
