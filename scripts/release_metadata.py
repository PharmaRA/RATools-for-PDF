"""Windows CPython 3.11 release lock, provenance and actual bundle evidence."""

import ast
import hashlib
import importlib
import importlib.metadata as metadata
import json
import marshal
import os
from pathlib import Path
import platform
import re
import struct
import subprocess
import sys
import types
import zipfile

from packaging.requirements import Requirement
from packaging.tags import sys_tags
from packaging.utils import canonicalize_name, parse_wheel_filename
from ratools_pdf.services.file_safety import check_unlinked_path


def sha256(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024*1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def load_lock(path):
    lock = json.loads(Path(path).read_text(encoding='utf-8'))
    if lock.get('schema_version') != '1.0' or lock.get('target') != 'CPython-3.11-Windows-AMD64':
        raise ValueError('Unsupported release lock target/schema')
    if sys.version_info[:2] != (3, 11) or sys.platform != 'win32' or platform.machine().lower() not in ('amd64', 'x86_64'):
        raise ValueError('Release lock requires CPython 3.11 Windows AMD64')
    names = set()
    for package in lock['packages']:
        name = canonicalize_name(package['name'])
        if name in names or not re.fullmatch('[0-9a-f]{64}', package['sha256']):
            raise ValueError('Duplicate package or invalid wheel digest')
        names.add(name)
        if not package['url'].startswith('https://files.pythonhosted.org/packages/') or not package['url'].endswith('/'+package['wheel']):
            raise ValueError('Wheel must use its explicit official PyPI HTTPS URL')
        wheel_name, version, _, tags = parse_wheel_filename(package['wheel'])
        if wheel_name != name or str(version) != package['version'] or not tags.intersection(sys_tags()):
            raise ValueError('Incompatible wheel/version in release lock')
    validate_dependency_closure(lock)
    return lock


def validate_dependency_closure(lock):
    packages = {canonicalize_name(item['name']): item for item in lock['packages']}
    for package in packages.values():
        for text in package.get('requires_dist', []):
            requirement = Requirement(text)
            if requirement.marker and not requirement.marker.evaluate({'extra': ''}):
                continue
            dependency = packages.get(canonicalize_name(requirement.name))
            if dependency is None or dependency['version'] not in requirement.specifier:
                raise ValueError(f"{package['name']} requires unlocked/incompatible {requirement.name}")


def pip_requirements(lock):
    validate_dependency_closure(lock)
    return ''.join(f"{item['name']} @ {item['url']} --hash=sha256:{item['sha256']}\n"
                   for item in lock['packages'])


def verify_environment(lock, dependency_path=None, verify_records=False):
    from base64 import urlsafe_b64encode
    result = {}
    for item in lock['packages']:
        try:
            dist = metadata.distribution(item['name'])
        except metadata.PackageNotFoundError:
            raise ValueError(f"Missing locked distribution: {item['name']}") from None
        if dist.version != item['version']:
            raise ValueError(f"Locked version mismatch: {item['name']}")
        if dependency_path and not Path(dist.locate_file('')).resolve().is_relative_to(Path(dependency_path).resolve()):
            raise ValueError(f"Distribution escaped isolated target: {item['name']}")
        actual = dict(item, requires_dist=dist.requires or [])
        validate_dependency_closure(dict(lock, packages=[actual if p['name'] == item['name'] else p for p in lock['packages']]))
        if verify_records:
            records = dist.files
            if not records or not any(str(entry).endswith('.dist-info/RECORD') for entry in records):
                raise ValueError(f"Missing installed RECORD: {item['name']}")
            verified = 0
            for entry in records:
                if not entry.hash:
                    if str(entry).endswith('.dist-info/RECORD') or str(entry).endswith('.pyc'):
                        continue
                    raise ValueError(f"Unhashed installed component: {item['name']}")
                if entry.hash.mode != 'sha256':
                    raise ValueError(f"Unsupported installed digest: {item['name']}")
                path = Path(dist.locate_file(entry))
                # pip --target relocates scripts while RECORD retains its
                # installation scheme. Verify relocated bytes rather than skip.
                if not path.is_file() and dependency_path and str(entry).replace('\\','/').startswith('../../bin/'):
                    path = Path(dependency_path)/'bin'/Path(str(entry)).name
                if not path.is_file():
                    raise ValueError(f"Missing installed component: {item['name']}")
                digest = urlsafe_b64encode(bytes.fromhex(sha256(path))).decode().rstrip('=')
                if digest != entry.hash.value:
                    raise ValueError(f"Installed component hash mismatch: {item['name']}")
                verified += 1
            if not verified:
                raise ValueError(f"No installed component hashes: {item['name']}")
        result[item['name']] = {'version': dist.version, 'role': item['role']}
    return result


def verify_import_origins(dependency_path):
    root = Path(dependency_path).resolve()
    origins = {}
    for name, module in tuple(sys.modules.items()):
        if name.split('.')[0] not in ('packaging','PyInstaller','PySide6','shiboken'+'6','pymupdf','fitz','PIL','pywin','win32com','win32api','win32ui','pythoncom','pywintypes'):
            continue
        file = getattr(module,'__file__',None)
        if file:
            # shiboken embeds these support modules and gives them virtual
            # filenames. Bind them to the verified native container instead.
            spec = getattr(module,'__spec__',None)
            if (getattr(spec,'origin',None) is None and type(getattr(spec,'loader',None)).__name__ == 'EmbeddableZipImporter'
                    and str(file).replace('\\','/').startswith('shibokensupport/')
                    and not Path(file).is_file()):
                container = sys.modules.get('shiboken6.Shiboken')
                file = getattr(container,'__file__',None)
                if not file:
                    raise ValueError('Embedded PySide6 support has no verified container')
            if not Path(file).resolve().is_relative_to(root):
                raise ValueError('Imported module escaped locked target: '+name)
            origins[name] = sha256(file)
    return origins


def _constant_structure(value):
    kind = type(value)
    if kind is types.CodeType:
        return ('code',_code_structure(value))
    if kind is tuple:
        return ('tuple',tuple(_constant_structure(item) for item in value))
    if kind is frozenset:
        return ('frozenset',frozenset(_constant_structure(item) for item in value))
    if kind is float:
        return ('float',struct.pack('!d',value))
    if kind is complex:
        return ('complex',struct.pack('!dd',value.real,value.imag))
    return (kind.__name__,value)


def _code_structure(code):
    constants = tuple(_constant_structure(value) for value in code.co_consts)
    fields = ('co_argcount','co_posonlyargcount','co_kwonlyargcount','co_nlocals','co_stacksize',
              'co_flags','co_code','co_names','co_varnames','co_filename','co_name','co_qualname',
              'co_firstlineno','co_linetable','co_exceptiontable','co_freevars','co_cellvars')
    return tuple(getattr(code,name) for name in fields)+(constants,)


def generated_stdlib_archive(path, build_work):
    path, build_work = Path(path).resolve(), Path(build_work).resolve()
    if path.name != 'base_library.zip' or not path.is_relative_to(build_work):
        raise ValueError('Unowned generated stdlib archive')
    sources, pyc_hashes = {}, {}
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        if not names or len(names) != len(set(names)):
            raise ValueError('Empty/duplicate generated stdlib archive')
        for name in names:
            relative = Path(name)
            if relative.is_absolute() or '..' in relative.parts or relative.suffix != '.pyc':
                raise ValueError('Invalid generated stdlib member')
            source = Path(sys.base_prefix)/'Lib'/relative.with_suffix('.py')
            if not source.is_file() or 'site-packages' in source.parts:
                raise ValueError('Generated member lacks stdlib source')
            data = archive.read(name)
            if data[:4] != importlib.util.MAGIC_NUMBER:
                raise ValueError('Generated stdlib bytecode version mismatch')
            code = marshal.loads(data[16:])
            allowed_names = {str(source).replace('\\','/'),relative.with_suffix('.py').as_posix()}
            if not isinstance(code,types.CodeType) or code.co_filename.replace('\\','/') not in allowed_names:
                raise ValueError('Generated stdlib source filename mismatch')
            expected = compile(source.read_bytes(),code.co_filename,'exec',dont_inherit=True,optimize=0)
            if _code_structure(expected) != _code_structure(code):
                raise ValueError('Generated stdlib bytecode differs from source')
            sources[relative.with_suffix('.py').as_posix()] = sha256(source)
            pyc_hashes[name] = hashlib.sha256(data).hexdigest()
    return {'verified_members':len(sources),'stdlib_sources':sources,'pyc_sha256':pyc_hashes}


def classify_bundle_inputs(inputs, owners, sources, app_root, build_work=None):
    app_root = Path(app_root).resolve()
    python_root = Path(sys.base_prefix).resolve()
    os_root = Path(os.environ.get('SystemRoot','C:/Windows')).resolve()
    inventory = []
    for name, source, kind in inputs:
        source = Path(source).resolve()
        owner = owners.get(source)
        generated = None
        if owner is None and build_work is not None and source.name == 'base_library.zip' and source.is_relative_to(Path(build_work).resolve()):
            generated = generated_stdlib_archive(source,build_work)
            owner = 'CPython generated stdlib archive'
        if owner is None and source.is_relative_to(app_root):
            relative = source.relative_to(app_root).as_posix()
            if relative in sources and sources[relative] == sha256(source):
                owner = 'application'
        if owner is None and source.is_relative_to(python_root):
            relative = source.relative_to(python_root)
            if 'site-packages' not in relative.parts and (relative.parts[0] in ('Lib','DLLs') or (source.parent == python_root and source.name.lower().startswith('python'))):
                owner = 'CPython'
        if owner is None and source.parent in (os_root/'System32',os_root/'SysWOW64'):
            owner = 'Windows system runtime'
        if owner is None:
            raise ValueError('Unowned Analysis input: '+name)
        entry = {'name':name,'kind':kind,'component':owner,'source_sha256':sha256(source)}
        if generated is not None:
            entry['generated_from'] = generated
        inventory.append(entry)
    return inventory


def runtime_versions():
    result = {}
    for distribution, module, attribute in [('PySide6','PySide6','__version__'),
            ('PyMuPDF','pymupdf','__version__'), ('Pillow','PIL','__version__'), ('pywin32','win32api',None)]:
        imported = importlib.import_module(module)
        actual_version = getattr(imported, attribute) if attribute else None
        distribution_version = metadata.version(distribution)
        if actual_version is not None and actual_version != distribution_version:
            raise ValueError(f'Imported runtime version mismatch: {distribution}')
        result[distribution] = {'distribution_version': distribution_version,
                                'imported_version': actual_version, 'module': module,
                                'imported_file_sha256':sha256(imported.__file__)}
    return result


def source_provenance(root):
    root = Path(root)
    flags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
    def git(*args):
        return subprocess.check_output(['git', '-C', str(root), *args], creationflags=flags).decode('utf-8').rstrip('\r\n')
    inputs = ['main.py','main_no_update.py','ratools_pdf','scripts/build_pyinstaller_safely.py',
              'scripts/build_cleanup_guard.py','scripts/release_metadata.py','scripts/sign_release.py',
              'scripts/lock_release_dependencies.py','patch_pe_subsystem.py','requirements.txt',
              'requirements-release-win311.lock.json','LICENSE','THIRD_PARTY_NOTICES.md','icon.ico','plugins']
    sources = {}
    for name in inputs:
        path = root/name
        files = sorted(path.rglob('*')) if path.is_dir() else [path]
        for file in files:
            if file.is_file() and '__pycache__' not in file.parts and (name != 'ratools_pdf' or file.suffix == '.py'):
                sources[file.relative_to(root).as_posix()] = sha256(file)
    status = git('status','--porcelain','--untracked-files=normal','--',*inputs)
    return {'commit': git('rev-parse','HEAD'), 'dirty': bool(status),
            'dirty_paths': [line[3:].strip('"') for line in status.splitlines()],
            'sources': sources, 'python': platform.python_version(), 'platform': platform.platform(),
            'reproducibility': 'Pinned wheels; local interpreter/OS/toolchain, not a hermetic or bit-reproducible build'}


def write_hash_manifest(root):
    root = Path(root)
    entries = {path.relative_to(root).as_posix(): sha256(path) for path in sorted(root.rglob('*'))
               if path.is_file() and path != root/'artifact-hashes.json'}
    (root/'artifact-hashes.json').write_text(json.dumps({'algorithm':'SHA-256','files':entries}, indent=2), encoding='utf-8')
    return entries


def verify_hash_manifest(root):
    root = Path(root)
    expected = json.loads((root/'artifact-hashes.json').read_text(encoding='utf-8'))['files']
    actual = {path.relative_to(root).as_posix(): sha256(path) for path in root.rglob('*')
              if path.is_file() and path != root/'artifact-hashes.json'}
    return actual == expected


def signing_plan(certificate_sha1=None, certificate_env=None):
    if certificate_sha1 and not re.fullmatch('[0-9a-fA-F]{40}', certificate_sha1):
        raise ValueError('Certificate store SHA1 selector must be 40 hexadecimal characters')
    if certificate_env and not re.fullmatch('[A-Za-z_][A-Za-z0-9_]*', certificate_env):
        raise ValueError('Invalid certificate environment variable name')
    if certificate_sha1 and certificate_env:
        raise ValueError('Choose one certificate reference')
    return {'status': 'pending' if certificate_sha1 or certificate_env else 'unsigned',
            'certificate_store_sha1': certificate_sha1, 'certificate_environment': certificate_env,
            'timestamp': 'not_requested', 'signature_verified': False}


def signing_preflight(executable, checker=check_unlinked_path):
    executable = checker(executable).resolve()
    bundle = executable.parent
    for path in (bundle/'release/build-provenance.json',bundle/'release/signing.json',
                 bundle/'artifact-hashes.json',bundle.parent/'build.json'):
        checker(path)
    if not executable.is_file() or not (bundle/'release/build-provenance.json').is_file() or not (bundle/'artifact-hashes.json').is_file():
        raise ValueError('Signing requires a complete existing release bundle')
    return executable


def record_signing_result(executable, plan, success):
    executable = signing_preflight(executable)
    bundle = executable.parent
    plan = dict(plan,status='signed' if success else 'failed',signature_verified=bool(success))
    provenance = bundle/'release/build-provenance.json'
    document = json.loads(provenance.read_text(encoding='utf-8'))
    document['signing'] = plan
    provenance.write_text(json.dumps(document,indent=2),encoding='utf-8')
    (bundle/'release/signing.json').write_text(json.dumps(plan,indent=2),encoding='utf-8')
    summary = bundle.parent/'build.json'
    if summary.is_file():
        document = json.loads(summary.read_text(encoding='utf-8'))
        states = document.setdefault('signing_by_executable',{})
        states[str(executable)] = plan
        statuses = [states.get(str(Path(path).resolve()),{'status':'unsigned'})['status'] for path in document['executables']]
        status = statuses[0] if len(set(statuses)) == 1 else 'mixed'
        document['signing'] = {'status':status,'signature_verified':all(value == 'signed' for value in statuses)}
        summary.write_text(json.dumps(document,indent=2),encoding='utf-8')
    write_hash_manifest(bundle)
    return plan


def bundle_evidence(bundle, analysis_toc, lock, provenance, variant):
    """Associate actual Analysis inputs to installed distributions, never all machine packages."""
    bundle = Path(bundle)
    inputs = []
    def visit(value):
        if isinstance(value, (tuple, list)):
            if len(value) == 3 and all(isinstance(v, str) for v in value):
                name, source, kind = value
                if Path(source).is_file():
                    inputs.append((name, Path(source).resolve(), kind))
            else:
                for child in value:
                    visit(child)
    visit(ast.literal_eval(Path(analysis_toc).read_text(encoding='utf-8')))
    components = []
    owners = {}
    for item in lock['packages']:
        dist = metadata.distribution(item['name'])
        files = {Path(dist.locate_file(entry)).resolve() for entry in dist.files or []}
        owners.update({path:item['name'] for path in files})
        bundled = [{'name': name, 'kind':kind, 'source_sha256':sha256(source)}
                   for name, source, kind in inputs if source in files]
        if not bundled:
            continue
        licenses = []
        for entry in dist.files or []:
            path = Path(str(entry))
            if any(part.lower() in ('licenses','license','licences') for part in path.parts) or path.name.lower().startswith(('license','copying','notice','copyright')):
                source = Path(dist.locate_file(entry))
                if source.is_file():
                    target = bundle/'release/licenses'/canonicalize_name(item['name'])/str(entry)
                    target.parent.mkdir(parents=True, exist_ok=True)
                    target.write_bytes(source.read_bytes())
                    licenses.append(target.relative_to(bundle).as_posix())
        components.append({'name':item['name'], 'version':dist.version, 'wheel_sha256':item['sha256'],
                           'license_expression':dist.metadata.get('License-Expression'),
                           'license_metadata':dist.metadata.get('License'),
                           'license_files': licenses, 'bundled_inputs':bundled})
    from ratools_pdf.pdf.qpdf import run_qpdf_command
    qpdf = run_qpdf_command(['--version'],timeout=15)
    if qpdf.returncode:
        raise RuntimeError('Bundled qpdf version check failed')
    plugin_files = [path for path in bundle.rglob('*') if path.is_file() and 'qpdf' in path.parts]
    components.append({'name':'qpdf', 'version_output':qpdf.stdout.splitlines()[0],
                       'bundled_files':[{ 'path':p.relative_to(bundle).as_posix(),'sha256':sha256(p)} for p in plugin_files],
                       'license_files':[p.relative_to(bundle).as_posix() for p in plugin_files
                           if any(word in p.name.lower() for word in ('license','notice','copying'))]})
    python_license = Path(sys.base_prefix)/'LICENSE.txt'
    python_files = [path for path in bundle.rglob('python*.dll')]
    if python_license.is_file():
        target = bundle/'release/licenses/python/LICENSE.txt'
        target.parent.mkdir(parents=True,exist_ok=True)
        target.write_bytes(python_license.read_bytes())
    components.append({'name':'CPython', 'version':platform.python_version(),
                       'bundled_files':[{'path':p.relative_to(bundle).as_posix(),'sha256':sha256(p)} for p in python_files],
                       'license_files':['release/licenses/python/LICENSE.txt'] if python_license.is_file() else []})
    release = bundle/'release'
    release.mkdir(exist_ok=True)
    inventory = classify_bundle_inputs(inputs,owners,provenance['sources'],Path(__file__).resolve().parents[1],Path(analysis_toc).parents[2])
    (release/'sbom.json').write_text(json.dumps({'schema_version':'1.0','basis':'PyInstaller actual Analysis inputs and bundled files',
        'components':components,'analysis_inputs':inventory,
        'limitations':['License collection is inventory evidence; legal/license compliance review is pending.']}, indent=2),encoding='utf-8')
    (release/'build-provenance.json').write_text(json.dumps({'variant':variant,'source':provenance,
        'runtime':runtime_versions(),'locked_environment':verify_environment(lock), 'signing':signing_plan()},indent=2),encoding='utf-8')
    (release/'requirements-release-win311.lock.json').write_text(json.dumps(lock,indent=2),encoding='utf-8')
    write_hash_manifest(bundle)
