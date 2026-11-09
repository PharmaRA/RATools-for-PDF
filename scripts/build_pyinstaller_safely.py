"""Build isolated release folders; all Python cleanup goes to the recycle bin.

python -B scripts/build_pyinstaller_safely.py [--install-dependency] [--variant normal|no-update|both]
Each invocation creates a new build/cache/dependency/output directory; older builds stay intact.
"""

import argparse
from contextlib import ExitStack, redirect_stderr, redirect_stdout
from io import StringIO
import json
import importlib
import os
from pathlib import Path
import site
import shutil
import subprocess
import sys
import tempfile
import uuid
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
sys.dont_write_bytecode = True
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'
from ratools_pdf.services.file_safety import check_unlinked_path, recycle_owned_path
from scripts.release_metadata import (bundle_evidence, load_lock, pip_requirements,
    sha256, source_provenance, verify_environment, verify_import_origins)


def run_locked_child(command):
    return subprocess.run(command,stdout=sys.stdout,stderr=sys.stderr,
                          creationflags=subprocess.CREATE_NO_WINDOW).returncode


def locked_python_paths(bootstrap,dependencies,root):
    dependencies=Path(dependencies)
    return [str(path) for path in (bootstrap,dependencies,dependencies/'win32',
        dependencies/'win32/lib',dependencies/'Pythonwin',dependencies/'pywin32_system32',root)]


def resolve_analysis_toc(directory):
    paths=list(Path(directory).glob('Analysis-*.toc'))
    if len(paths)!=1:
        raise ValueError('Expected exactly one actual Analysis output in the variant directory')
    return paths[0]


def build():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--install-dependency',action='store_true')
    parser.add_argument('--variant',choices=('normal','no-update','both'),default='both')
    parser.add_argument('--proxy', help='Optional dependency-download proxy URL')
    parser.add_argument('--dependency-path', help='Reuse an existing isolated PyInstaller dependency directory')
    parser.add_argument('--lock',type=Path,default=ROOT/'requirements-release-win311.lock.json')
    parser.add_argument('--locked-child',action='store_true',help=argparse.SUPPRESS)
    args=parser.parse_args()
    if os.name != 'nt':
        parser.error('This recycling build runner currently requires Windows')
    lock=load_lock(args.lock)
    if not args.install_dependency and not args.dependency_path:
        parser.error('Use --install-dependency or a previously locked --dependency-path')
    provenance=source_provenance(ROOT)
    if args.locked_child:
        from scripts.build_cleanup_guard import install_from_environment
        install_from_environment()
        work=check_unlinked_path(os.environ['RATOOLS_BUILD_WORK'])
        dist=check_unlinked_path(os.environ['RATOOLS_BUILD_DIST'])
        run_id=work.name
    else:
        run_id=uuid.uuid4().hex
        work=check_unlinked_path(ROOT/'build_pyinstaller'/run_id)
        dist=check_unlinked_path(ROOT/'dist'/run_id)
        work.mkdir(parents=True,exist_ok=False)
        dist.mkdir(parents=True,exist_ok=False)
    temp=work/'tmp'
    temp.mkdir(exist_ok=args.locked_child)
    tempfile.tempdir=str(temp)
    os.environ['TEMP']=os.environ['TMP']=str(temp)
    os.environ['PYINSTALLER_CONFIG_DIR']=str(work/'cache')
    os.environ['RATOOLS_BUILD_WORK']=str(work)
    os.environ['RATOOLS_BUILD_DIST']=str(dist)
    bootstrap=work/'bootstrap'
    bootstrap.mkdir(exist_ok=args.locked_child)
    # Python normally logs and ignores sitecustomize errors. Exit immediately
    # instead so a failed guard cannot leave an unprotected analysis process.
    (bootstrap/'sitecustomize.py').write_text(
        'import os\ntry:\n    from scripts.build_cleanup_guard import install_from_environment\n'
        '    install_from_environment()\nexcept BaseException:\n    os._exit(91)\n',encoding='utf-8')
    python_paths=[str(bootstrap),str(ROOT)]
    if args.dependency_path:
        dependencies=check_unlinked_path(args.dependency_path).resolve()
        if not dependencies.is_relative_to((ROOT/'build_pyinstaller').resolve()) or not dependencies.is_dir():
            raise ValueError('Dependency directory must belong to a previous isolated build')
        python_paths.append(str(dependencies))
    os.environ['PYTHONPATH']=os.pathsep.join(python_paths)
    def recycle(path,*_args,**_kwargs):
        candidate=Path(os.fsdecode(path)).absolute()
        for owner in (work,dist):
            if candidate.resolve().is_relative_to(owner.resolve()):
                if candidate.exists():
                    recycle_owned_path(candidate,owner)
                return
        raise RuntimeError('Build cleanup attempted outside the current owned directories')
    def audit(event,_args):
        if event in ('os.remove','os.rmdir','shutil.rmtree'):
            raise RuntimeError('Physical deletion is disabled during this build')
    sys.addaudithook(audit)
    with ExitStack() as stack:
        for name in ('remove','unlink','rmdir'):
            stack.enter_context(patch.object(os,name,recycle))
        stack.enter_context(patch.object(shutil,'rmtree',recycle))
        if args.install_dependency:
            from pip._internal.cli.main import main as pip_main
            dependencies=work/'dependencies'
            requirements=work/'locked-requirements.txt'
            requirements.write_text(pip_requirements(lock),encoding='utf-8')
            # Keep installer diagnostics private; dependency fetching uses the
            # explicit public index, with no user-configured credential output.
            install_args=['--isolated','install','--index-url','https://pypi.org/simple',
                          '--no-cache-dir','--no-compile','--only-binary=:all:',
                          '--no-deps','--require-hashes',
                          '--target',str(dependencies),'-r',str(requirements)]
            if args.proxy:
                install_args += ['--proxy', args.proxy]
            output, errors = StringIO(), StringIO()
            with redirect_stdout(output),redirect_stderr(errors):
                status=pip_main(install_args)
            if status:
                diagnostic = output.getvalue() + errors.getvalue()
                categories = [label for label in ('Physical deletion', 'outside the current',
                              'CERTIFICATE', 'ProxyError', 'PermissionError', 'Could not find',
                              'Connection', 'Timeout', 'FileNotFoundError', 'TypeError')
                              if label.lower() in diagnostic.lower()]
                raise RuntimeError(f'PyInstaller dependency installation failed (exit {status}; categories: {categories})')
            python_paths.append(str(dependencies))
            (dependencies/'ratools-lock-sha256.txt').write_text(sha256(args.lock),encoding='ascii')
        if (dependencies/'ratools-lock-sha256.txt').read_text(encoding='ascii') != sha256(args.lock):
            raise ValueError('Dependency target was not installed from the selected lock')
        site.addsitedir(str(dependencies))
        for folder in reversed([Path(path) for path in locked_python_paths(bootstrap,dependencies,ROOT)[1:-1]]):
            if folder.is_dir():
                sys.path.insert(0,str(folder))
                python_paths.append(str(folder))
        importlib.invalidate_caches()
        # Fresh interpreter: locked paths precede all ambient site packages
        # before packaging, PyInstaller, native runtimes or metadata are imported.
        python_paths=locked_python_paths(bootstrap,dependencies,ROOT)
        os.environ['PYTHONPATH']=os.pathsep.join(python_paths)
        if not args.locked_child:
            command=[sys.executable,'-X','utf8','-B',str(Path(__file__).resolve()),
                     '--locked-child','--dependency-path',str(dependencies),'--variant',args.variant,
                     '--lock',str(args.lock.resolve())]
            status=run_locked_child(command)
            raise SystemExit(status)
        environment=verify_environment(lock,dependencies,verify_records=True)
        print('LOCK_VERIFIED packages='+str(len(environment)),flush=True)
        import PyInstaller
        from PyInstaller.__main__ import run
        from PyInstaller import isolated
        from scripts.build_cleanup_guard import deletion_is_blocked
        origins=verify_import_origins(dependencies)
        if not deletion_is_blocked() or not isolated.call(deletion_is_blocked):
            raise RuntimeError('Build cleanup guard did not reach the analysis subprocess')
        print('BUILD_GUARD parent=blocked isolated_child=blocked',flush=True)
        from ratools_pdf.config.version import APP_COMPANY, APP_NAME, APP_VERSION_STR, APP_WINDOWS_VERSION
        from patch_pe_subsystem import patch_subsystem, WINDOWS_GUI_SUBSYSTEM
        version=work/'version.txt'
        strings={'CompanyName':APP_COMPANY,'FileDescription':APP_NAME,'FileVersion':APP_VERSION_STR,
                 'ProductName':APP_NAME,'ProductVersion':APP_VERSION_STR}
        version.write_text(f"VSVersionInfo(ffi=FixedFileInfo(filevers={APP_WINDOWS_VERSION!r},prodvers={APP_WINDOWS_VERSION!r},mask=63,flags=0,OS=0x40004,fileType=1,subtype=0,date=(0,0)),kids=[StringFileInfo([StringTable('080404B0',["+
                           ','.join(f'StringStruct({key!r},{value!r})' for key,value in strings.items())+
                           "])]),VarFileInfo([VarStruct('Translation',[2052,1200])])])",encoding='utf-8')
        variants=[('normal','main.py','RATools-for-PDF'),('no-update','main_no_update.py','RATools-for-PDF-NoUpdate')]
        outputs=[]
        for kind,entry,name in variants:
            if args.variant not in ('both',kind):
                continue
            options=[str(ROOT/entry),'--console','--onedir','--noupx','--name',name,
                     '--distpath',str(dist),'--workpath',str(work/kind),'--specpath',str(work),
                     '--icon',str(ROOT/'icon.ico'),'--version-file',str(version)]
            for module in ('torch','torchvision','easyocr','tensorflow','pandas','numpy','scipy','sitecustomize','scripts.build_cleanup_guard'):
                options += ['--exclude-module',module]
            if kind=='no-update':
                options += ['--exclude-module','ratools_pdf.services.update_checker']
            for package in lock['packages']:
                if package['role']=='runtime':
                    options += ['--copy-metadata',package['name']]
            for resource,target in (('LICENSE','.'),('THIRD_PARTY_NOTICES.md','.'),('icon.ico','.'),('plugins','plugins')):
                options += ['--add-data',f'{ROOT/resource};{target}']
            run(options)
            executable=dist/name/(name+'.exe')
            patch_subsystem(executable,WINDOWS_GUI_SUBSYSTEM)
            plugin=dist/name/'_internal/PySide6/plugins/tls/qopensslbackend.dll'
            if plugin.exists():
                recycle(plugin)
            bundle_evidence(executable.parent,resolve_analysis_toc(work/kind/name),lock,provenance,kind)
            verify_import_origins(dependencies)
            outputs.append(str(executable))
        (dist/'build.json').write_text(json.dumps({'version':APP_VERSION_STR,'pyinstaller':PyInstaller.__version__,
            'lock_sha256':sha256(args.lock),'source':provenance,'environment':environment,'initial_import_origins':origins,
            'signing':{'status':'unsigned','signature_verified':False},'executables':outputs},indent=2),encoding='utf-8')
        print('BUILD_OUTPUT='+str(dist),flush=True)


if __name__=='__main__':
    build()
