"""Explicit, isolated, synthetic source/frozen release smoke verification."""

from datetime import date
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import sys
import shutil
import tempfile

from ratools_pdf.config.features import ENABLE_UPDATE_CHECK
from ratools_pdf.services.file_safety import check_unlinked_path, recycle_owned_path


def verify_release(output_directory):
    root = check_unlinked_path(output_directory)
    root.mkdir(parents=True,exist_ok=False)
    temporary = root/'temporary'
    temporary.mkdir()
    tempfile.tempdir = str(temporary)
    def recycle(path, *_args, **_kwargs):
        candidate = check_unlinked_path(os.fsdecode(path))
        if not candidate.resolve().is_relative_to(root.resolve()):
            raise RuntimeError('Release cleanup attempted outside its output directory')
        if candidate.exists():
            recycle_owned_path(candidate,root)
    for operation in ('remove','unlink','rmdir'):
        setattr(os,operation,recycle)
    shutil.rmtree = recycle
    report = {'schema_version':'1.0', 'ok':False,
              'variant':'normal' if ENABLE_UPDATE_CHECK else 'no-update',
              'frozen':bool(getattr(sys,'frozen',False)), 'checks':{},
              'boundaries':['Synthetic smoke only; real business samples, Acrobat, signing, and legal qualification remain pending.',
                            'Network observation covers Python socket attempts in this process; normal live update is not exercised.']}
    network_attempts = []
    def observe(event, _args):
        if event in ('os.remove','os.rmdir','shutil.rmtree'):
            raise RuntimeError('Physical deletion is disabled during release selfcheck')
        if event in ('socket.connect','socket.getaddrinfo','socket.sendto'):
            network_attempts.append(event)
            raise RuntimeError('Release selfcheck blocks network access')
    # This command exits after the check. The process-local audit hook also
    # covers updater imports and startup/About routes, not only PDF operations.
    sys.addaudithook(observe)
    os.environ['QT_QPA_PLATFORM'] = 'offscreen'
    app = view = controller = None
    current = 'gui'
    try:
        from PySide6.QtCore import QSettings
        from PySide6.QtWidgets import QApplication
        from ratools_pdf.ui.main_window import MainWindow
        from ratools_pdf.controllers.main_controller import MainController
        settings = root/'settings/settings.ini'
        settings.parent.mkdir()
        initial = QSettings(str(settings),QSettings.IniFormat)
        # Normal builds keep their real entry but do not perform a live update
        # during offline smoke. NoUpdate also exercises the disabled route below.
        initial.setValue('Update/LastSilentCheckDate',date.today().isoformat())
        initial.sync()
        app = QApplication.instance() or QApplication(['release-selfcheck'])
        view = MainWindow(settings_path=str(settings))
        controller = MainController(view)
        view.show()
        app.processEvents()
        controller.check_updates_on_startup()
        view.btn_top_about.click()
        app.processEvents()
        if not hasattr(view,'about_dialog'):
            raise RuntimeError('About route failed')
        about = view.about_dialog
        update_entry = hasattr(about,'btn_check_updates')
        if not ENABLE_UPDATE_CHECK:
            controller.updates.check_updates_manually()
        updater_imported = 'ratools_pdf.services.update_checker' in sys.modules
        updater_available = importlib.util.find_spec('ratools_pdf.services.update_checker') is not None
        worker_created = controller.updates.update_worker is not None
        if update_entry != ENABLE_UPDATE_CHECK or worker_created or network_attempts:
            raise RuntimeError('Unexpected updater entry, worker or network request')
        if not ENABLE_UPDATE_CHECK and (updater_imported or (report['frozen'] and updater_available)):
            raise RuntimeError('NoUpdate contains or imports the updater')
        report['checks']['update_boundary'] = {'status':'passed','update_entry':update_entry,
            'updater_imported':updater_imported,'updater_available':updater_available,
            'worker_created':worker_created,'network_attempts':len(network_attempts),
            'startup':'disabled' if not ENABLE_UPDATE_CHECK else 'daily marker suppresses live update',
            'observation':'Python socket audit; real startup controller and About button route'}
        about.close()
        if not view.close():
            raise RuntimeError('GUI refused idle close')
        view.app_settings.sync()
        controller.shutdown()
        app.processEvents()
        report['checks']['gui'] = {'status':'passed','initialized':True,'closed':True,
                                    'settings':'settings/settings.ini','platform':'offscreen'}

        current = 'core_pdf'
        import fitz
        from ratools_pdf.pdf.processor import process_document
        source, output = root/'release-core.pdf', root/'core-output.pdf'
        with fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((72,72),'Synthetic release smoke')
            doc.set_metadata({'title':'before'})
            doc.save(source)
        source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
        ok, _message = process_document(str(source),str(output),{'title_from_filename','open_page_first'},processing_mode='force')
        if not ok:
            raise RuntimeError('Core PDF pipeline failed')
        with fitz.open(output) as doc:
            if doc.metadata.get('title') != 'release-core' or doc.page_count != 1 or not doc[0].get_pixmap().samples:
                raise RuntimeError('Core PDF metadata/page/render postcondition failed')
        if hashlib.sha256(source.read_bytes()).hexdigest() != source_hash:
            raise RuntimeError('Core pipeline changed the input')
        report['checks']['core_pdf'] = {'status':'passed','rules':['title_from_filename','open_page_first'],
            'input_sha256':source_hash,'output_sha256':hashlib.sha256(output.read_bytes()).hexdigest(),
            'input_preserved':True,'page_count':1,'rendered':True}

        current = 'qpdf'
        from ratools_pdf.pdf.qpdf import check_pdf_syntax, get_qpdf_path, run_qpdf_command
        qpdf_path = Path(get_qpdf_path())
        if report['frozen'] and not qpdf_path.resolve().is_relative_to(Path(sys._MEIPASS).resolve()):
            raise RuntimeError('Frozen qpdf resolved outside the bundle')
        version = run_qpdf_command(['--version'],timeout=20)
        rewritten = root/'qpdf-output.pdf'
        rewrite = run_qpdf_command(['--linearize',str(output),str(rewritten)],timeout=30)
        syntax = check_pdf_syntax(str(rewritten)) if rewrite.returncode == 0 else {}
        if version.returncode or rewrite.returncode or not syntax.get('is_healthy'):
            raise RuntimeError('qpdf version/rewrite/syntax failed')
        report['checks']['qpdf'] = {'status':'passed','version':version.stdout.splitlines()[0],
            'executable_sha256':hashlib.sha256(qpdf_path.read_bytes()).hexdigest(),
            'syntax_returncode':syntax['returncode'],'rewrite_returncode':rewrite.returncode,
            'process_control':'run_owned_command (Windows atomic Job ownership)'}

        current = 'image_compression'
        from ratools_pdf.services.compression_check import verify_image_compression
        if verify_image_compression(root/'image-compression'):
            raise RuntimeError('Actual image compression check failed')
        compression = json.loads((root/'image-compression/verification.json').read_text(encoding='utf-8'))
        report['checks']['image_compression'] = dict(compression,status='passed')
        report['ok'] = not network_attempts
    except Exception as exc:
        report['checks'][current] = {'status':'failed','error_type':type(exc).__name__,
                                     'reason':str(exc) if isinstance(exc,RuntimeError) else 'Check could not complete'}
    finally:
        if controller is not None:
            controller.shutdown()
        if view is not None:
            view.close()
            view.app_settings.sync()
        if app is not None:
            app.processEvents()
        report['network_attempts'] = len(network_attempts)
        (root/'release-verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    return 0 if report['ok'] else 1
