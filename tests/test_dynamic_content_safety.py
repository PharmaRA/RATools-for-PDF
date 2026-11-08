"""Dynamic cleanup preserves non-dynamic navigation, files and annotations."""

import tempfile
import unittest
from pathlib import Path

import fitz

from ratools_pdf.pdf.processor import _PipelineContext, _PhaseProfiler, _step_cleanup, process_document
from ratools_pdf.pdf.precheck import build_precheck_report
from ratools_pdf.pdf.dynamic_content import has_dynamic_content


class DynamicContentSafetyTests(unittest.TestCase):
    def document(self):
        doc = fitz.open()
        self.addCleanup(doc.close)
        doc.new_page().insert_text((72,72),'Keep this text')
        return doc

    def clean(self, doc):
        context = _PipelineContext(doc, {'cleanup_remove_dynamic_content'}, 'synthetic.pdf', _PhaseProfiler())
        _step_cleanup(context)
        return context

    def test_cleanup_preserves_named_destinations_and_embedded_files(self):
        doc = self.document()
        doc.embfile_add('keep.txt',b'Keep this attachment')
        catalog = doc.pdf_catalog()
        names = int(doc.xref_get_key(catalog,'Names')[1].split()[0]) if doc.xref_get_key(catalog,'Names')[0]=='xref' else catalog
        prefix = '' if names != catalog else 'Names/'
        doc.xref_set_key(names,prefix+'Dests',f'<< /Names [(Target) [{doc[0].xref} 0 R /XYZ 72 700 0]] >>')
        doc.xref_set_key(names,prefix+'JavaScript','<< /Names [(Run) << /S /JavaScript /JS (void 0;) >>] >>')
        doc.set_toc([[1,'Target',1]])
        outline = doc.get_toc(simple=False)[0][3]['xref']
        doc.xref_set_key(outline,'A','<< /S /GoTo /D (Target) >>')
        self.clean(doc)
        with fitz.open(stream=doc.tobytes(),filetype='pdf') as result:
            self.assertEqual(result.embfile_get('keep.txt'),b'Keep this attachment')
            self.assertIn('Target',result.resolve_names())
            self.assertEqual(result.get_toc(simple=False)[0][3]['page'],0)
            self.assertEqual(result.xref_get_key(result.pdf_catalog(),'Names/JavaScript')[0],'null')

    def test_open_action_script_is_removed_but_following_local_navigation_is_kept(self):
        doc = self.document()
        doc.xref_set_key(doc.pdf_catalog(),'OpenAction',
                        f'<< /S /JavaScript /JS (void 0;) /Next << /S /GoTo /D [{doc[0].xref} 0 R /Fit] >> >>')
        self.clean(doc)
        self.assertEqual(doc.xref_get_key(doc.pdf_catalog(),'OpenAction/S'),('name','/GoTo'))
        self.assertEqual(doc.xref_get_key(doc.pdf_catalog(),'OpenAction/JS')[0],'null')

    def test_scripts_in_page_annotation_and_field_actions_are_removed(self):
        doc = self.document()
        page = doc[0]
        annotation = page.add_text_annot((72,100),'Keep this comment')
        widget = fitz.Widget()
        widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
        widget.field_name = 'field'
        widget.field_value = 'Keep this value'
        widget.rect = fitz.Rect(72,150,220,180)
        field = page.add_widget(widget)
        for owner in (doc.pdf_catalog(),page.xref,annotation.xref,field.xref):
            doc.xref_set_key(owner,'AA','<< /O << /S /JavaScript /JS (void 0;) >> >>')
        doc.xref_set_key(annotation.xref,'A','<< /S /URI /URI (https://example.invalid/) '
                         '/Next [<< /S /JavaScript /JS (void 0;) >> << /S /Named /N /NextPage >>] >>')
        self.clean(doc)
        for owner in (doc.pdf_catalog(),page.xref,annotation.xref,field.xref):
            self.assertEqual(doc.xref_get_key(owner,'AA/O')[0],'null')
        self.assertEqual(doc.xref_get_key(annotation.xref,'A/S'),('name','/URI'))
        self.assertNotIn('/JavaScript',doc.xref_get_key(annotation.xref,'A/Next')[1])
        self.assertIn('/NextPage',doc.xref_get_key(annotation.xref,'A/Next')[1])
        with fitz.open(stream=doc.tobytes(),filetype='pdf') as result:
            self.assertEqual(next(result[0].annots()).info['content'],'Keep this comment')
            self.assertEqual(next(result[0].widgets()).field_value,'Keep this value')

    def test_3d_and_rich_media_are_removed_without_losing_static_annotations(self):
        doc = self.document()
        page = doc[0]
        comment = page.add_text_annot((72,100),'Keep comment')
        doc.xref_set_key(page.xref,'Annots',
                        f'[{comment.xref} 0 R << /Type /Annot /Subtype /3D /Rect [10 10 50 50] /3DD << >> >> '
                        '<< /Type /Annot /Subtype /RichMedia /Rect [60 10 90 50] /RichMediaContent << >> >>]')
        self.clean(doc)
        raw = doc.xref_get_key(page.xref,'Annots')[1]
        self.assertNotIn('/3D',raw)
        self.assertNotIn('/RichMedia',raw)
        with fitz.open(stream=doc.tobytes(),filetype='pdf') as result:
            self.assertEqual(len(list(result[0].annots())),1)

    def test_smart_precheck_detects_open_action_without_names_and_pipeline_removes_it(self):
        doc = self.document()
        doc.xref_set_key(doc.pdf_catalog(),'OpenAction','<< /S /JavaScript /JS (void 0;) >>')
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory)/'source.pdf'
            output = Path(directory)/'output.pdf'
            doc.save(source)
            before = source.read_bytes()
            report = build_precheck_report(str(source),selected_options={'cleanup_remove_dynamic_content'})
            self.assertTrue(report['available'])
            self.assertIn('cleanup_remove_dynamic_content',report['suggestions'])
            self.assertEqual(source.read_bytes(),before,'precheck must not modify input')
            ok,message = process_document(str(source),str(output),{'cleanup_remove_dynamic_content'})
            self.assertTrue(ok,message)
            with fitz.open(output) as result:
                self.assertEqual(result.xref_get_key(result.pdf_catalog(),'OpenAction')[0],'null')
                self.assertIn('Keep this text',result[0].get_text())

    def test_static_names_and_navigation_are_not_marked_as_changed(self):
        doc = self.document()
        doc.embfile_add('keep.txt',b'Keep')
        doc.xref_set_key(doc.pdf_catalog(),'OpenAction',f'[{doc[0].xref} 0 R /Fit]')
        context = self.clean(doc)
        self.assertFalse(context.changed)
        self.assertEqual(doc.embfile_get('keep.txt'),b'Keep')

    def test_detection_is_read_only_for_shared_script_and_cyclic_actions(self):
        doc = self.document()
        script = doc.get_new_xref()
        doc.update_object(script,f'<< /S /JavaScript /JS (void 0;) /Next {script} 0 R >>')
        doc.xref_set_key(doc.pdf_catalog(),'OpenAction',f'{script} 0 R')
        doc.set_toc([[1,'Bookmark',1]])
        outline = doc.get_toc(simple=False)[0][3]['xref']
        doc.xref_set_key(outline,'A',f'{script} 0 R')
        before = [doc.xref_object(i) for i in range(1,doc.xref_length())]
        self.assertTrue(has_dynamic_content(doc))
        self.assertEqual([doc.xref_object(i) for i in range(1,doc.xref_length())],before)
        self.clean(doc)
        self.assertFalse(has_dynamic_content(doc))
        self.assertEqual(doc.xref_get_key(doc.pdf_catalog(),'OpenAction')[0],'null')
        self.assertEqual(doc.xref_get_key(outline,'A')[0],'null')

    def test_removing_script_promotes_multiple_safe_next_actions_in_order(self):
        doc = self.document()
        doc.xref_set_key(doc.pdf_catalog(),'OpenAction',
                        '<< /S /JavaScript /JS (void 0;) /Next ['
                        '<< /S /Named /N /NextPage >> << /S /Named /N /LastPage >>] >>')
        self.clean(doc)
        self.assertEqual(doc.xref_get_key(doc.pdf_catalog(),'OpenAction/N'),('name','/NextPage'))
        self.assertIn('/LastPage',doc.xref_get_key(doc.pdf_catalog(),'OpenAction/Next')[1])
        self.assertFalse(has_dynamic_content(doc))

    def test_nested_form_field_script_is_cleaned_without_changing_field_value(self):
        doc = self.document()
        field = doc.get_new_xref()
        doc.update_object(field,'<< /FT /Tx /T (Nested) /V (Preserved) '
                          '/AA << /K << /S /JavaScript /JS (void 0;) >> >> >>')
        parent = doc.get_new_xref()
        doc.update_object(parent,f'<< /T (Parent) /Kids [{field} 0 R] >>')
        doc.xref_set_key(doc.pdf_catalog(),'AcroForm',f'<< /Fields [{parent} 0 R] >>')
        self.assertTrue(has_dynamic_content(doc))
        self.clean(doc)
        self.assertEqual(doc.xref_get_key(field,'AA')[0],'null')
        self.assertEqual(doc.xref_get_key(field,'V'),('string','Preserved'))
        self.assertFalse(has_dynamic_content(doc))

    def test_precheck_detects_rich_media_without_catalog_names(self):
        doc = self.document()
        appearance = doc.get_new_xref()
        doc.update_object(appearance,'<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << >> >>')
        doc.update_stream(appearance,b'q Q')
        doc.xref_set_key(doc[0].xref,'Annots',
                        f'[<< /Subtype /RichMedia /Rect [0 0 10 10] /AP << /N {appearance} 0 R >> >>]')
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory)/'media.pdf'
            output = Path(directory)/'clean.pdf'
            doc.save(path)
            report = build_precheck_report(str(path),selected_options={'cleanup_remove_dynamic_content'})
            self.assertIn('cleanup_remove_dynamic_content',report['suggestions'])
            ok,message = process_document(str(path),str(output),{'cleanup_remove_dynamic_content'})
            self.assertTrue(ok,message)
            with fitz.open(output) as result:
                self.assertFalse(has_dynamic_content(result))
                self.assertEqual(result[0].get_text().strip(),'Keep this text')


if __name__ == '__main__':
    unittest.main()
