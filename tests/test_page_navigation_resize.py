"""Resize local destinations by the target page's original PDF coordinates."""

import unittest
import tempfile
from pathlib import Path

import fitz

from ratools_pdf.pdf.page_layout import resize_pages_with_padding
from ratools_pdf.pdf.processor import process_document


class NavigationResizeTests(unittest.TestCase):
    def document(self, sizes):
        doc = fitz.open()
        self.addCleanup(doc.close)
        for width,height in sizes:
            doc.new_page(width=width,height=height).insert_text((20,100),'TARGET')
        return doc

    def array(self, doc, xref, key):
        kind,value = doc.xref_get_key(xref,key)
        if kind == 'xref':
            value = doc.xref_object(int(value.split()[0]))
        return ' '.join(value.replace('[',' ').replace(']',' ').replace('/',' /').split())

    def test_bookmark_coordinates_are_read_before_page_height_changes(self):
        doc = self.document([(200,400)])
        doc.set_toc([[1,'Target',1,{'kind':fitz.LINK_GOTO,'page':0,'to':fitz.Point(20,100),'zoom':1.5}]])
        resize_pages_with_padding(doc,fitz.Rect(0,0,400,800))
        with fitz.open(stream=doc.tobytes(),filetype='pdf') as result:
            dest = result.get_toc(simple=False)[0][3]
            self.assertEqual(tuple(dest['to']),(40,200))
            self.assertEqual(dest['zoom'],1.5)

    def test_links_from_unchanged_pages_follow_scaled_targets_in_both_orders(self):
        for target_index in (0,1):
            with self.subTest(target_index=target_index):
                sizes = [(400,800),(400,800)]
                sizes[target_index] = (200,400)
                doc = self.document(sizes)
                source_index = 1-target_index
                page = doc[source_index]
                page.insert_link({'kind':fitz.LINK_GOTO,'from':fitz.Rect(10,10,80,40),
                                  'page':target_index,'to':fitz.Point(20,100)})
                resize_pages_with_padding(doc,fitz.Rect(0,0,400,800))
                with fitz.open(stream=doc.tobytes(),filetype='pdf') as result:
                    link = result[source_index].get_links()[0]
                    self.assertEqual(tuple(link['to']),(40,200))
                    self.assertEqual(tuple(link['from']),(10,10,80,40))

    def test_rotated_target_preserves_raw_destination_and_link_action_flags(self):
        doc = self.document([(600,800),(200,400)])
        doc[1].set_rotation(90)
        target = doc[1].xref
        annotation = doc.get_new_xref()
        doc.update_object(annotation,f'<< /Type /Annot /Subtype /Link /Rect [10 750 80 780] '
                          f'/A << /S /GoTo /D [{target} 0 R /XYZ 20 300 2.5] /NewWindow true >> >>')
        doc.xref_set_key(doc[0].xref,'Annots',f'[{annotation} 0 R]')
        resize_pages_with_padding(doc,fitz.Rect(0,0,600,800))
        self.assertEqual(self.array(doc,annotation,'A/D'),f'{target} 0 R /XYZ 140 600 2.5')
        self.assertEqual(doc.xref_get_key(annotation,'A/NewWindow'),('bool','true'))

    def test_shared_named_destinations_are_scaled_once_and_names_preserved(self):
        doc = self.document([(200,400)])
        target = doc[0].xref
        dest = doc.get_new_xref()
        doc.update_object(dest,f'[{target} 0 R /XYZ 20 300 2]')
        leaf = doc.get_new_xref()
        doc.update_object(leaf,f'<< /Names [(Alpha) {dest} 0 R (Beta) << /D {dest} 0 R >>] >>')
        doc.xref_set_key(doc.pdf_catalog(),'Names',f'<< /Dests << /Kids [{leaf} 0 R] >> >>')
        doc.xref_set_key(doc.pdf_catalog(),'Dests',f'<< /Legacy {dest} 0 R >>')
        doc.set_toc([[1,'Target',1]])
        outline = doc.get_toc(simple=False)[0][3]['xref']
        doc.xref_set_key(outline,'A','<< /S /GoTo /D (Alpha) >>')
        resize_pages_with_padding(doc,fitz.Rect(0,0,400,800))
        names = doc.resolve_names()
        for name in ('Alpha','Beta','Legacy'):
            self.assertEqual(tuple(names[name]['to']),(40,600))
            self.assertEqual(names[name]['zoom'],2)
        self.assertEqual(doc.xref_get_key(outline,'A/D'),('string','Alpha'))

    def test_open_and_additional_actions_transform_fit_coordinates_preserving_nulls(self):
        doc = self.document([(200,400)])
        target = doc[0].xref
        catalog = doc.pdf_catalog()
        doc.xref_set_key(catalog,'OpenAction',f'[{target} 0 R /XYZ null 300 0]')
        doc.xref_set_key(doc[0].xref,'AA',
                        f'<< /O << /S /GoTo /D [{target} 0 R /FitH 300] '
                        f'/Next [<< /S /GoTo /D [{target} 0 R /FitV 20] >> '
                        f'<< /S /GoTo /D [{target} 0 R /FitR 10 20 100 300] >>] >> >>')
        resize_pages_with_padding(doc,fitz.Rect(0,0,400,800))
        self.assertEqual(self.array(doc,catalog,'OpenAction'),f'{target} 0 R /XYZ null 600 0')
        self.assertEqual(self.array(doc,doc[0].xref,'AA/O/D'),f'{target} 0 R /FitH 600')
        self.assertIn('/FitV 40',doc.xref_get_key(doc[0].xref,'AA/O/Next')[1])
        self.assertIn('/FitR 20 40 200 600',doc.xref_get_key(doc[0].xref,'AA/O/Next')[1])

    def test_external_destination_is_not_transformed_when_sharing_local_array(self):
        doc = self.document([(200,400)])
        target = doc[0].xref
        destination = doc.get_new_xref()
        doc.update_object(destination,'[0 /XYZ 20 300 2]')
        doc.xref_set_key(doc.pdf_catalog(),'OpenAction',f'{destination} 0 R')
        link = doc.get_new_xref()
        doc.update_object(link,f'<< /Type /Annot /Subtype /Link /Rect [10 300 50 330] '
                          f'/A << /S /GoToR /F (other.pdf) /D {destination} 0 R /NewWindow true >> >>')
        doc.xref_set_key(target,'Annots',f'[{link} 0 R]')
        resize_pages_with_padding(doc,fitz.Rect(0,0,400,800))
        self.assertEqual(self.array(doc,doc.pdf_catalog(),'OpenAction'),'0 /XYZ 40 600 2')
        self.assertEqual(self.array(doc,link,'A/D'),'0 /XYZ 20 300 2')
        self.assertEqual(doc.xref_get_key(link,'A/NewWindow'),('bool','true'))

    def test_scaled_source_preserves_unchanged_target_and_scales_inline_hit_area(self):
        doc = self.document([(200,400),(400,800)])
        target = doc[1].xref
        doc.xref_set_key(doc[0].xref,'Annots',
                        f'[<< /Type /Annot /Subtype /Link /Rect [10 300 50 330] '
                        f'/Dest [{target} 0 R /XYZ 20 700 0] >>]')
        resize_pages_with_padding(doc,fitz.Rect(0,0,400,800))
        with fitz.open(stream=doc.tobytes(),filetype='pdf') as result:
            link = result[0].get_links()[0]
            self.assertEqual(tuple(link['from']),(20,140,100,200))
            self.assertEqual(tuple(link['to']),(20,100))

    def test_cyclic_action_chain_is_processed_once(self):
        doc = self.document([(200,400)])
        target = doc[0].xref
        action = doc.get_new_xref()
        doc.update_object(action,f'<< /S /GoTo /D [{target} 0 R /XYZ 20 300 0] /Next {action} 0 R >>')
        doc.xref_set_key(doc.pdf_catalog(),'OpenAction',f'{action} 0 R')
        resize_pages_with_padding(doc,fitz.Rect(0,0,400,800))
        self.assertEqual(self.array(doc,action,'D'),f'{target} 0 R /XYZ 40 600 0')
        self.assertEqual(doc.xref_get_key(action,'Next'),('xref',f'{action} 0 R'))

    def test_saved_pipeline_preserves_bookmark_hierarchy_and_inline_destination(self):
        doc = self.document([(200,400)])
        doc.set_toc([[1,'Parent',1],[2,'Child',1]])
        outlines = doc.get_toc(simple=False)
        for entry in outlines:
            doc.xref_set_key(entry[3]['xref'],'A/D',f'[{doc[0].xref} 0 R /XYZ 20 300 1.5]')
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory)/'source.pdf'
            output = Path(directory)/'output.pdf'
            doc.save(source)
            ok,message = process_document(str(source),str(output),{'page_size_letter'},processing_mode='force')
            self.assertTrue(ok,message)
            with fitz.open(output) as result:
                toc = result.get_toc(simple=False)
                self.assertEqual([(item[0],item[1]) for item in toc],[(1,'Parent'),(2,'Child')])
                # Letter 612x792: scale=1.98, horizontal padding=108.
                for item in toc:
                    self.assertAlmostEqual(item[3]['to'].x,147.6,delta=.01)
                    self.assertAlmostEqual(item[3]['to'].y,198,delta=.01)
                    self.assertEqual(item[3]['zoom'],1.5)

    def test_remote_open_action_is_unchanged_but_its_local_next_action_is_scaled(self):
        for indirect in (False,True):
            with self.subTest(indirect=indirect):
                doc = self.document([(200,400)])
                target = doc[0].xref
                action_text = (f'<< /S /GoToR /F (other.pdf) /D [0 /XYZ 20 300 2] /NewWindow true '
                               f'/Next << /S /GoTo /D [{target} 0 R /XYZ 20 300 0] >> >>')
                catalog = doc.pdf_catalog()
                if indirect:
                    action = doc.get_new_xref()
                    doc.update_object(action,action_text)
                    doc.xref_set_key(catalog,'OpenAction',f'{action} 0 R')
                else:
                    doc.xref_set_key(catalog,'OpenAction',action_text)
                resize_pages_with_padding(doc,fitz.Rect(0,0,400,800))
                self.assertEqual(self.array(doc,catalog,'OpenAction/D'),'0 /XYZ 20 300 2')
                self.assertEqual(self.array(doc,catalog,'OpenAction/Next/D'),f'{target} 0 R /XYZ 40 600 0')
                self.assertEqual(doc.xref_get_key(catalog,'OpenAction/NewWindow'),('bool','true'))


if __name__ == '__main__':
    unittest.main()
