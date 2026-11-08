"""CSV round trips preserve named PDF actions, and reject unsupported data."""

import csv
import tempfile
import unittest
from pathlib import Path

import fitz

from ratools_pdf.pdf.bookmarks_links import export_bookmarks, import_bookmarks


class BookmarkCsvTargetTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.source, self.csv, self.output = (self.root/name for name in ('source.pdf','toc.csv','out.pdf'))

    def make_source(self, actions):
        with fitz.open() as doc:
            doc.new_page()
            doc.new_page()
            name = fitz.get_pdf_str('Section (one) \\ appendix')
            doc.xref_set_key(doc.pdf_catalog(), 'Names',
                             f'<</Dests<</Names[{name} [{doc.page_xref(1)} 0 R /XYZ 42 700 1.5]]>>>>')
            doc.set_toc([[1 if index == 0 else 2, f'Entry {index}', 1] for index in range(len(actions))])
            for xref, action in zip(doc.get_outline_xrefs(), actions):
                doc.xref_set_key(xref, 'A', action)
            doc.save(self.source)

    def test_local_named_destination_preserves_name_hierarchy_and_resolved_target(self):
        name = fitz.get_pdf_str('Section (one) \\ appendix')
        self.make_source([f'<</S/GoTo/D{name}>>', '<</S/URI/URI(https://example.org/path)>>'])
        with fitz.open(self.source) as doc:
            before = doc.get_toc(simple=False)
            resolved = doc.resolve_names()
        export_bookmarks(self.source, self.csv)
        import_bookmarks(self.source, self.csv, self.output)
        with fitz.open(self.output) as doc:
            after = doc.get_toc(simple=False)
            self.assertEqual(after[0][3]['kind'], fitz.LINK_NAMED)
            self.assertEqual(after[0][3]['nameddest'], before[0][3]['nameddest'])
            self.assertEqual(after[0][3]['page'], 1)
            self.assertEqual(doc.resolve_names(), resolved)
            self.assertEqual(after[1][:3], before[1][:3])
            self.assertEqual(after[1][3]['uri'], 'https://example.org/path')

    def test_remote_named_and_explicit_destinations_preserve_action_and_new_window(self):
        self.make_source([
            '<</S/GoToR/F<</Type/Filespec/F(remote.pdf)/UF(remote.pdf)>>/D(section)/NewWindow true>>',
            '<</S/GoToR/F(remote.pdf)/D[3/XYZ 24 48 2]/NewWindow false>>',
        ])
        original = self.source.read_bytes()
        export_bookmarks(self.source, self.csv)
        self.assertEqual(self.source.read_bytes(), original)
        import_bookmarks(self.source, self.csv, self.output)
        with fitz.open(self.output) as doc:
            first, second = doc.get_outline_xrefs()
            self.assertEqual(doc.xref_get_key(first, 'A/S'), ('name','/GoToR'))
            self.assertEqual(doc.xref_get_key(first, 'A/D'), ('string','section'))
            self.assertEqual(doc.xref_get_key(first, 'A/F/UF'), ('string','remote.pdf'))
            self.assertEqual(doc.xref_get_key(first, 'A/NewWindow'), ('bool','true'))
            self.assertEqual(doc.xref_get_key(second, 'A/S'), ('name','/GoToR'))
            self.assertEqual(doc.xref_get_key(second, 'A/D'), ('array','[3/XYZ 24 48 2]'))
            self.assertEqual(doc.xref_get_key(second, 'A/NewWindow'), ('bool','false'))

    def test_unsupported_pdf_action_fails_before_overwriting_existing_csv(self):
        self.make_source(['<</S/Named/N/NextPage>>'])
        self.csv.write_text('previous export', encoding='utf-8')
        with self.assertRaisesRegex(ValueError, '不支持'):
            export_bookmarks(self.source, self.csv)
        self.assertEqual(self.csv.read_text(encoding='utf-8'), 'previous export')

    def test_unsupported_csv_kind_is_explicit_error_and_preserves_existing_output(self):
        self.make_source(['<</S/URI/URI(https://example.org)>>'])
        self.csv.write_text('Level,Title,Page,Kind\n1,unsupported,1,999\n', encoding='utf-8-sig')
        self.output.write_bytes(b'earlier PDF')
        with self.assertRaisesRegex(ValueError, '不支持'):
            import_bookmarks(self.source, self.csv, self.output)
        self.assertEqual(self.output.read_bytes(), b'earlier PDF')

    def test_malformed_named_target_is_reported_instead_of_becoming_empty_action(self):
        self.make_source(['<</S/URI/URI(https://example.org)>>'])
        self.csv.write_text(f'Level,Title,Page,Kind,NamedDest\n1,missing,1,{fitz.LINK_NAMED},\n', encoding='utf-8-sig')
        with self.assertRaisesRegex(ValueError, '命名目标'):
            import_bookmarks(self.source, self.csv, self.output)
        self.assertFalse(self.output.exists())

    def test_old_three_column_csv_remains_supported(self):
        self.make_source(['<</S/URI/URI(https://example.org)>>'])
        self.csv.write_text('Level,Title,Page\n1,Legacy title,2\n', encoding='utf-8-sig')
        import_bookmarks(self.source, self.csv, self.output)
        with fitz.open(self.output) as doc:
            self.assertEqual(doc.get_toc(), [[1,'Legacy title',2]])

    def test_remote_pdf_name_escapes_are_preserved_as_the_same_target(self):
        self.make_source(['<</S/GoToR/F(remote.pdf)/D/section#20one>>'])
        export_bookmarks(self.source, self.csv)
        with self.csv.open(encoding='utf-8-sig', newline='') as stream:
            self.assertEqual(next(csv.DictReader(stream))['NamedDest'], 'section one')
        import_bookmarks(self.source, self.csv, self.output)
        with fitz.open(self.output) as doc:
            xref = doc.get_outline_xrefs()[0]
            self.assertEqual(doc.xref_get_key(xref, 'A/D'), ('string','section one'))

    def test_absent_new_window_setting_retains_viewer_default(self):
        self.make_source(['<</S/GoToR/F(remote.pdf)/D(section)>>'])
        export_bookmarks(self.source, self.csv)
        import_bookmarks(self.source, self.csv, self.output)
        with fitz.open(self.output) as doc:
            xref = doc.get_outline_xrefs()[0]
            self.assertEqual(doc.xref_get_key(xref, 'A/NewWindow'), ('null','null'))

    def test_launch_action_remains_launch_with_original_window_setting(self):
        self.make_source(['<</S/Launch/F(other.pdf)/NewWindow true>>',
                          '<</S/Launch/F(other.pdf)/NewWindow false>>',
                          '<</S/Launch/F(other.pdf)>>'])
        export_bookmarks(self.source, self.csv)
        import_bookmarks(self.source, self.csv, self.output)
        with fitz.open(self.output) as doc:
            for xref, expected in zip(doc.get_outline_xrefs(), [('bool','true'),('bool','false'),('null','null')]):
                self.assertEqual(doc.xref_get_key(xref, 'A/S'), ('name','/Launch'))
                self.assertEqual(doc.xref_get_key(xref, 'A/F/UF'), ('string','other.pdf'))
                self.assertEqual(doc.xref_get_key(xref, 'A/NewWindow'), expected)


if __name__ == '__main__':
    unittest.main()
