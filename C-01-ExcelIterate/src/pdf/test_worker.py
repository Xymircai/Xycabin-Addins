"""End-to-end tests against the compiled worker (test-only: pymupdf and Pillow).

Run: python src/pdf/build_worker.py && python src/pdf/test_worker.py
Retains fixtures and rendered samples in build/pdf-tests for review.
"""
from pathlib import Path
import hashlib
import uuid
import subprocess
import unittest
import xml.etree.ElementTree as ET

import pymupdf as pdf
from PIL import Image

ROOT = Path(__file__).resolve().parents[2]
WORKER = ROOT / 'build/pdf-worker/PdfWorker.exe'
OUT = ROOT / 'build/pdf-tests'


class WorkerTests(unittest.TestCase):
    def setUp(self):
        self.folder = OUT / (self._testMethodName + '-' + uuid.uuid4().hex[:8])
        self.folder.mkdir(parents=True, exist_ok=True)
        self.output = self.folder / 'output'
        self.output.mkdir(exist_ok=True)

    def fixture(self, name='source.pdf', count=3, rotate=False):
        path = self.folder / name
        doc = pdf.open()
        for i in range(count):
            page = doc.new_page(width=420 if i % 2 else 595, height=595 if i % 2 else 842)
            page.insert_text((50, 70), name + ' page ' + str(i + 1))
            if rotate and i == 1:
                page.set_rotation(90)
        doc.save(path)
        doc.close()
        return path

    def run_job(self, mode, files, cancel=False, **options):
        job = ET.Element('PdfJob', mode=mode, output=str(self.output), **options)
        for path in files:
            ET.SubElement(job, 'File').text = str(path)
        request = self.folder / 'request.xml'
        ET.ElementTree(job).write(request, encoding='utf-8', xml_declaration=True)
        if cancel:
            Path(str(request) + '.cancel').write_text('cancel')
        result = subprocess.run([str(WORKER), str(request)], timeout=60, capture_output=True)
        report = ET.parse(str(request) + '.result').getroot()
        trace = Path(str(request) + '.cancel.log').read_text(encoding='utf-8-sig')
        self.assertIn('worker-start pid=', trace)
        self.assertIn('worker-complete exit=' + str(result.returncode), trace)
        return result.returncode, report

    def saved(self, report):
        return [Path(node.text) for node in report.findall('Success')]

    def texts(self, path):
        with pdf.open(path) as doc:
            return [page.get_text() for page in doc]

    def test_split_ranges_preserve_order_and_source(self):
        source = self.fixture()
        before = hashlib.sha256(source.read_bytes()).digest()
        code, result = self.run_job('split', [source], ranges='2-3,1')
        self.assertEqual(code, 0, ET.tostring(result))
        files = self.saved(result)
        self.assertEqual([len(self.texts(file)) for file in files], [2, 1])
        self.assertIn('page 2', self.texts(files[0])[0])
        self.assertIn('page 1', self.texts(files[1])[0])
        self.assertEqual(before, hashlib.sha256(source.read_bytes()).digest())

    def test_split_each_page(self):
        code, result = self.run_job('split', [self.fixture()], ranges='')
        self.assertEqual(code, 0)
        self.assertEqual(len(self.saved(result)), 3)
        self.assertTrue(all(len(self.texts(p)) == 1 for p in self.saved(result)))

    def test_split_rejects_bad_ranges_before_writing(self):
        source = self.fixture()
        for value in ['0', '4', '3-1', '1-', '1,,2', '1-2,9', '-1', '1-2-3']:
            code, result = self.run_job('split', [source], ranges=value)
            self.assertEqual(code, 1, value)
            self.assertEqual(self.saved(result), [])
        self.assertEqual(list(self.output.iterdir()), [])

    def test_merge_uses_selected_order(self):
        first, second = self.fixture('first.pdf', 2), self.fixture('second.pdf', 1)
        code, result = self.run_job('merge', [second, first])
        self.assertEqual(code, 0, ET.tostring(result))
        texts = self.texts(self.saved(result)[0])
        self.assertEqual(len(texts), 3)
        self.assertIn('second.pdf', texts[0])
        self.assertIn('first.pdf page 2', texts[2])

    def test_merge_bad_input_does_not_publish(self):
        code, result = self.run_job('merge', [self.fixture(), self.folder / 'missing.pdf'])
        self.assertEqual(code, 1)
        self.assertEqual(list(self.output.iterdir()), [])

    def test_merge_preserves_image_resources_and_rotation(self):
        source = self.fixture(rotate=True)
        image_doc = self.folder / 'image.pdf'
        with pdf.open() as doc:
            page = doc.new_page()
            bitmap = pdf.Pixmap(pdf.csRGB, pdf.IRect(0, 0, 40, 40), False)
            bitmap.clear_with(80)
            page.insert_image(pdf.Rect(20, 20, 200, 200), pixmap=bitmap)
            doc.save(image_doc)
        code, result = self.run_job('merge', [source, image_doc])
        self.assertEqual(code, 0, ET.tostring(result))
        with pdf.open(self.saved(result)[0]) as doc:
            self.assertEqual(doc[1].rotation, 90)
            self.assertEqual(len(doc[3].get_images()), 1)
            self.assertGreater(len(set(doc[3].get_pixmap().samples)), 1)

    def test_watermark_chinese_rotation_and_original_text(self):
        source = self.fixture(rotate=True)
        code, result = self.run_job('watermark', [source], text='内部资料 仅供审核', opacity='25', size='48', angle='-35')
        self.assertEqual(code, 0, ET.tostring(result))
        target = self.saved(result)[0]
        with pdf.open(source) as original, pdf.open(target) as modified:
            self.assertEqual(len(original), len(modified))
            for a, b in zip(original, modified):
                self.assertEqual(a.rotation, b.rotation)
                self.assertIn('page', b.get_text())
                self.assertIn('内部资料', b.get_text().replace(' ', ''))
                self.assertNotEqual(a.get_pixmap().samples, b.get_pixmap().samples)
            modified[0].get_pixmap(matrix=pdf.Matrix(1, 1)).save(self.folder / 'watermark.png')
            modified[1].get_pixmap().save(self.folder / 'watermark-rotated.png')

    def test_watermark_invalid_parameters(self):
        source = self.fixture()
        for options in [dict(text='', opacity='20', size='48', angle='0'), dict(text='A', opacity='NaN', size='48', angle='0')]:
            code, result = self.run_job('watermark', [source], **options)
            self.assertEqual(code, 1)
            self.assertEqual(self.saved(result), [])

    def test_images_and_multi_page_tiff(self):
        png = self.folder / '中文 image.png'
        first, second = Image.new('RGB', (800, 400), 'navy'), Image.new('RGB', (400, 800), 'orange')
        first.save(png)
        tiff = self.folder / 'pages.tiff'
        first.save(tiff, save_all=True, append_images=[second])
        code, result = self.run_job('convert', [png, tiff])
        self.assertEqual(code, 0, ET.tostring(result))
        files = self.saved(result)
        self.assertEqual([len(self.texts(p)) for p in files], [1, 2])
        with pdf.open(files[0]) as doc:
            self.assertGreater(doc[0].rect.width, doc[0].rect.height)
            doc[0].get_pixmap().save(self.folder / 'image.png')

    def test_collision_and_partial_failure(self):
        source = self.fixture()
        existing = self.output / 'source_001.pdf'
        existing.write_bytes(b'existing file')
        code, result = self.run_job('split', [source, self.folder / 'missing.pdf'], ranges='1')
        self.assertEqual(code, 1)
        self.assertEqual(existing.read_bytes(), b'existing file')
        self.assertEqual(len(self.saved(result)), 1)
        self.assertNotEqual(self.saved(result)[0], existing)
        self.assertFalse(any(p.suffix == '.tmp' for p in self.output.iterdir()))

    def test_cancel_does_not_write(self):
        code, result = self.run_job('split', [self.fixture()], cancel=True, ranges='')
        self.assertEqual(code, 3)
        self.assertEqual(list(self.output.iterdir()), [])

    def test_encrypted_pdf_rejected(self):
        source = self.fixture()
        encrypted = self.folder / 'locked.pdf'
        with pdf.open(source) as doc:
            doc.save(encrypted, encryption=pdf.PDF_ENCRYPT_AES_256, owner_pw='owner', user_pw='secret')
        code, result = self.run_job('split', [encrypted], ranges='')
        self.assertEqual(code, 1)
        self.assertEqual(list(self.output.iterdir()), [])

    def test_mixed_conversion_merge_is_atomic(self):
        first, second = self.folder / 'wide.png', self.folder / 'tall.jpg'
        Image.new('RGB', (600, 200), 'red').save(first)
        Image.new('RGB', (200, 600), 'blue').save(second)
        code, result = self.run_job('convertmerge', [first, second])
        self.assertEqual(code, 0, ET.tostring(result))
        self.assertEqual(len(self.saved(result)), 1)
        with pdf.open(self.saved(result)[0]) as doc:
            self.assertEqual(len(doc), 2)
            self.assertGreater(doc[0].rect.width, doc[0].rect.height)
            self.assertGreater(doc[1].rect.height, doc[1].rect.width)

    def test_failed_mixed_merge_publishes_nothing(self):
        image = self.folder / 'image.png'
        Image.new('RGB', (80, 120), 'green').save(image)
        bad = self.folder / 'bad.txt'
        bad.write_text('unsupported')
        code, result = self.run_job('convertmerge', [image, bad])
        self.assertEqual(code, 1)
        self.assertEqual(list(self.output.iterdir()), [])

    def test_open_office_source_pauses_entire_batch(self):
        image = self.folder / 'first.png'
        Image.new('RGB', (80, 120), 'green').save(image)
        book = self.folder / 'opened.xlsx'
        book.write_bytes(b'fixture')
        with book.open('rb'):
            code, result = self.run_job('convert', [image, book])
        self.assertEqual(code, 4, ET.tostring(result))
        self.assertIsNotNone(result.find('Blocked'))
        self.assertEqual(list(self.output.iterdir()), [])

    def test_publish_current_document_ranges(self):
        source = self.fixture()
        code, result = self.run_job('publish', [source], name='current-export', ranges='3,1-2')
        self.assertEqual(code, 0, ET.tostring(result))
        texts = self.texts(self.saved(result)[0])
        self.assertEqual(len(texts), 3)
        self.assertIn('page 3', texts[0])
        self.assertIn('page 1', texts[1])

    def test_publish_whole_document_retains_bytes(self):
        source = self.fixture()
        code, result = self.run_job('publish', [source], name='full-export', ranges='')
        self.assertEqual(code, 0, ET.tostring(result))
        self.assertEqual(source.read_bytes(), self.saved(result)[0].read_bytes())

    def test_colored_tiled_watermark(self):
        source = self.fixture(count=1)
        code, result = self.run_job('watermark', [source], text='内部资料', opacity='70', size='24', angle='-30', color='FF0000', layout='tile')
        self.assertEqual(code, 0, ET.tostring(result))
        with pdf.open(self.saved(result)[0]) as doc:
            self.assertGreater(doc[0].get_text().count('内部资料'), 3)
            samples = doc[0].get_pixmap().samples
            self.assertTrue(any(samples[i] > samples[i+1] + 50 and samples[i] > samples[i+2] + 50 for i in range(0, len(samples)-2, 3)))
            doc[0].get_pixmap().save(self.folder / 'tile.png')

    def test_unsupported_conversion_reports_failure(self):
        source = self.folder / 'test.txt'
        source.write_text('test')
        code, result = self.run_job('convert', [source])
        self.assertEqual(code, 1)
        self.assertEqual(self.saved(result), [])


if __name__ == '__main__':
    unittest.main(verbosity=2)
