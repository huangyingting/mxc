import base64
import io
import json
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
import unittest
from unittest.mock import patch

from PIL import Image
from pptx import Presentation
from pptx.util import Emu, Pt

from presentation.src import build_html as H, build_pptx as D


class BuilderTests(unittest.TestCase):
    def test_notes_share_content_and_cannot_terminate_the_script(self):
        note = '\u8bb2\u8005\u5907\u6ce8\n</script><script>alert("notes")</script>'
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with patch.object(D, 'slides', {0: []}), patch.object(D, 'notes', {0: note}):
                H.build(str(root / 'index.html'), str(root), 1600)
                D.build(str(root / 'deck.pptx'), str(root))
            doc = (root / 'index.html').read_text(encoding='utf-8')
            match = re.search(r'window\.SLIDE_NOTES=(.*?);</script>', doc)
            self.assertIsNotNone(match)
            self.assertEqual(json.loads(match.group(1)), {'0': note})
            self.assertNotIn(note, doc)
            self.assertEqual(Presentation(root / 'deck.pptx').slides[0].notes_slide.notes_text_frame.text, note)

    def test_paragraph_spacing_and_links_survive_both_renderers(self):
        with patch.object(D, 'slides', {}):
            D.tb(0, 0, 0, 3000000, paras=[
                {'text': 'Reference', 'href': 'https://example.com/?a=1&b=2', 'spaceAfterPt': 9},
                {'text': 'Explanation'},
            ], space_after=4)
            operation = D.slides[0][0]
            markup = H.text_html(operation)
            self.assertIn('margin-bottom:12.0px;', markup)
            self.assertIn('href="https://example.com/?a=1&amp;b=2"', markup)
            presentation = Presentation()
            slide = presentation.slides.add_slide(presentation.slide_layouts[6])
            frame = slide.shapes.add_textbox(0, 0, Emu(3000000), Emu(2000000)).text_frame
            D.write_text(frame, operation)
            self.assertEqual(frame.paragraphs[0].space_after, Pt(9))
            self.assertEqual(frame.paragraphs[0].runs[0].hyperlink.address, 'https://example.com/?a=1&b=2')
            D.tb(1, 0, 0, 3000000, paras=[{'text': 'Reference', 'href': 'https://example.com'}])
            self.assertIn('href="https://example.com"', H.text_html(D.slides[1][0]))

    def test_multiline_code_preserves_indentation(self):
        code = '{\n  "nested": {\n    "enabled": true\n  }\n}'
        with patch.object(D, 'slides', {}):
            D.tb(0, 0, 0, 3000000, code, font=D.MF)
            operation = D.slides[0][0]
            markup = H.text_html(operation)
            self.assertIn('white-space:pre-wrap;', markup)
            self.assertIn('    &quot;enabled&quot;: true', markup)
            presentation = Presentation()
            slide = presentation.slides.add_slide(presentation.slide_layouts[6])
            frame = slide.shapes.add_textbox(0, 0, Emu(3000000), Emu(2000000)).text_frame
            D.write_text(frame, operation)
            self.assertEqual(frame.text.replace('\v', '\n'), code)

    def test_image_cache_respects_requested_resolution(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'wide.png'
            Image.new('RGB', (400, 200)).save(path)
            cache = {}
            for maxw in (100, 200):
                uri = H.data_uri(str(path), maxw, cache)
                image = Image.open(io.BytesIO(base64.b64decode(uri.split(',', 1)[1])))
                self.assertEqual(image.size, (maxw, maxw // 2))

    def test_svg_remains_vector_in_html_and_is_contained_in_pptx(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            svg = b'<svg xmlns="http://www.w3.org/2000/svg" width="400" height="200"/>'
            (root / 'wide.svg').write_bytes(svg)
            Image.new('RGB', (400, 200)).save(root / 'wide.png')
            uri = H.data_uri(str(root / 'wide.svg'), 1600, {})
            self.assertEqual(base64.b64decode(uri.split(',', 1)[1]), svg)
            with patch.object(D, 'slides', {}):
                D.img(0, 'wide.svg', 0, 0, 2000000, 2000000, 'Chart', fit='contain')
                operation = D.slides[0][0]
                self.assertIn('contain-img', H.slide_html(0, [operation], str(root), 1600))
                presentation = Presentation()
                slide = presentation.slides.add_slide(presentation.slide_layouts[6])
                D.add_picture(slide, operation, str(root))
                picture = slide.shapes[0]
                self.assertEqual((picture.left, picture.top, picture.width, picture.height), (0, 500000, 2000000, 1000000))
                self.assertEqual((picture.crop_top, picture.crop_bottom), (0, 0))

    def test_layout_rejects_overlap_and_out_of_bounds(self):
        with patch.object(D, 'slides', {}):
            D.tb(0, 0, 0, 3000000, 'First', cy=500000)
            D.tb(0, 0, 200000, 3000000, 'Second', cy=500000)
            with self.assertRaisesRegex(ValueError, 'textbox overlap'):
                D.validate_layout()
        with patch.object(D, 'slides', {}):
            D.rect(0, D.W - 100000, 0, 200000, 500000)
            with self.assertRaisesRegex(ValueError, 'invalid bounds'):
                D.validate_layout()

    def test_cover_omits_duration_footer_and_renders_pages_qr_code(self):
        cover_texts = [op.get('text', '') for op in D.slides[0] if op['op'] == 'add_text_box']
        self.assertFalse(any('November 2026' in t or '45 minutes' in t for t in cover_texts))
        self.assertTrue(all(op.get('name') != 'footer_src' for ops in D.slides.values() for op in ops))
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            url = 'https://huangyingting.github.io/mxc/'
            with patch.object(D, 'slides', {0: [], 1: []}), patch.object(D, 'notes', {}):
                H.build(str(root / 'index.html'), str(root), 1600, pages_url=url)
            doc = (root / 'index.html').read_text(encoding='utf-8')
            self.assertEqual(doc.count('class="a pages-qr-card"'), 1)
            self.assertIn('href="https://huangyingting.github.io/mxc/"', doc)
            self.assertIn('huangyingting.github.io/mxc', doc)
            self.assertIn('<svg class="pages-qr-svg"', doc)

    @unittest.skipUnless(shutil.which('node'), 'node is required to syntax-check embedded JS')
    def test_embedded_javascript_has_valid_syntax(self):
        with tempfile.TemporaryDirectory() as directory:
            js_path = Path(directory) / 'deck.js'
            js_path.write_text(H.JS, encoding='utf-8')
            subprocess.run(['node', '--check', str(js_path)], check=True)


if __name__ == '__main__':
    unittest.main()
