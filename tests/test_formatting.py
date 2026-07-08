import sys
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import append_science_text, render_section_label


class AppendScienceTextTests(unittest.TestCase):
    def test_parses_bold_and_subscript_tags(self):
        doc = Document()
        paragraph = doc.add_paragraph()

        append_science_text(paragraph, "Use **bold** and H<sub>2</sub>O")

        self.assertEqual(paragraph.runs[0].text, "Use ")
        self.assertFalse(paragraph.runs[0].font.bold)

        self.assertEqual(paragraph.runs[1].text, "bold")
        self.assertTrue(paragraph.runs[1].font.bold)

        self.assertEqual(paragraph.runs[2].text, " and H")
        self.assertFalse(paragraph.runs[2].font.bold)

        self.assertEqual(paragraph.runs[3].text, "2")
        self.assertTrue(paragraph.runs[3].font.subscript)

        self.assertEqual(paragraph.runs[4].text, "O")
        self.assertFalse(paragraph.runs[4].font.subscript)

    def test_render_section_label_uses_strong_html(self):
        with patch("app.st.markdown") as mock_markdown:
            render_section_label("Year Group")

        mock_markdown.assert_called_once_with("<strong>Year Group</strong>", unsafe_allow_html=True)


if __name__ == "__main__":
    unittest.main()
