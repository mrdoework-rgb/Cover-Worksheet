import unittest

from docx import Document

from app import append_science_text


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


if __name__ == "__main__":
    unittest.main()
