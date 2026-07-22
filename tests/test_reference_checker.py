import unittest

from checker.reference_checker import check_figure_references


class FigureReferenceCheckerTests(unittest.TestCase):
    def make_item(self, text, location, style=""):
        return {
            "file_name": "manual.docx",
            "heading": "1 Overview",
            "section_key": "1 Overview",
            "location": location,
            "text": text,
            "xml_text": text,
            "paragraph_style": style,
            "is_toc": False,
        }

    def test_caption_without_body_reference_is_a_hint(self):
        issues = check_figure_references([
            self.make_item("Figure 1-1 Planning Process", "正文-段落1", "Caption"),
        ])
        self.assertEqual("引用提示", issues[0]["issue_type"])
        self.assertIn("未检测到正文中的显式编号引用", issues[0]["issue"])

    def test_body_reference_without_caption_is_an_issue(self):
        issues = check_figure_references([
            self.make_item("The process is shown in Figure 1-1.", "正文-段落1"),
        ])
        self.assertEqual("引用问题", issues[0]["issue_type"])
        self.assertIn("未检测到对应图题", issues[0]["issue"])

    def test_caption_and_body_reference_pass(self):
        issues = check_figure_references([
            self.make_item("The process is shown in Figure 1-1.", "正文-段落1"),
            self.make_item("Figure 1-1 Planning Process", "正文-段落2", "Caption"),
        ])
        self.assertEqual([], issues)

    def test_similar_numbers_do_not_collide(self):
        issues = check_figure_references([
            self.make_item("Figure 1-23 First Figure", "正文-段落1", "Caption"),
            self.make_item("Figure 12-3 Second Figure", "正文-段落2", "Caption"),
        ])
        self.assertEqual(2, len(issues))


if __name__ == "__main__":
    unittest.main()
