import unittest

from checker.english_format_checker import check_english_format, check_item


class EnglishFormatCheckerTests(unittest.TestCase):
    def make_item(self, text, **overrides):
        item = {
            "file_name": "manual.docx",
            "heading": "Installation",
            "location": "正文-段落1",
            "text": text,
        }
        item.update(overrides)
        return item

    def test_powershell_relative_path_is_not_reported(self):
        text = r"Get-FileHash .\snap-k8s-offline.tar.gz -Algorithm MD5"
        self.assertEqual([], check_item(self.make_item(text)))

    def test_ip_placeholder_is_not_reported(self):
        text = 'Enter "XX.XX.XX.XX" in your browser and press Enter.'
        self.assertEqual([], check_item(self.make_item(text)))

    def test_url_version_and_ip_are_not_reported(self):
        text = "Open https://www.example.com for V2.5.0 on 100.7.32.233."
        self.assertEqual([], check_item(self.make_item(text)))

    def test_standalone_file_extension_is_not_reported(self):
        text = "Save the template in .xml format locally."
        self.assertEqual([], check_item(self.make_item(text)))

    def test_independent_command_is_not_reported(self):
        self.assertEqual([], check_item(self.make_item("sudo imcli status")))

    def test_code_style_is_not_reported(self):
        item = self.make_item("NAME        STATUS", is_code_style=True)
        self.assertEqual([], check_item(item))

    def test_real_spacing_errors_are_still_reported(self):
        issues = check_english_format("This  is wrong . Another sentence.Next one.")
        self.assertIn("发现连续两个或多个空格", issues)
        self.assertIn("英文标点前可能多了空格", issues)
        self.assertIn("英文标点后可能缺少空格", issues)

    def test_chinese_punctuation_in_prose_is_still_reported(self):
        issues = check_english_format("Run the command，and check the result.")
        self.assertIn("发现中文标点：，", issues)


if __name__ == "__main__":
    unittest.main()
