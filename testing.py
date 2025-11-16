import unittest
import re

# Mock config with corrected patterns
MOCK_CONFIG = {
    "test_project": {
        "file_rename": {
            "block_number_in_filename": {
                "pattern": "re.compile(r'test_pattern')",
                "replacement": "'new_replacement'"
            },
            "starts_with_digit": {
                "pattern": "re.compile(r'test_pattern')",
                "replacement": "'new_replacement'"
            }
        }
    }
}

# Mock the _load_patterns to use MOCK_CONFIG
def mock_load_patterns(self, project, parser_name):
    patterns = []
    rules = MOCK_CONFIG.get(project, {}).get(parser_name, {})
    for rule_name, rule in rules.items():
        try:
            pattern = eval(rule["pattern"], {"re": re})
            replacement = eval(rule["replacement"], {"self": self})
            patterns.append((pattern, replacement))
        except Exception as e:
            pass
    return patterns

# Define the processor classes with mocked load_patterns
class MockExcelProcessor:
    def __init__(self, replacement_digit, project):
        self.replacement_digit = str(replacement_digit)
        self.patterns = mock_load_patterns(self, project, "excel_parser")

    def _apply_replacements(self, text):
        for pattern, repl in self.patterns:
            text = pattern.sub(repl, text)
        return text

class MockWordProcessor:
    def __init__(self, replacement_digit, project):
        self.replacement_digit = str(replacement_digit)
        self.patterns = mock_load_patterns(self, project, "word_parser")

    def _apply_replacements(self, text):
        for pattern, repl in self.patterns:
            text = pattern.sub(repl, text)
        return text

class MockAutoCADProcessor:
    def __init__(self, replacement_digit, project):
        self.replacement_digit = str(replacement_digit)
        self.patterns = mock_load_patterns(self, project, "dwg_parser")

    def _apply_replacements(self, text):
        for pattern, repl in self.patterns:
            text = pattern.sub(repl, text)
        return text

class MockShaProcessor:
    def __init__(self, replacement_digit, project):
        self.replacement_digit = str(replacement_digit)
        self.patterns = mock_load_patterns(self, project, "sha_parser")

    def _apply_replacements(self, text):
        for pattern, repl in self.patterns:
            text = pattern.sub(repl, text)
        return text

# For file_rename, define a function
def apply_file_rename(filename, replacement_digit, project):
    rules = MOCK_CONFIG.get(project, {}).get("file_rename", {})
    new_name = filename
    for rule in rules.values():
        pattern = eval(rule["pattern"], {"re": re})
        replacement = eval(rule["replacement"], {"replacement_digit": replacement_digit})
        new_name = pattern.sub(replacement, new_name)
    return new_name

# Specific test lists
EXCEL_TESTS = [
    ("old_text", "new_text"),
]

WORD_TESTS = [
    ("old_text", "new_text"),
]

DWG_TESTS = [
    ("old_text", "new_text"),
]

SHA_TESTS = [
    ("old_text", "new_text"),
    ]


FILENAME_TESTS = [
    ("old_text", "new_text"),
]

PROJECT = "test_project"
DIGIT = 'test_digit'

class TestProcessors(unittest.TestCase):

    def test_excel_processor(self):
        processor = MockExcelProcessor(DIGIT, PROJECT)
        for input_text, expected in EXCEL_TESTS:
            result = processor._apply_replacements(input_text)
            self.assertEqual(result, expected, f"Failed for {input_text} in excel")

    def test_word_processor(self):
        processor = MockWordProcessor(DIGIT, PROJECT)
        for input_text, expected in WORD_TESTS:
            result = processor._apply_replacements(input_text)
            self.assertEqual(result, expected, f"Failed for {input_text} in word")

    def test_dwg_processor(self):
        processor = MockAutoCADProcessor(DIGIT, PROJECT)
        for input_text, expected in DWG_TESTS:
            result = processor._apply_replacements(input_text)
            self.assertEqual(result, expected, f"Failed for {input_text} in dwg")

    def test_sha_processor(self):
        processor = MockShaProcessor(DIGIT, PROJECT)
        for input_text, expected in SHA_TESTS:
            result = processor._apply_replacements(input_text)
            self.assertEqual(result, expected, f"Failed for {input_text} in sha")

    def test_file_rename(self):
        for input_name, expected in FILENAME_TESTS:
            result = apply_file_rename(input_name, DIGIT, PROJECT)
            self.assertEqual(result, expected, f"Failed for {input_name}")

if __name__ == '__main__':
    unittest.main()