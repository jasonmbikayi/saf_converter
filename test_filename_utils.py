#!/usr/bin/env python3
import unittest
from filename_utils import clean_filename

class TestCleanFilename(unittest.TestCase):
    def test_diacritics(self):
        original, cleaned = clean_filename("Piò_Dead_Christ.pdf")
        self.assertEqual(cleaned, "pio_dead_christ.pdf")

    def test_spaces(self):
        original, cleaned = clean_filename("My File Name.DOCX")
        self.assertEqual(cleaned, "my_file_name.docx")

    def test_apostrophes(self):
        original, cleaned = clean_filename("Rome’s_File.txt")
        # Apostrophe → hyphen, underscore preserved
        self.assertEqual(cleaned, "rome-s_file.txt")

    def test_parentheses(self):
        original, cleaned = clean_filename("Report(2023).pdf")
        self.assertEqual(cleaned, "report_2023.pdf")

    def test_multiple_underscores(self):
        original, cleaned = clean_filename("___Draft__File___.jpg")
        self.assertEqual(cleaned, "draft_file.jpg")

    def test_strip_edges(self):
        original, cleaned = clean_filename("_-Weird-Name-_.png")
        self.assertEqual(cleaned, "weird-name.png")

    def test_mixed_versioning(self):
        original, cleaned = clean_filename("Résumé (Final) v2.0.pdf")
        # Diacritics removed, parentheses cleaned, space → underscore
        self.assertEqual(cleaned, "resume_final_v2.0.pdf")

if __name__ == "__main__":
    unittest.main()
