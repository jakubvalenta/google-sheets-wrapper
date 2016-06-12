#!/usr/bin/env python

import unittest

from google_sheets_wrapper import sheets


class Test(unittest.TestCase):

    def test_a1(self):
        self.assertEqual(sheets.a1(0, 0), 'A1')
        self.assertEqual(sheets.a1(1, 1), 'B2')
        self.assertEqual(sheets.a1(10, 10), 'K11')
        self.assertEqual(sheets.a1(25, 25), 'Z26')
        self.assertEqual(sheets.a1(26, 26), 'AA27')
        self.assertEqual(sheets.a1(36, 36), 'KK37')

    def test_format_formula_image(self):
        self.assertEqual(
            sheets.format_formula_image('http://www.example.com/example.jpg'),
            '=IMAGE("http://www.example.com/example.jpg")'
        )

    def test_parse_formula_image(self):
        self.assertEqual(
            sheets.parse_formula_image(
                '=IMAGE("http://www.example.com/example.jpg")'
            ),
            'http://www.example.com/example.jpg'
        )


if __name__ == '__main__':
    unittest.main()
