"""Validate the committed example outputs of ``read_mdd``.

The files in ``tests/fixtures/*.mdd.json`` are real outputs of the MDD reader. They
cannot be regenerated here (reading an MDD needs Windows + IBM/Unicom Professional),
so these tests pin the report shape to keep the examples valid inputs for the
reporters and the diff tool.
"""

import json
import unittest
from pathlib import Path

FIXTURES = Path(__file__).resolve().parent / 'fixtures'

# Each one was produced by: python src/read_mdd.py --mdd <file>.mdd
READER_OUTPUTS = [
    'a1_v1.mdd.json',
    'a1_v2.mdd.json',
    'b1_today.mdd.json',
    'b1_tomorrow.mdd.json',
]


def _load(name):
    with (FIXTURES / name).open(encoding='utf-8') as f:
        return json.load(f)


class TestReaderExampleOutputs(unittest.TestCase):
    def test_every_reader_output_is_present(self):
        for name in READER_OUTPUTS:
            self.assertTrue((FIXTURES / name).is_file(), name)

    def test_reader_outputs_are_valid_json_reports(self):
        for name in READER_OUTPUTS:
            with self.subTest(name=name):
                report = _load(name)
                self.assertEqual(report.get('report_type'), 'MDD')
                self.assertIn('report_scheme', report)
                self.assertIn('columns', report['report_scheme'])
                self.assertIsInstance(report['report_scheme']['columns'], list)
                self.assertIn('sections', report)
                self.assertTrue(report['sections'], 'expected at least one section')
                for section in report['sections']:
                    self.assertIn('name', section)
                    self.assertIn('content', section)
                    self.assertIsInstance(section['content'], list)


if __name__ == '__main__':
    unittest.main()
