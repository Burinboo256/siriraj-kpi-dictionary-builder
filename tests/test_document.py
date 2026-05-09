from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from docx import Document

from src.generate_document import generate_kpi_dictionary
from src.generate_template import create_template_workbook


class DocumentGenerationTests(unittest.TestCase):
    def test_generate_kpi_dictionary_creates_search_indexes_and_reference_style_detail_pages(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            template_path = Path(tmpdir) / "template.xlsx"
            output_path = Path(tmpdir) / "KPI_Dictionary.docx"
            report_path = Path(tmpdir) / "validation_report.xlsx"

            create_template_workbook(template_path)
            generate_kpi_dictionary(template_path, output_path, validation_report_path=report_path)

            self.assertTrue(output_path.exists())
            self.assertTrue(report_path.exists())

            document = Document(output_path)
            paragraph_text = "\n".join(p.text for p in document.paragraphs)

            self.assertIn("KPI Dictionary", paragraph_text)
            self.assertIn("Index by KPI Code", paragraph_text)
            self.assertIn("Index by KPI Category", paragraph_text)
            self.assertIn("DH0101", paragraph_text)
            self.assertGreaterEqual(len(document.tables), 6)

            detail_table = document.tables[-1]
            detail_labels = [detail_table.cell(row_index, 0).text.strip() for row_index in range(len(detail_table.rows))]
            detail_values = [detail_table.cell(row_index, 1).text.strip() for row_index in range(len(detail_table.rows))]

            self.assertIn("หมวดตัวชี้วัด", detail_labels)
            self.assertIn("ประเภทตัวชี้วัด", detail_labels)
            self.assertIn("รหัสตัวชี้วัด", detail_labels)
            self.assertIn("ชื่อตัวชี้วัด (ภาษาไทย)", detail_labels)
            self.assertIn("ข้อมูลที่ต้องการ", detail_labels)
            self.assertIn("รหัสโรค/หัตถการที่เกี่ยวข้อง", detail_labels)
            self.assertIn("PS0301", detail_values)


if __name__ == "__main__":
    unittest.main()
