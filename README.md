# KPI Dictionary Generator

เครื่องมือ Python สำหรับสร้างเอกสาร KPI Dictionary แบบไทย/อังกฤษจากไฟล์ Excel โดยผลลัพธ์เป็นไฟล์ `.docx` พร้อมใช้งาน ซึ่งประกอบด้วย:

- หน้าปก
- สารบัญ
- ดัชนีค้นหาหลายรูปแบบ
- หน้า KPI detail แยกตามตัวชี้วัด
- footer พร้อมเลขหน้า
- รูปแบบเอกสารภาษาไทยเป็นหลัก และมีภาษาอังกฤษประกอบ
- ไฟล์ validation report

รูปแบบเอกสารถูกออกแบบให้ใกล้เคียงกับคู่มือ KPI ของโรงพยาบาล / เอกสารอ้างอิงลักษณะ THIP

## Repository

GitHub repository:

`https://github.com/Burinboo256/siriraj-kpi-dictionary-builder.git`

Clone โปรเจกต์ด้วยคำสั่ง:

```bash
git clone https://github.com/Burinboo256/siriraj-kpi-dictionary-builder.git
cd siriraj-kpi-dictionary-builder
```

## โครงสร้างโฟลเดอร์

```text
docx_kpi_dictionary_codex/
├── assets/
│   └── logo.png
├── input/
│   └── kpi_dictionary_template.example.xlsx
├── output/
│   ├── KPI_Dictionary.docx
│   ├── KPI_Dictionary.pdf
│   └── validation_report.xlsx
├── src/
│   ├── __init__.py
│   ├── generate_document.py
│   ├── generate_template.py
│   ├── styles.py
│   └── utils.py
├── temp/
├── tests/
│   ├── test_document.py
│   ├── test_template.py
│   └── test_utils.py
├── .gitignore
├── requirements.txt
└── README.md
```

## Excel Template

repo นี้เก็บไฟล์ตัวอย่างไว้ที่ `input/kpi_dictionary_template.example.xlsx`

ไฟล์ใช้งานจริงของคุณควรเป็น `input/kpi_dictionary_template.xlsx` ซึ่งถูก ignore จาก Git

เตรียมไฟล์ใช้งานจริงด้วยคำสั่ง:

```bash
cp input/kpi_dictionary_template.example.xlsx input/kpi_dictionary_template.xlsx
```

การใช้งานปกติ:

- อ่าน `01_Instruction`
- กรอกเฉพาะ `02_KPI_Input_Form`
- ใช้ `03_KPI_Input_Dictionary` เป็นคู่มืออธิบายแต่ละคอลัมน์ ว่าหมายถึงอะไร ช่องไหน required/optional มีตัวอย่างอะไร และมี allowed values อะไรบ้าง

ชีตที่เหลือเป็น internal hidden sheets ที่ตัว generator ใช้งานเอง และไม่จำเป็นต้องแก้ไขในการกรอกข้อมูลทั่วไป

ชีต input ถูกออกแบบตามหน้า KPI detail โดยตรง และเก็บเฉพาะข้อมูลที่จำเป็นต่อการสร้างเอกสาร:

- ช่อง required ถูก highlight ไว้สำหรับ category, type, code, ชื่อไทย/อังกฤษ, definition, objective, formula, numerator, denominator และ related code sets
- ช่อง optional ไม่ถูก highlight เช่น frequency, benchmark, interpretation, team/reference, data source, version dates, revision reason และ note
- มีชีต dictionary ประกอบสำหรับอธิบายคอลัมน์ทั้งหมดของชีต input
- มี dropdown validation สำหรับ category, type และ frequency ในชีตที่ผู้ใช้กรอก

ตัว generator ยังรองรับ Google Sheet รูปแบบเดิมที่มีชีต `KPIs` แบบ flat columns ได้ด้วย

## แนวทางการเก็บไฟล์ใน Git

ไฟล์ที่ควร track:

- `src/`
- `tests/`
- `assets/`
- `input/kpi_dictionary_template.example.xlsx`
- `requirements.txt`
- `README.md`

ไฟล์ที่ควรเป็น local/untracked:

- `input/kpi_dictionary_template.xlsx`
- `output/` สำหรับไฟล์ `.docx`, `.pdf` และ validation workbook ที่ generate แล้ว
- `temp/` สำหรับไฟล์ชั่วคราวหรือไฟล์ export จาก Google Sheet
- `.venv/`
- Python cache files
- ไฟล์ระบบของ macOS

## การติดตั้ง

repo นี้แนะนำให้ใช้ `uv` เป็นหลัก:

```bash
cd docx_kpi_dictionary_codex
uv venv
source .venv/bin/activate
uv pip install -r requirements.txt
```

ถ้าไม่มี `uv` สามารถใช้ Python virtual environment ปกติได้:

```bash
cd docx_kpi_dictionary_codex
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## เริ่มใช้งานเร็ว

สร้างไฟล์ทำงานจริงจากไฟล์ตัวอย่าง:

```bash
cp input/kpi_dictionary_template.example.xlsx input/kpi_dictionary_template.xlsx
```

จากนั้นสร้างเอกสาร Word ด้วยคำสั่ง:

```bash
python3 src/generate_document.py input/kpi_dictionary_template.xlsx --output output/KPI_Dictionary.docx --validation-report output/validation_report.xlsx
```

## การใช้งาน

ถ้าคุณกำลังพัฒนาโครงสร้าง template เอง และต้องการ regenerate ไฟล์ทำงานจากโค้ด:

```bash
python3 src/generate_template.py
```

สร้างเอกสาร Word:

```bash
python3 src/generate_document.py input/kpi_dictionary_template.xlsx --output output/KPI_Dictionary.docx --validation-report output/validation_report.xlsx
```

สร้างเอกสารจาก public Google Sheet URL หรือ sheet ID:

```bash
python3 src/generate_document.py --google-sheet "https://docs.google.com/spreadsheets/d/1s6LB0a_j4k-RhBTKAfZwQVDda25xzgmbGFDTLRC1LvM/edit?usp=sharing" --output output/KPI_Dictionary.docx --validation-report output/validation_report.xlsx
```

เมื่อใช้ `--google-sheet` ระบบจะ export workbook มาเก็บใน `temp/` เป็น `.xlsx` ก่อน แล้วจึงประมวลผลด้วย pipeline เดียวกัน

ถ้าต้องการลองสร้างทั้ง Word และ PDF:

```bash
python3 src/generate_document.py input/kpi_dictionary_template.xlsx --output output/KPI_Dictionary.docx --validation-report output/validation_report.xlsx --pdf
```

## ผลลัพธ์ที่ได้

ไฟล์ Word ที่สร้างจะประกอบด้วย:

- หน้าปกและสารบัญ
- ดัชนีค้นหาตาม KPI code
- ดัชนีค้นหาตาม KPI category
- ดัชนีค้นหาตาม KPI type
- ดัชนีค้นหาตาม frequency
- ดัชนีค้นหาตาม reference team
- ดัชนีค้นหาตาม ICD / procedure code
- หน้า KPI detail สำหรับแต่ละ KPI ตามรูปแบบ reference section
- แถวข้อมูล required เช่น category, type, code, ชื่อไทย/อังกฤษ, definition, objective, formula, numerator, denominator และ related code sets
- แถวข้อมูล optional เช่น frequency, benchmark, interpretation, team/reference, effective date, last update, revision reason และ note
- footer ที่มี version ของเอกสารและเลขหน้า

validation report จะตรวจเรื่อง:

- missing required fields
- duplicate KPI codes
- invalid dropdown values
- missing formula
- ICD format warning
- related KPI code not found

## หมายเหตุเรื่องภาษาไทย

- ค่า default ของเอกสารใช้ `TH Sarabun New` สำหรับภาษาไทย และ `Arial` สำหรับภาษาอังกฤษ
- label และ KPI name ภาษาไทยจะแสดงก่อน
- KPI name ภาษาอังกฤษจะแสดงเป็นตัวเอียงประกอบ
- เพื่อให้ไฟล์ `.docx` แสดงผลดีที่สุด ควรติดตั้ง font ตามที่กำหนดไว้ในเครื่องที่ใช้เปิดเอกสาร

## หมายเหตุเรื่องเลขหน้าและสารบัญ

- footer page number ถูกสร้างเป็น Word field
- footer จะแสดง version ของเอกสารด้วย
- สารบัญถูกแทรกเป็น Word TOC field
- หลังเปิดไฟล์ `.docx` ใน Microsoft Word ควร update fields เพื่อให้สารบัญและเลขหน้าตรงกับเอกสารล่าสุด
- เลขหน้าใน search indexes ถูกคำนวณตามลำดับเอกสารเป็นค่าเริ่มต้นเชิงปฏิบัติ

## การทดสอบ

รัน test suite ด้วยคำสั่ง:

```bash
python3 -m unittest discover -s tests -v
```

## ขั้นตอนก่อน commit

คำสั่งหลักที่ควรรันก่อน commit:

```bash
python3 -m unittest discover -s tests -v
python3 src/generate_template.py
python3 src/generate_document.py input/kpi_dictionary_template.xlsx --output output/KPI_Dictionary.docx --validation-report output/validation_report.xlsx
```

## รายละเอียดการทำงานภายใน

- `src/generate_template.py`: สร้าง Excel template แบบ KPI detail-first พร้อม required-field highlighting
- `src/utils.py`: จัดการ source ของ workbook, รวมข้อมูลจากชีตที่เกี่ยวข้อง และสร้าง validation findings
- `src/styles.py`: helper สำหรับ formatting เอกสาร เช่น border, shading, TOC และ page number
- `src/generate_document.py`: สร้างไฟล์ `.docx`, search indexes และ KPI detail tables

## หมายเหตุเรื่อง PDF

การ export PDF เป็น optional feature และพึ่งพา `docx2pdf` ซึ่งโดยทั่วไปต้องใช้ Microsoft Word บน Windows หรือ macOS ถ้า PDF export ใช้งานไม่ได้ ระบบยังคงสร้าง `.docx` ได้ตามปกติ
