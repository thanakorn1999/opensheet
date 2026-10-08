# Cross-Platform XLSX Editor

## 1. เป้าหมาย

สร้างโปรแกรมสำหรับเปิด อ่าน แก้ไข และบันทึกไฟล์ `.xlsx` รองรับ:

* macOS
* Windows

โปรแกรมควรสามารถทำงานกับไฟล์ Excel ได้โดยไม่จำเป็นต้องติดตั้ง Microsoft Excel

---

# 2. Technology Stack

## ภาษาและ Framework

### C# + .NET 10 (LTS) + Avalonia UI

รองรับ:

```text
macOS
Windows
```

ข้อดี:

* ใช้ภาษาเดียว (C#) ทั้ง UI และ XLSX engine — ไม่ต้องมี FFI / Bridge
* Avalonia วาด UI เอง (Skia) หน้าตาเหมือนกันทั้ง macOS และ Windows
* รองรับ keyboard / mouse / shortcut บน Desktop เต็มรูปแบบ
* ใช้ MVVM + data binding เหมาะกับ application แบบ editor
* License: Avalonia = MIT

---

# 3. XLSX Engine

## แนะนำ: OpenXML SDK + ClosedXML

| Library | หน้าที่ | License |
| --- | --- | --- |
| **DocumentFormat.OpenXml** (OpenXML SDK) | อ่าน/เขียนโครงสร้าง OOXML ระดับล่าง — library ทางการของ Microsoft | MIT |
| **ClosedXML** | API ระดับสูงบน OpenXML SDK (Cell, Style, Merge, Freeze) + คำนวณ Formula | MIT |

แนวทาง:

* ใช้ **ClosedXML** เป็นหลัก — งานทั่วไปเขียนสั้นกว่ามาก
* ลงไปใช้ **OpenXML SDK** ตรงๆ เฉพาะส่วนที่ ClosedXML ไม่รองรับ
* ข้อจำกัด: ClosedXML โหลดทั้ง workbook เข้า memory — ไฟล์ใหญ่มาก (หลักแสนแถวขึ้นไป) ให้ใช้ `OpenXmlReader` (SAX) อ่านแบบ streaming สำหรับโหมด Viewer

Architecture:

```text
Avalonia UI (Views + ViewModels)
    │
    │ เรียก method ตรงๆ (project reference)
    ▼
XlsxEditor.Core (C# class library)
    │
    ├── Open XLSX
    ├── Read Cell
    ├── Write Cell
    ├── Sheet
    ├── Row / Column
    ├── Formula
    ├── Style
    ├── Merge Cell
    └── Save XLSX
    │
    ▼
ClosedXML → OpenXML SDK
```

เหตุผลที่แยก XLSX Engine เป็น project ของตัวเอง:

* ไม่ผูก logic กับ UI
* ทดสอบ engine แยกจาก UI ได้ (xUnit)
* เปลี่ยน/เพิ่ม UI ภายหลังได้โดยไม่แตะ engine

---

# 4. Platform Architecture

```text
                    ┌───────────────┐
                    │  Avalonia UI  │
                    │  (C# / XAML)  │
                    └───────┬───────┘
                            │
                     same process
                            │
                    ┌───────▼───────┐
                    │ XlsxEditor    │
                    │    .Core      │
                    └───────┬───────┘
                            │
                ClosedXML / OpenXML SDK
                            │
                       XLSX File
                            │
                    ┌───────┴───────┐
                    ▼               ▼
                  macOS          Windows
```

---

# 5. Core Features

## File

* Open `.xlsx`
* Save `.xlsx`
* Save As
* Create new workbook
* Import workbook
* Export workbook

---

## Workbook

รองรับ:

* Multiple worksheets
* Add worksheet
* Delete worksheet
* Rename worksheet
* Reorder worksheet
* Hide worksheet

ตัวอย่าง:

```text
Workbook
│
├── Sheet1
├── Sheet2
├── Sales
└── Users
```

---

# 6. Cell Editor

รองรับการแก้ไข:

```text
A1
A2
B1
B2
...
```

ประเภทข้อมูล:

* String
* Integer
* Decimal
* Boolean
* Date
* Time
* Formula
* Empty

ตัวอย่าง:

```text
A1 = "Name"
B1 = "Age"

A2 = "John"
B2 = 25
```

---

# 7. Formula

ควรรองรับ Formula พื้นฐานก่อน

```text
=SUM(A1:A10)

=AVERAGE(B1:B10)

=COUNT(A1:A10)

=MAX(A1:A10)

=MIN(A1:A10)

=A1+B1

=A1*B1
```

ไม่ควรพยายามทำ Formula Engine ครบทุกอย่างใน Version แรก

---

# 8. Cell Formatting

รองรับ:

* Font
* Font size
* Bold
* Italic
* Underline
* Text alignment
* Background
* Text color
* Border
* Number format
* Date format
* Decimal format

ตัวอย่าง:

```text
┌──────────┬───────┬────────┐
│ Name     │ Age   │ Salary │
├──────────┼───────┼────────┤
│ John     │ 25    │ 30,000 │
│ Jane     │ 28    │ 40,000 │
└──────────┴───────┴────────┘
```

---

# 9. Row / Column

รองรับ:

* Add row
* Delete row
* Add column
* Delete column
* Resize column
* Resize row
* Hide row
* Hide column
* Freeze row
* Freeze column

---

# 10. Clipboard

```text
Ctrl/Cmd + C
Ctrl/Cmd + X
Ctrl/Cmd + V
```

ควรรองรับการ copy หลาย cells:

```text
A1:B5
```

---

# 11. Search

รองรับ:

```text
Search
Find
Find Next
Replace
Replace All
```

ตัวอย่าง:

```text
Search: "John"

A2 = John
A10 = John
B20 = John
```

---

# 12. UI

ใช้ UI แบบ Spreadsheet เต็มรูปแบบ:

```text
┌──────────────────────────────────────────────┐
│ File Edit View Insert Format                 │
├──────────────────────────────────────────────┤
│ A1 │ =SUM(A1:A10)                            │
├────┬────────┬────────┬────────┬──────────────┤
│    │ A      │ B      │ C      │ D            │
├────┼────────┼────────┼────────┼──────────────┤
│ 1  │        │        │        │              │
│ 2  │        │        │        │              │
│ 3  │        │        │        │              │
│ 4  │        │        │        │              │
└────┴────────┴────────┴────────┴──────────────┘

[Sheet1] [Sheet2] [+]
```

## Grid Control

`DataGrid` ของ Avalonia ออกแบบมาสำหรับตารางข้อมูล ไม่ใช่ spreadsheet (ไม่มี cell selection แบบช่วง, merge, freeze)

ควรเขียน control เอง: `SheetGrid : Control`

* override `Render(DrawingContext)` วาดเฉพาะ cell ที่มองเห็น (virtualization)
* วาด header แถว/คอลัมน์, เส้นตาราง, selection เอง
* ใช้ `TextBox` ลอยทับ cell เฉพาะตอนกำลังแก้ไข
* รองรับไฟล์หลายแสนแถวได้เพราะไม่สร้าง control ต่อ cell

---

# 13. File Storage

เปิดไฟล์จาก:

```text
Finder
Windows Explorer
Drag & Drop
```

เช่น:

```text
Open With → My XLSX Editor
```

---

# 14. Offline First

โปรแกรมควรสามารถ:

```text
Open XLSX
    ↓
Edit
    ↓
Save
```

โดยไม่ต้องมี Internet

ไม่ควรบังคับให้ Upload ไฟล์ขึ้น Server

ข้อดี:

* Privacy
* ใช้งาน Offline
* เร็ว
* ไม่มีค่า Server
* เหมาะกับเอกสารบริษัท

---

# 15. Security

ไฟล์ `.xlsx` ควรประมวลผลบนเครื่องเป็นหลัก

```text
User File
    ↓
Device
    ↓
XLSX Engine
    ↓
Device
```

ไม่จำเป็นต้องส่งไฟล์ขึ้น Server

---

# 16. Project Structure

```text
xlsx-editor/
│
├── src/
│   ├── XlsxEditor.App/            ← Avalonia app
│   │   ├── Views/
│   │   ├── ViewModels/
│   │   ├── Controls/
│   │   │   └── SheetGrid.cs
│   │   ├── Services/
│   │   └── XlsxEditor.App.csproj
│   │
│   └── XlsxEditor.Core/           ← XLSX engine (ไม่มี reference ไป UI)
│       ├── Workbook/
│       ├── Worksheet/
│       ├── Cell/
│       ├── Formula/
│       ├── Style/
│       └── XlsxEditor.Core.csproj
│
├── tests/
│   └── XlsxEditor.Core.Tests/     ← xUnit
│
├── docs/
│
├── XlsxEditor.sln
└── README.md
```

---

# 17. Alternatives ที่พิจารณาแล้ว

| Stack | เหตุผลที่ไม่เลือก |
| --- | --- |
| Flutter + Rust | ข้อดีหลักคือรองรับ Mobile ซึ่งตัดออกแล้ว, ต้องมี FFI bridge, Rust XLSX library ยังไม่ครบเท่า |
| Go + Wails + excelize | excelize ดีมาก แต่ต้องแยกภาษา UI (TypeScript) กับ engine (Go) |
| Tauri + Rust | แอปเล็ก/เร็ว แต่ library XLSX ฝั่ง Rust แยกอ่าน/เขียนคนละตัว |

C# + OpenXML SDK ได้ความเข้ากันได้กับไฟล์ Excel สูงสุด เพราะเป็น library ของ Microsoft เอง

---

# 18. Recommended Technology

| ส่วน        | Technology                          |
| ----------- | ----------------------------------- |
| UI          | Avalonia UI                         |
| Language    | C#                                  |
| Runtime     | .NET 10 (LTS)                       |
| XLSX Core   | ClosedXML + OpenXML SDK             |
| XLSX Format | OOXML                               |
| Platform    | macOS + Windows                     |
| Storage     | Local filesystem                    |
| State       | MVVM — CommunityToolkit.Mvvm        |
| Testing     | xUnit                               |
| Packaging   | macOS `.app` (notarize) + Windows MSIX / installer |
| CI/CD       | GitHub Actions                      |

---

# 19. Development Phases

## Phase 1 — Basic Viewer

```text
Open XLSX
↓
Show Sheets
↓
Show Cells
```

ยังไม่ต้องแก้ไข

---

## Phase 2 — Editor

เพิ่ม:

```text
Edit Cell
Add Row
Delete Row
Add Column
Delete Column
Save
```

---

## Phase 3 — Formatting

เพิ่ม:

```text
Bold
Italic
Font
Alignment
Border
Background
Number Format
```

---

## Phase 4 — Formula

เพิ่ม:

```text
SUM
AVERAGE
COUNT
MIN
MAX
Basic arithmetic
```

---

## Phase 5 — Advanced

เพิ่ม:

```text
Merge
Freeze
Filter
Sort
Conditional Formatting
Charts
Images
```

---

# 20. Important Limitation

อย่าตั้งเป้าว่า Version แรกจะต้องเป็น:

> "Microsoft Excel ที่เขียนขึ้นมาใหม่"

เพราะ Excel มีระบบจำนวนมหาศาล

เป้าหมายที่เหมาะสมกว่า:

> "Fast cross-platform XLSX viewer/editor"

แล้วค่อยเพิ่มความสามารถทีละส่วน

---

# 21. Final Recommendation

ถ้าต้องเลือก Stack เดียว:

```text
C# / .NET 10
   +
Avalonia UI
   +
ClosedXML / OpenXML SDK
```

สำหรับ:

```text
                 ┌────────────┐
                 │  Avalonia  │
                 └─────┬──────┘
                       │
                ┌──────▼──────┐
                │ C# .NET     │
                │ XLSX Core   │
                └──────┬──────┘
                       │
                  ┌────▼────┐
                  │  .xlsx  │
                  └─────────┘

                ┌────────┬────────┐
                │ macOS  │Windows │
                └────────┴────────┘
```

**Avalonia = UI**

**C# + ClosedXML / OpenXML SDK = XLSX engine**

**`.xlsx` = File format**

ไม่จำเป็นต้องมี Microsoft Excel ติดตั้งอยู่บนเครื่องผู้ใช้
