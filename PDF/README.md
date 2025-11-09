# PDF Operations 📄

VBA modules for exporting Excel worksheets to PDF format.

## SaveAsPDF.bas

Export Excel sheets to PDF with customizable print areas.

### Functions

#### 1. SaveActiveSheetAsPDF_SetPrintArea()

Export a single worksheet to PDF with a fixed print area.

**Features:**
- Set custom print area (e.g., $A$1:$G$50)
- Configure page orientation and scaling
- Auto-generate filename with timestamp
- Save in same folder as Excel file

**Usage:**
```vba
' In VBA Editor, run this macro
SaveActiveSheetAsPDF_SetPrintArea
```

**Customization:**
- Change `ws = wb.Worksheets("Sheet1")` to your sheet name
- Or use `Set ws = ActiveSheet` for current sheet
- Modify print area: `ws.PageSetup.PrintArea = "$A$1:$G$50"`
- Change orientation: `.Orientation = xlLandscape` or `xlPortrait`

---

#### 2. SaveWorkbookAsSinglePDF_SetPrintAreas()

Export entire workbook as a single PDF file with dynamic print areas.

**Features:**
- Automatically detect used range for each sheet
- Export all sheets to one PDF file
- Smart print area detection
- Configurable column limits

**Usage:**
```vba
' In VBA Editor, run this macro
SaveWorkbookAsSinglePDF_SetPrintAreas
```

**Customization:**
- To limit columns (e.g., A to G only):
  ```vba
  ws.PageSetup.PrintArea = ws.Range("A1", ws.Cells(ws.Rows.Count, "G").End(xlUp)).Address
  ```
- To use fixed range for all sheets:
  ```vba
  ws.PageSetup.PrintArea = "$A$1:$G$100"
  ```

---

## Common Modifications

### Change Output Folder
```vba
' Instead of using workbook folder:
pdfPath = "C:\MyReports\"
```

### Open PDF After Export
```vba
ws.ExportAsFixedFormat Type:=xlTypePDF, _
    Filename:=fullPath, _
    OpenAfterPublish:=True  ' Changed to True
```

### Change PDF Quality
```vba
Quality:=xlQualityMinimum  ' Smaller file size
' or
Quality:=xlQualityStandard ' Better quality (default)
```

### Custom Filename Format
```vba
' Add prefix/suffix
pdfName = "Report_" & ws.Name & "_" & Format(Now, "yyyymmdd") & ".pdf"
```

---

## Troubleshooting

**"กรุณาบันทึกไฟล์ Excel ก่อน" message:**
- Save your Excel file first (Ctrl+S) before running the macro
- The macro needs to know the file location to save PDF there

**Print area not correct:**
- Manually set print area: Page Layout → Print Area → Set Print Area
- Then check the range in VBA: `Debug.Print ws.PageSetup.PrintArea`

**PDF is blank:**
- Check if print area contains data
- Verify sheet is not hidden
- Ensure print area is not set to empty range

---

## Requirements

- Excel 2007 or later (for PDF export support)
- File must be saved as `.xlsm` (macro-enabled)
- Workbook must be saved before running (to determine output path)
