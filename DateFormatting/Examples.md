# ตัวอย่างการใช้งาน DateFormatting Module 📖

## English Summary

This document provides examples for using the DateFormatting module functions. Each example includes VBA code, expected input/output, and use cases.

---

## Function 1: TextToDate()

### ตัวอย่างที่ 1.1 - ใช้ใน VBA
```vba
Sub Example_TextToDate_VBA()
    Dim myDate As Date
    
    ' แปลง text เป็น date
    myDate = TextToDate("27/11/2025")
    
    ' แสดงผล
    Debug.Print myDate          ' 27/11/2025 (เป็น Date value จริง)
    Debug.Print myDate + 7      ' 04/12/2025 (บวกได้ 7 วัน)
End Sub
```

### ตัวอย่างที่ 1.2 - ใช้เป็น UDF ใน Cell
```
' สมมติ A1 มีข้อความ "27/11/2025"

' ใน B1 พิมพ์:
=TextToDate(A1)

' ผลลัพธ์: 27/11/2025 (เป็น Date value ที่คำนวณได้)
```

### ตัวอย่างที่ 1.3 - แปลงแล้วนำไปคำนวณ
```vba
Sub Example_TextToDate_Calculate()
    Dim startDate As Date
    Dim endDate As Date
    Dim daysDiff As Long
    
    startDate = TextToDate("01/11/2025")
    endDate = TextToDate("27/11/2025")
    
    daysDiff = endDate - startDate
    Debug.Print "จำนวนวัน: " & daysDiff    ' จำนวนวัน: 26
End Sub
```

---

## Function 2: InsertFormattedDate()

### ตัวอย่างที่ 2.1 - ใส่วันที่ปัจจุบัน
```vba
Sub Example_InsertDate()
    ' ใส่วันที่วันนี้ในเซลล์ A1
    InsertFormattedDate Range("A1")
    
    ' ผลลัพธ์ใน A1:
    ' - ค่า: 27/11/2025 (Date value)
    ' - Format: [$-409]dd/mm/yyyy
End Sub
```

### ตัวอย่างที่ 2.2 - ใส่วันที่พร้อมเวลา
```vba
Sub Example_InsertDateTime()
    ' ใส่วันที่และเวลาในเซลล์ A1
    InsertFormattedDate Range("A1"), True
    
    ' ผลลัพธ์ใน A1:
    ' - ค่า: 27/11/2025 14:30:22 (DateTime value)
    ' - แสดงผล: 27/11/2025 (ตาม format ที่กำหนด)
End Sub
```

### ตัวอย่างที่ 2.3 - ใส่วันที่หลายเซลล์
```vba
Sub Example_InsertMultipleDates()
    Dim i As Long
    
    ' ใส่วันที่ในคอลัมน์ A, แถว 1-10
    For i = 1 To 10
        InsertFormattedDate Cells(i, 1)
    Next i
End Sub
```

### ตัวอย่างที่ 2.4 - ใช้กับ UserForm
```vba
Private Sub btnInsertDate_Click()
    ' เมื่อกดปุ่มให้ใส่วันที่ในเซลล์ที่เลือก
    If Not TypeOf Selection Is Range Then Exit Sub
    InsertFormattedDate Selection.Cells(1, 1)
End Sub
```

---

## Function 3: ConvertRangeTextToDate()

### ตัวอย่างที่ 3.1 - แปลงข้อมูลใน Range
```vba
Sub Example_ConvertRange()
    ' สมมติ A1:A10 มีข้อมูล text date
    ' เช่น "27/11/2025", "28/11/2025", ...
    
    ConvertRangeTextToDate Range("A1:A10")
    
    ' ผลลัพธ์:
    ' - ข้อมูลทั้งหมดแปลงเป็น Date value
    ' - Format เป็น dd/mm/yyyy
    ' - แสดง MsgBox บอกจำนวนที่แปลงสำเร็จ
End Sub
```

### ตัวอย่างที่ 3.2 - แปลงเซลล์ที่เลือก
```vba
Sub Example_ConvertSelection()
    ' ให้ผู้ใช้เลือก range แล้วแปลง
    ConvertRangeTextToDate Selection
End Sub
```

### ตัวอย่างที่ 3.3 - แปลงทั้งคอลัมน์ (เฉพาะที่มีข้อมูล)
```vba
Sub Example_ConvertColumn()
    Dim lastRow As Long
    
    ' หา row สุดท้ายที่มีข้อมูล
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' แปลงเฉพาะส่วนที่มีข้อมูล
    ConvertRangeTextToDate Range("A1:A" & lastRow)
End Sub
```

---

## Function 4: GetFormattedDateString()

### ตัวอย่างที่ 4.1 - สร้างชื่อไฟล์
```vba
Sub Example_CreateFilename()
    Dim fileName As String
    
    ' สร้างชื่อไฟล์พร้อม timestamp
    fileName = "Report_" & GetFormattedDateString(Now, "yyyymmdd_hhmmss") & ".xlsx"
    
    Debug.Print fileName    ' Report_20251127_143022.xlsx
End Sub
```

### ตัวอย่างที่ 4.2 - แสดงวันที่ในรูปแบบต่างๆ
```vba
Sub Example_DateFormats()
    Dim today As Date
    today = Date
    
    Debug.Print GetFormattedDateString(today)                     ' 27/11/2025
    Debug.Print GetFormattedDateString(today, "dd-mm-yyyy")       ' 27-11-2025
    Debug.Print GetFormattedDateString(today, "yyyy-mm-dd")       ' 2025-11-27
    Debug.Print GetFormattedDateString(today, "dd mmm yyyy")      ' 27 Nov 2025
    Debug.Print GetFormattedDateString(today, "mmmm dd, yyyy")    ' November 27, 2025
End Sub
```

---

## Use Case: แก้ปัญหา Date Format ต่างกันระหว่าง PC

### สถานการณ์
- PC เครื่อง A ใช้ format dd/mm/yyyy
- PC เครื่อง B ใช้ format mm/dd/yyyy
- ต้องการให้แสดงผลเหมือนกันทุกเครื่อง

### วิธีแก้
```vba
Sub FixDateFormatProblem()
    ' แทนที่จะใช้:
    ' Range("A1").Value = WorksheetFunction.Text(Now(), "dd/mm/yyyy")
    ' ซึ่งได้ผลเป็น TEXT ไม่ใช่ DATE
    
    ' ให้ใช้:
    InsertFormattedDate Range("A1")
    
    ' ซึ่งจะ:
    ' 1. ใส่ค่าเป็น Date value จริง
    ' 2. กำหนด format [$-409]dd/mm/yyyy
    ' 3. แสดงผลเหมือนกันทุกเครื่อง
    ' 4. สามารถคำนวณและเรียงลำดับได้
End Sub
```

---

## Use Case: Import ข้อมูลจาก Text File

### สถานการณ์
- Import ข้อมูลจาก CSV ที่มีคอลัมน์วันที่
- วันที่ถูก import เป็น text

### วิธีแก้
```vba
Sub AfterImportCSV()
    ' หลังจาก import เสร็จแล้ว
    ' แปลงคอลัมน์วันที่จาก text เป็น date
    
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' สมมติคอลัมน์ B เป็นวันที่
    ConvertRangeTextToDate Range("B2:B" & lastRow)
End Sub
```

---

## เปรียบเทียบผลลัพธ์

### ก่อนใช้ Module (วิธีเดิม)
| เซลล์ | ค่า | ประเภท | ปัญหา |
|-------|-----|--------|-------|
| A1 | 27/11/2025 | Text | ❌ เรียงลำดับผิด, คำนวณไม่ได้ |
| A2 | 28/11/2025 | Text | ❌ เรียงลำดับผิด, คำนวณไม่ได้ |

### หลังใช้ Module (วิธีใหม่)
| เซลล์ | ค่า | ประเภท | ผลลัพธ์ |
|-------|-----|--------|---------|
| A1 | 45988 | Date (แสดง 27/11/2025) | ✅ เรียงลำดับถูก, คำนวณได้ |
| A2 | 45989 | Date (แสดง 28/11/2025) | ✅ เรียงลำดับถูก, คำนวณได้ |

---

## Tips & Tricks

### Tip 1: ตรวจสอบว่าเป็น Date หรือ Text
```vba
Sub CheckIfDate()
    If IsDate(Range("A1").Value) And VarType(Range("A1").Value) = vbDate Then
        Debug.Print "เป็น Date value จริง"
    Else
        Debug.Print "เป็น Text หรือค่าอื่น"
    End If
End Sub
```

### Tip 2: กำหนด Keyboard Shortcut
1. กด Alt + F8
2. เลือก Macro ที่ต้องการ
3. กด Options
4. กำหนด Shortcut key

### Tip 3: เพิ่มปุ่มใน Quick Access Toolbar
1. File → Options → Quick Access Toolbar
2. Choose commands from: Macros
3. เลือก Macro แล้วกด Add
