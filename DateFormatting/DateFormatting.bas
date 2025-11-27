Attribute VB_Name = "DateFormatting"
Option Explicit

' =====================================================
' DateFormatting Module
' โมดูลสำหรับจัดการ Date Format ใน Excel VBA
' แก้ปัญหา format วันที่ต่างกันระหว่าง PC
' =====================================================

' -----------------------------------------------------
' Function 1: TextToDate
' แปลง text date เป็น date value
' สามารถใช้เป็น UDF (User Defined Function) ใน cell ได้
'
' Parameters:
'   dateText - ข้อความวันที่ที่ต้องการแปลง (เช่น "27/11/2025")
'
' Returns:
'   Date value ที่สามารถใช้คำนวณได้
'
' ตัวอย่างการใช้งาน:
'   ใน VBA: result = TextToDate("27/11/2025")
'   ใน Cell: =TextToDate(A1)
'
' หมายเหตุ:
'   ถ้าแปลงไม่สำเร็จ จะ return #VALUE! error (เมื่อใช้เป็น UDF)
'   หรือ raise error เมื่อเรียกจาก VBA
' -----------------------------------------------------
Public Function TextToDate(dateText As String) As Date
    ' แปลง text เป็น date value
    ' ถ้า dateText ไม่ใช่รูปแบบวันที่ที่ถูกต้อง จะเกิด error
    TextToDate = DateValue(dateText)
End Function

' -----------------------------------------------------
' Function 2: InsertFormattedDate
' ลงวันที่พร้อม format ที่ถูกต้องในเซลล์ที่ระบุ
'
' ข้อดี:
'   - ใส่ค่าเป็น Date value จริง (ไม่ใช่ text)
'   - กำหนด format แยก ทำให้สามารถคำนวณและเรียงลำดับได้
'   - ใช้ [$-409] เพื่อบังคับ format แบบ English
'
' Parameters:
'   targetCell  - เซลล์ที่ต้องการใส่วันที่
'   includeTime - (Optional) ถ้า True จะใส่เวลาด้วย, Default = False
'
' ตัวอย่างการใช้งาน:
'   InsertFormattedDate Range("A1")              ' ใส่วันที่อย่างเดียว
'   InsertFormattedDate Range("A1"), True        ' ใส่วันที่พร้อมเวลา
' -----------------------------------------------------
Public Sub InsertFormattedDate(targetCell As Range, Optional includeTime As Boolean = False)
    On Error GoTo ErrHandler
    
    With targetCell
        ' ใส่ค่าเป็น Date value จริง
        If includeTime Then
            .Value = Now()
        Else
            .Value = Date
        End If
        
        ' กำหนด format โดยใช้ [$-409] เพื่อบังคับ English format
        ' ป้องกันปัญหา format ต่างกันระหว่าง PC
        .NumberFormat = "[$-409]dd/mm/yyyy"
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' -----------------------------------------------------
' Function 3: ConvertRangeTextToDate
' แปลงช่วงข้อมูลที่เป็น text date ให้เป็น date value
'
' ใช้กรณีที่มีข้อมูลเดิมเป็น text อยู่แล้ว
' และต้องการแปลงให้เป็น date จริงทั้งหมด
'
' Parameters:
'   targetRange - ช่วงเซลล์ที่ต้องการแปลง
'
' ตัวอย่างการใช้งาน:
'   ConvertRangeTextToDate Range("A1:A100")
'   ConvertRangeTextToDate Selection
' -----------------------------------------------------
Public Sub ConvertRangeTextToDate(targetRange As Range)
    On Error GoTo ErrHandler
    
    Dim cell As Range
    Dim convertedCount As Long
    
    Application.ScreenUpdating = False
    convertedCount = 0
    
    For Each cell In targetRange
        ' ตรวจสอบว่าเซลล์มีข้อมูลและเป็น text
        If Not IsEmpty(cell.Value) And cell.Value <> "" Then
            ' แปลงเฉพาะเซลล์ที่เป็น text (ไม่ใช่ Date value อยู่แล้ว)
            If VarType(cell.Value) = vbString Then
                ' พยายามแปลงเป็น date
                On Error Resume Next
                Err.Clear
                cell.Value = DateValue(cell.Value)
                If Err.Number = 0 Then
                    cell.NumberFormat = "dd/mm/yyyy"
                    convertedCount = convertedCount + 1
                End If
                Err.Clear
                On Error GoTo ErrHandler
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "แปลงข้อมูลสำเร็จ " & convertedCount & " เซลล์", vbInformation
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' -----------------------------------------------------
' Function 4: GetFormattedDateString
' สร้าง date string ที่มี format คงที่
' ใช้สำหรับ filename หรือ text ที่ต้องการความสม่ำเสมอ
'
' Parameters:
'   dateValue    - วันที่ที่ต้องการ format
'   formatString - (Optional) รูปแบบที่ต้องการ, Default = "dd/mm/yyyy"
'
' Returns:
'   String ของวันที่ในรูปแบบที่กำหนด
'
' ตัวอย่างการใช้งาน:
'   result = GetFormattedDateString(Date)                    ' "27/11/2025"
'   result = GetFormattedDateString(Now, "yyyymmdd_hhmmss")  ' "20251127_143022"
' -----------------------------------------------------
Public Function GetFormattedDateString(dateValue As Date, Optional formatString As String = "dd/mm/yyyy") As String
    On Error GoTo ErrHandler
    
    GetFormattedDateString = Format(dateValue, formatString)
    Exit Function
    
ErrHandler:
    GetFormattedDateString = ""
End Function
