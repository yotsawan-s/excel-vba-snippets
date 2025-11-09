Option Explicit

' ตัวอย่าง 1:
' ตั้ง PrintArea ให้ชีทที่ต้องการ (แก้ "Sheet1" หรือใช้ ActiveSheet)
' แล้วบันทึกชีทนั้นเป็น PDF ในโฟลเดอร์เดียวกับ ThisWorkbook
Public Sub SaveActiveSheetAsPDF_SetPrintArea()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pdfPath As String, pdfName As String, fullPath As String

    ' ถ้าคุณต้องการใช้โฟลเดอร์ของไฟล์ที่เปิดอยู่ ให้เปลี่ยน ThisWorkbook เป็น ActiveWorkbook
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Sheet1") ' <-- เปลี่ยนเป็นชื่อชีทที่ต้องการ หรือใช้: Set ws = ActiveSheet

    ' ตัวอย่างกำหนด Print Area แบบคงที่ (แก้ช่วงได้)
    ws.PageSetup.PrintArea = "$A$1:$G$50"

    ' ตั้งค่าหน้ากระดาษเพิ่มเติม (ปรับตามต้องการ)
    With ws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

    pdfPath = wb.Path
    If Len(pdfPath) = 0 Then
        MsgBox "กรุณาบันทึกไฟล์ Excel ก่อน (Save) เพื่อให้กำหนดโฟลเดอร์สำหรับบันทึก PDF ได้", vbExclamation
        Exit Sub
    End If

    pdfName = ws.Name & " - " & Format(Now, "yyyy-mm-dd_hhmmss") & ".pdf"
    fullPath = pdfPath & Application.PathSeparator & pdfName

    ws.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=fullPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "บันทึก PDF เรียบร้อย: " & fullPath, vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ตัวอย่าง 2:
' สร้าง Print Area แบบไดนามิก (ใช้ UsedRange หรือระบุตามคอลัมน์ที่ต้องการ)
' แล้ว Export ทั้ง workbook เป็น PDF ไฟล์เดียว ในโฟลเดอร์เดียวกับ ThisWorkbook
Public Sub SaveWorkbookAsSinglePDF_SetPrintAreas()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pdfPath As String, pdfName As String, fullPath As String

    Set wb = ThisWorkbook ' เปลี่ยนเป็น ActiveWorkbook ถ้าจำเป็น

    pdfPath = wb.Path
    If Len(pdfPath) = 0 Then
        MsgBox "กรุณาบันทึกไฟล์ Excel ก่อน (Save) เพื่อให้กำหนดโฟลเดอร์สำหรับบันทึก PDF ได้", vbExclamation
        Exit Sub
    End If

    ' ตั้งค่า PrintArea ให้ทุกชีท (ตัวอย่างใช้ UsedRange)
    For Each ws In wb.Worksheets
        ' ถ้าต้องการจำกัดคอลัมน์ เช่น A ถึง G ให้ใช้:
        ' ws.PageSetup.PrintArea = ws.Range("A1", ws.Cells(ws.Rows.Count, "G").End(xlUp)).Address
        If Application.WorksheetFunction.CountA(ws.Cells) > 0 Then
            ws.PageSetup.PrintArea = ws.UsedRange.Address
        Else
            ws.PageSetup.PrintArea = "" ' ว่างถ้าไม่มีข้อมูล
        End If

        With ws.PageSetup
            .Orientation = xlPortrait
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
        End With
    Next ws

    pdfName = wb.Name
    ' เอา .xls หรือ .xlsm ออก แล้วต่อด้วย timestamp
    If InStrRev(pdfName, ".") > 0 Then pdfName = Left(pdfName, InStrRev(pdfName, ".") - 1)
    pdfName = pdfName & " - " & Format(Now, "yyyy-mm-dd_hhmmss") & ".pdf"
    fullPath = pdfPath & Application.PathSeparator & pdfName

    ' Export ทั้ง workbook เป็นไฟล์ PDF เดียว
    wb.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=fullPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "บันทึก PDF (ทั้งเล่ม) เรียบร้อย: " & fullPath, vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume Cleanup
End Sub