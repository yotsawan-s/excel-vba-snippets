Option Explicit

' ฟังก์ชันหลักในการตรวจสอบไฟล์ Backup
Public Sub CheckBackupFiles()
    Dim mainPath As String
    Dim projectCodes As Variant
    Dim resultSheet As Worksheet
    Dim i As Long
    
    ' กำหนด Path หลัก (แก้ไขตาม Path จริงของคุณ)
    mainPath = "C:\Backup\" ' *** แก้ไข Path ตามต้องการ ***
    
    ' รายการ Project Code ที่ต้องการตรวจสอบ
    projectCodes = Array("TP001", "TP002", "TP003") ' *** แก้ไขรายการ Project Code ***
    
    ' สร้าง/เตรียม Worksheet สำหรับแสดงผล
    On Error Resume Next
    Set resultSheet = ThisWorkbook.Worksheets("BackupCheckResult")
    If resultSheet Is Nothing Then
        Set resultSheet = ThisWorkbook.Worksheets.Add
        resultSheet.Name = "BackupCheckResult"
    Else
        resultSheet.Cells.Clear
    End If
    On Error GoTo 0
    
    ' สร้างหัวตาราง
    With resultSheet
        .Range("A1:E1").Value = Array("Project Code", "Pattern", "File Found", "File Path", "Status")
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(217, 217, 217)
    End With
    
    Dim currentRow As Long
    currentRow = 2
    
    ' ตรวจสอบแต่ละ Project Code
    For i = LBound(projectCodes) To UBound(projectCodes)
        currentRow = CheckProjectBackup(mainPath, CStr(projectCodes(i)), resultSheet, currentRow)
    Next i
    
    ' ปรับความกว้างคอลัมน์อัตโนมัติ
    resultSheet.Columns("A:E").AutoFit
    
    MsgBox "การตรวจสอบเสร็จสิ้น!", vbInformation
End Sub

' ฟังก์ชันตรวจสอบ Backup ของแต่ละ Project
Private Function CheckProjectBackup(ByVal mainPath As String, _
                                     ByVal projectCode As String, _
                                     ByVal resultSheet As Worksheet, _
                                     ByVal startRow As Long) As Long
    
    ' กำหนด Pattern ของไฟล์ที่ต้องตรวจสอบ (9 รูปแบบ)
    Dim filePatterns As Variant
    filePatterns = Array( _
        "P1_*_V*.*.xl??", _
        "P2_*_V*.*.xl??", _
        "P3_*_V*.*.xl??", _
        "D1_*_V*.*.xl??", _
        "D2_*_V*.*.xl??", _
        "D3_*_V*.*.do??", _
        "R1_*_V*.*.xl??", _
        "R2_*_V*.*.pd?", _
        "M1_*_V*.*.xl??" _
    ) ' *** แก้ไข Pattern ตามต้องการ ***
    
    Dim misFolder As String
    Dim documentPath As String
    Dim currentRow As Long
    Dim i As Long
    
    currentRow = startRow
    
    ' สร้าง Path ตามโครงสร้าง: MIS > TP001_xxxx > Document
    misFolder = mainPath & "MIS\"
    
    ' ตรวจสอบว่ามี Folder MIS หรือไม่
    If Not FolderExists(misFolder) Then
        resultSheet.Cells(currentRow, 1).Value = projectCode
        resultSheet.Cells(currentRow, 2).Value = "MIS Folder"
        resultSheet.Cells(currentRow, 3).Value = "Not Found"
        resultSheet.Cells(currentRow, 4).Value = misFolder
        resultSheet.Cells(currentRow, 5).Value = "❌ ERROR"
        resultSheet.Cells(currentRow, 5).Interior.Color = RGB(255, 199, 206)
        CheckProjectBackup = currentRow + 1
        Exit Function
    End If
    
    ' ค้นหา Folder ที่ขึ้นต้นด้วย Project Code
    Dim projectFolder As String
    projectFolder = FindProjectFolder(misFolder, projectCode)
    
    If projectFolder = "" Then
        resultSheet.Cells(currentRow, 1).Value = projectCode
        resultSheet.Cells(currentRow, 2).Value = "Project Folder"
        resultSheet.Cells(currentRow, 3).Value = "Not Found"
        resultSheet.Cells(currentRow, 4).Value = misFolder & projectCode & "_*"
        resultSheet.Cells(currentRow, 5).Value = "❌ ERROR"
        resultSheet.Cells(currentRow, 5).Interior.Color = RGB(255, 199, 206)
        CheckProjectBackup = currentRow + 1
        Exit Function
    End If
    
    ' สร้าง Path ของ Document Folder
    documentPath = misFolder & projectFolder & "\Document\"
    
    ' ตรวจสอบว่ามี Folder Document หรือไม่
    If Not FolderExists(documentPath) Then
        resultSheet.Cells(currentRow, 1).Value = projectCode
        resultSheet.Cells(currentRow, 2).Value = "Document Folder"
        resultSheet.Cells(currentRow, 3).Value = "Not Found"
        resultSheet.Cells(currentRow, 4).Value = documentPath
        resultSheet.Cells(currentRow, 5).Value = "❌ ERROR"
        resultSheet.Cells(currentRow, 5).Interior.Color = RGB(255, 199, 206)
        CheckProjectBackup = currentRow + 1
        Exit Function
    End If
    
    ' ตรวจสอบไฟล์ตาม Pattern แต่ละรูปแบบ
    For i = LBound(filePatterns) To UBound(filePatterns)
        Dim foundFile As String
        foundFile = FindFileByPattern(documentPath, CStr(filePatterns(i)))
        
        resultSheet.Cells(currentRow, 1).Value = projectCode
        resultSheet.Cells(currentRow, 2).Value = filePatterns(i)
        
        If foundFile <> "" Then
            resultSheet.Cells(currentRow, 3).Value = foundFile
            resultSheet.Cells(currentRow, 4).Value = documentPath & foundFile
            resultSheet.Cells(currentRow, 5).Value = "✓ OK"
            resultSheet.Cells(currentRow, 5).Interior.Color = RGB(198, 239, 206)
        Else
            resultSheet.Cells(currentRow, 3).Value = "Not Found"
            resultSheet.Cells(currentRow, 4).Value = documentPath
            resultSheet.Cells(currentRow, 5).Value = "❌ MISSING"
            resultSheet.Cells(currentRow, 5).Interior.Color = RGB(255, 235, 156)
        End If
        
        currentRow = currentRow + 1
    Next i
    
    CheckProjectBackup = currentRow
End Function

' ฟังก์ชันตรวจสอบว่า Folder มีอยู่หรือไม่
Private Function FolderExists(ByVal folderPath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(folderPath)
    Set fso = Nothing
End Function

' ฟังก์ชันค้นหา Folder ที่ขึ้นต้นด้วย Project Code
Private Function FindProjectFolder(ByVal misPath As String, ByVal projectCode As String) As String
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(misPath) Then
        FindProjectFolder = ""
        Exit Function
    End If
    
    Set folder = fso.GetFolder(misPath)
    
    For Each subFolder In folder.SubFolders
        If Left(subFolder.Name, Len(projectCode)) = projectCode Then
            FindProjectFolder = subFolder.Name
            Exit Function
        End If
    Next subFolder
    
    FindProjectFolder = ""
    Set fso = Nothing
End Function

' ฟังก์ชันค้นหาไฟล์ตาม Pattern
Private Function FindFileByPattern(ByVal folderPath As String, ByVal pattern As String) As String
    Dim fileName As String
    
    fileName = Dir(folderPath & pattern)
    
    If fileName <> "" Then
        FindFileByPattern = fileName
    Else
        FindFileByPattern = ""
    End If
End Function

' ฟังก์ชันตรวจสอบไฟล์ที่ตรงกับ Pattern แบบละเอียด (ใช้ Regex)
Private Function IsFileMatchPattern(ByVal fileName As String, ByVal pattern As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = True
        .IgnoreCase = True
        .pattern = pattern
    End With
    
    IsFileMatchPattern = regex.Test(fileName)
    Set regex = Nothing
End Function
