Sub SaveAsToFolder_Timestamped()
    Dim FileName As String
    Dim FolderPath As String
    Dim FullPath As String

    ' Set the folder path to the Output folder
    FolderPath = "C:\YourPath\Output\" 

    ' Create a timestamped subfolder
    Dim SubFolder As String
    SubFolder = Format(Now, "YYYYMMDD-HHMMSS")

    ' Create full path
    FullPath = FolderPath & SubFolder & "\"

    ' Create the folder if it doesn't exist
    If Dir(FullPath, vbDirectory) = "" Then
        MkDir FullPath
    End If

    ' Save the file in the timestamped folder
    FileName = "YourFileName.xlsx"
    ActiveWorkbook.SaveCopyAs FullPath & FileName
End Sub

Sub SaveAsToFolder_Overwrite()
    Dim FileName As String
    Dim FolderPath As String

    ' Set the folder path to the Output folder
    FolderPath = "C:\YourPath\Output\" 

    ' Set the file name
    FileName = "YourFileName.xlsx"

    ' Save the file, overwriting if it exists
    ActiveWorkbook.SaveCopyAs FolderPath & FileName
End Sub
