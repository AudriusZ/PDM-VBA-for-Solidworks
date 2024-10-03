Attribute VB_Name = "Module1"
Sub UpdateSolidWorksProperties()
    Dim swApp As Object
    Dim swModel As Object
    Dim FilePath As String
    Dim FileName As String
    Dim PartNo As String
    Dim Revision As String
    Dim Description As String
    Dim Material As String
    Dim i As Integer
    Dim folderPath As String
    Dim fileType As Integer
    Dim Errors As Long
    Dim Warnings As Long
    Dim saveStatus As Long
    Dim logMissingFiles As String
    Dim startPartNo As String
    Dim endPartNo As String
    Dim foundStart As Boolean

    ' Set the folder path where SolidWorks files are stored
    folderPath = "C:\Data\Upwork-PPH Large Files\20231110 Jeevan Technology\Solidworks CAD Library\"

    ' Set the start and end part numbers for the subset
    startPartNo = "JCS00001" ' Replace with your start part number
    endPartNo = "JCS00250" ' Replace with your end part number

    ' Initialize variables
    foundStart = False

    ' Connect to SolidWorks
    On Error Resume Next
    Set swApp = GetObject(, "SldWorks.Application")
    If swApp Is Nothing Then
        Set swApp = CreateObject("SldWorks.Application")
    End If
    If swApp Is Nothing Then
        MsgBox "SolidWorks could not be started. Please ensure it is installed."
        Exit Sub
    End If
    On Error GoTo 0

    ' Initialize missing files log
    logMissingFiles = ""

    ' Loop through each row in Excel
    For i = 3 To ThisWorkbook.Sheets("JCS Database").Cells(Rows.Count, 1).End(xlUp).Row
        PartNo = ThisWorkbook.Sheets("JCS Database").Cells(i, 1).Value

        ' Check if the current part is within the subset range
        If PartNo = startPartNo Then
            foundStart = True
        End If
        If Not foundStart Then GoTo SkipIteration
        If PartNo = endPartNo Then foundStart = False

        ' Get the file name for the part or assembly
        FileName = Dir(folderPath & PartNo & "*.SLDPRT")
        If FileName = "" Then
            FileName = Dir(folderPath & PartNo & "*.SLDASM")
            fileType = 2 ' Assembly files
        Else
            fileType = 1 ' Part files
        End If

        If FileName = "" Then
            logMissingFiles = logMissingFiles & PartNo & vbCrLf
            GoTo SkipIteration
        End If

        FilePath = folderPath & FileName

        Revision = ThisWorkbook.Sheets("JCS Database").Cells(i, 2).Value
        Description = ThisWorkbook.Sheets("JCS Database").Cells(i, 3).Value
        Material = ThisWorkbook.Sheets("JCS Database").Cells(i, 4).Value

        ' Open the SolidWorks file
        On Error Resume Next
        Set swModel = swApp.OpenDoc6(FilePath, fileType, 0, "", Errors, Warnings)
        If Err.Number <> 0 Or swModel Is Nothing Then
            MsgBox "Failed to open: " & FilePath & " Error: " & Err.Description & " Errors: " & Errors & " Warnings: " & Warnings
            Err.Clear
            On Error GoTo 0
            GoTo SkipIteration
        End If
        On Error GoTo 0

        ' Add custom properties if non-existing
        Call swModel.AddCustomInfo3("", "PartNo", 30, PartNo)
        Call swModel.AddCustomInfo3("", "Revision", 30, Revision)
        Call swModel.AddCustomInfo3("", "Description", 30, Description)
        Call swModel.AddCustomInfo3("", "Material", 30, Material)

        ' Set new custom properties if already existing
        Call swModel.CustomInfo2("", "PartNo", PartNo)
        Call swModel.CustomInfo2("", "Revision", Revision)
        Call swModel.CustomInfo2("", "Description", Description)
        Call swModel.CustomInfo2("", "Material", Material)

        ' Save and close the document with full path
        saveStatus = swModel.SaveAs(FilePath)
        If Not saveStatus Then
            MsgBox "Failed to SaveAs: " & FilePath
        End If
        swApp.CloseDoc FilePath

SkipIteration:
        ' Reset FileName for the next iteration
        FileName = ""
    Next i

    ' Check if there were any missing files and save the log
    If Len(logMissingFiles) > 0 Then
        Dim fs As Object, a As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(folderPath & "missing_files_log.txt", True)
        a.WriteLine "Missing Files:" & vbCrLf & logMissingFiles
        a.Close
        MsgBox "Some files were not found. Check the 'missing_files_log.txt' file in the SolidWorks folder."
    Else
        MsgBox "Properties updated successfully!"
    End If

    ' Clean up
    Set swModel = Nothing
    Set swApp = Nothing
End Sub
