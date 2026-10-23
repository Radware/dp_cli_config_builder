Attribute VB_Name = "DeviceConfigGenerator"
Option Explicit
Private Const SHEET_NAME As String = "DP Template Variables"
Private Const TOKEN_PATTERN As String = "\$[A-Za-z_][A-Za-z0-9_-]*"

Public Sub ValidateTemplateVariables()
    RunGenerator 0
End Sub

Public Sub GenerateSelectedDevicePreview()
    Dim answer As String, col As Long, i As Long, letter As Long
    On Error GoTo Failed
    answer = UCase$(Trim$(InputBox("Enter the device column letter (B, C, ... AA):", "Select device column", "B")))
    If Len(answer) = 0 Then Exit Sub
    If Len(answer) > 3 Then GoTo InvalidColumn
    For i = 1 To Len(answer)
        letter = AscW(Mid$(answer, i, 1)) - AscW("A") + 1
        If letter < 1 Or letter > 26 Then GoTo InvalidColumn
        col = col * 26 + letter
    Next i
    If col < 2 Or col > ThisWorkbook.Worksheets(SHEET_NAME).Columns.Count Then GoTo InvalidColumn
    If Len(Trim$(CellText(ThisWorkbook.Worksheets(SHEET_NAME).Cells(1, col)))) = 0 Then
        MsgBox "Column " & answer & " has no device header in row 1.", vbExclamation
        Exit Sub
    End If
    RunGenerator 1, col
    Exit Sub
InvalidColumn:
    MsgBox "Enter a column letter between B and XFD, such as C or AA.", vbExclamation
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Column selection failed"
End Sub

Public Sub GenerateAllDeviceConfigs()
    RunGenerator 2
End Sub

' Mode 0 validates;
' Mode 1 previews selectedCol, tolerating missing/duplicate keys;
' Mode 2 generates all device columns only after strict validation passes.
Private Sub RunGenerator(ByVal mode As Long, Optional ByVal selectedCol As Long = 2)
    Dim ws As Worksheet, dlg As Object, rows As Object, rx As Object
    Dim matches As Object, m As Object, required As Object, fso As Object
    Dim jobs As Collection, paths As Object, job As Variant, key As Variant
    Dim template As String, issues As String, folder As String, path As String
    Dim token As String, label As String, content As String, value As String
    Dim r As Long, c As Long, lastRow As Long, lastCol As Long, count As Long, firstCol As Long
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    Set dlg = Application.FileDialog(3)
    dlg.Title = "Choose the text file with DP Configuration Template (UTF-8)"
    dlg.AllowMultiSelect = False
    If dlg.Show <> -1 Then Exit Sub
    template = ReadUtf8(dlg.SelectedItems(1))
    Set rows = CreateObject("Scripting.Dictionary")
    rows.CompareMode = vbBinaryCompare
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        token = Trim$(CellText(ws.Cells(r, 1)))
        If Len(token) > 0 Then
            If rows.Exists(token) Then
                issues = issues & "Duplicate " & token & " at A" & rows(token) & " and A" & r & vbCrLf
            Else
                rows.Add token, r
            End If
        End If
    Next r
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = TOKEN_PATTERN
    rx.Global = True
    Set matches = rx.Execute(template)
    Set required = CreateObject("Scripting.Dictionary")
    required.CompareMode = vbBinaryCompare
    For Each m In matches
        If Not required.Exists(m.Value) Then required.Add m.Value, True
    Next m
    For Each key In required.Keys
        If Not rows.Exists(key) Then issues = issues & "Missing " & key & vbCrLf
    Next key
    If mode = 0 Then
        If Len(issues) = 0 Then issues = "All " & required.Count & " template variables exist. No duplicate variable names."
        MsgBox issues, vbInformation, "Variable validation"
        Exit Sub
    End If
    If mode = 2 And Len(issues) > 0 Then
        MsgBox "Correct these issues before generating all devices:" & vbCrLf & issues, vbExclamation
        Exit Sub
    End If
    Set dlg = Application.FileDialog(4)
    dlg.Title = "Choose output folder"
    If dlg.Show <> -1 Then Exit Sub
    folder = dlg.SelectedItems(1)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set jobs = New Collection
    Set paths = CreateObject("Scripting.Dictionary")
    paths.CompareMode = vbTextCompare
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    firstCol = 2
    If mode = 1 Then
        firstCol = selectedCol
        lastCol = selectedCol
    End If
    For c = firstCol To lastCol
        If Len(Trim$(CellText(ws.Cells(1, c)))) > 0 Then
            label = ""
            If rows.Exists("$device-name") Then label = Trim$(CellText(ws.Cells(rows("$device-name"), c)))
            If Len(label) = 0 Then label = Trim$(CellText(ws.Cells(1, c)))
            ' Column number keeps filenames unique even when site/device names repeat.
            path = fso.BuildPath(folder, SafeName(label) & "_col" & c & IIf(mode = 1, "_preview", "") & ".txt")
            If paths.Exists(path) Or fso.FileExists(path) Then Err.Raise vbObjectError + 1, , "Output already exists: " & path
            paths.Add path, True
            content = ExpandTemplate(template, matches, rows, ws, c)
            jobs.Add Array(path, content)
        End If
    Next c
    If jobs.Count = 0 Then Err.Raise vbObjectError + 2, , "No device headers in row 1, starting at B1."
    For Each job In jobs
        WriteUtf8 CStr(job(0)), CStr(job(1))
        count = count + 1
    Next job
    MsgBox "Created " & count & " file(s) in " & folder & IIf(mode = 1 And Len(issues) > 0, vbCrLf & "Preview only. Missing variables use no-value; duplicate variables use the first row." & vbCrLf & issues, ""), vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description & vbCrLf & "Files created: " & count, vbExclamation, "Generation stopped"
End Sub

Private Function ExpandTemplate(ByVal template As String, ByVal matches As Object, ByVal rows As Object, ByVal ws As Worksheet, ByVal col As Long) As String
    Dim m As Object, value As String, result As String, pos As Long
    pos = 1
    For Each m In matches
        value = "no-value"
        If rows.Exists(m.Value) Then
            value = CellText(ws.Cells(rows(m.Value), col))
            If Len(Trim$(value)) = 0 Then value = "no-value"
        End If
        result = result & Mid$(template, pos, m.FirstIndex + 1 - pos) & value
        pos = m.FirstIndex + m.Length + 1
    Next m
    ExpandTemplate = result & Mid$(template, pos)
End Function

Private Function CellText(ByVal cell As Range) As String
    If IsError(cell.Value2) Then Err.Raise vbObjectError + 3, , "Excel error at " & cell.Address(External:=True)
    CellText = CStr(cell.Value2)
End Function

Private Function SafeName(ByVal value As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
        value = Replace(value, CStr(ch), "_")
    Next ch
    Dim i As Long
    For i = 0 To 31
        value = Replace(value, Chr$(i), "_")
    Next i
    value = Trim$(value)
    Do While Len(value) > 0
        If Right$(value, 1) <> "." And Right$(value, 1) <> " " Then Exit Do
        value = Left$(value, Len(value) - 1)
    Loop
    If Len(value) = 0 Then value = "device"
    SafeName = Left$(value, 100)
End Function

Private Function ReadUtf8(ByVal path As String) As String
    Dim s As Object
    Set s = CreateObject("ADODB.Stream")
    s.Type = 2
    s.Charset = "utf-8"
    s.Open
    s.LoadFromFile path
    ReadUtf8 = s.ReadText
    s.Close
End Function

Private Sub WriteUtf8(ByVal path As String, ByVal content As String)
    Dim s As Object, b As Object
    Set s = CreateObject("ADODB.Stream")
    s.Type = 2
    s.Charset = "utf-8"
    s.Open
    s.WriteText content
    s.Position = 0
    s.Type = 1
    s.Position = 3
    Set b = CreateObject("ADODB.Stream")
    b.Type = 1
    b.Open
    s.CopyTo b
    b.SaveToFile path, 1
    b.Close
    s.Close
End Sub
