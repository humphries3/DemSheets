Attribute VB_Name = "Module1"
Option Explicit
Public gFilSelDir As Boolean

Public Type typShCtl
    Sh As Worksheet
    RowLast As Long
    RowCurr As Long
    Used As Range
    colMap As New Scripting.Dictionary
    End Type
    
Public logCtl As typShCtl

Function CombSheets() As Boolean
''''''''''''''''''''''''''''''''
Dim Fd As FileDialog
Dim FdSel As Integer
Dim FnArr() As String
Dim Fn As String
''''''''''''''''
frmFilSel.Show vbModal

Select Case True

    'Directory selection requested:
    Case gFilSelDir
    Set Fd = Application.FileDialog(msoFileDialogFolderPicker)
    Fd.Title = "Select a Folder for consolidation"
    Fd.AllowMultiSelect = False
    Fd.InitialFileName = Application.DefaultFilePath
    If Environ("username") = "WAH" Then
        Fd.InitialFileName = "D:\Users\WAH\Google Drive\DemSheets\Combine-Samples\"
        End If
    
    Select Case Fd.Show
    
        'Directory specified:
        Case -1
        Fn = Dir(Fd.SelectedItems(1) & "\*.xlsx")
        Do While Len(Fn) > 0
            FdSel = FdSel + 1
            ReDim Preserve FnArr(1 To FdSel)
            FnArr(FdSel) = Fd.SelectedItems(1) & "\" & Fn
            Fn = Dir
            Loop
            
        CombSheets = CombSheetsRun(FnArr)
        
        'No directory specified:
        Case Else
        CombSheets = True
        End Select
    
    'Individual file selection requested
    Case Else
    Set Fd = Application.FileDialog(msoFileDialogFilePicker)
    Fd.AllowMultiSelect = True
    Fd.Filters.Add "Documents", "*.xls; *.xlsx"
    Fd.Title = "Select workbooks for consolidation"
    If Environ("username") = "WAH" Then Fd.InitialFileName = "D:\Users\WAH\Google Drive\DemSheets\Combine-Samples\"
    
    Select Case Fd.Show
    
        'Files specified:
        Case -1
        For FdSel = 1 To Fd.SelectedItems.Count
            ReDim Preserve FnArr(1 To FdSel)
            FnArr(FdSel) = Fd.SelectedItems(FdSel)
            Next FdSel
        CombSheets = CombSheetsRun(FnArr)
        
        'No files specified:
        Case Else
        CombSheets = True
        End Select
    
    End Select
    
End Function

Function CombSheetsRun(FnArr() As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim OutWbk As Workbook
Dim Fd As FileDialog
Dim FdSel  As Integer
Dim lineText As String
Dim outWsh As Worksheet
Dim logWsh As Worksheet
Dim regExp As New regExp
Dim outRow As Integer
Dim outErr As Integer
Dim matchList As MatchCollection
Dim matchC As Integer
Dim newData(1 To 3)
Dim fnCurr As Integer
Dim iWbk As Workbook
Dim iWsh As Worksheet
Dim oWsh As Worksheet
Dim Sh As Worksheet
Dim iRng As Range
Dim colIdx As Integer
Dim oRows As Integer
Dim conCtl As typShCtl
Dim inpCtl As typShCtl
Dim logColNames
Dim fnSuggest
Dim shInventory As New Scripting.Dictionary
Dim verMap As New Scripting.Dictionary
Dim jnkMap As New Scripting.Dictionary
''''''''''''''''''''''''''''''''''''''''
    
'1.0 instantiate output workbook
''''''''''''''''''''''''''''''
Set OutWbk = Workbooks.Add
verMap.CompareMode = TextCompare

'1.1 need to pre-emptively ACTIVATE (any) sheet before calling FREEZE:
'but let's do SHEET2 so user can see data build....
OutWbk.Worksheets(1).Activate
shInventory.CompareMode = TextCompare

'1.2 build LOG sheet:
logColNames = Array("Time", "Code", "!", "Topic", "Detail")
logCtl = newShCtl()
With logCtl
    Set .Sh = OutWbk.Worksheets.Add
    .Sh.Name = "Log"
    .Sh.Rows(1).Font.Bold = True
    .Sh.Rows(1).HorizontalAlignment = xlCenter
    .Sh.Cells.Font.Name = "ms reference sans serif"
    .Sh.Cells.Font.Size = 10
    .colMap.CompareMode = TextCompare
    .RowLast = 1
    For colIdx = 1 To UBound(logColNames) + 1
        .Sh.Cells(.RowLast, colIdx) = logColNames(colIdx - 1)
        .colMap.Add logColNames(colIdx - 1), colIdx
        Next colIdx
    .RowLast = 1
    .Sh.Columns(.colMap("Code")).NumberFormat = "000"
    shInventory.Add .Sh.Name, 1
    Call Freeze(.Sh)
    End With

'1.3 build CONSOLIDATED sheet:
With conCtl
    Set .Sh = OutWbk.Worksheets.Add
    .Sh.Name = "Consolidated"
    .Sh.Rows(1).Font.Bold = True
    .Sh.Rows(1).HorizontalAlignment = xlCenter
    shInventory.Add .Sh.Name, 2
    Call Freeze(.Sh)
    End With

'1.4 delete extraneous sheets:
Application.DisplayAlerts = False
For Each Sh In OutWbk.Worksheets
    Select Case True
        Case shInventory.Exists(Sh.Name)
        'nop
        Case Else
        Sh.Delete
        End Select
    Next Sh
Application.DisplayAlerts = True

'1.5 name and save output book:
fnSuggest = "DemSheets-Consol-" & Format(Now, "yyyymmdd-hhmmss") & ".xlsx"
Set Fd = Application.FileDialog(msoFileDialogSaveAs)
Fd.Title = "Name the output workbook"
Fd.InitialFileName = Application.DefaultFilePath & fnSuggest
'NB: If folder doesn't exist FD just uses default path - no error msg!
If Environ("username") = "WAH" Then Fd.InitialFileName = "D:\Users\WAH\Google Drive\DemSheets\Combine-Out\" & fnSuggest
Call logMsg("001I Combining worksheets to...", Fd.InitialFileName)

Select Case Fd.Show

    Case -1
    OutWbk.SaveAs Fd.SelectedItems(1)

    For fnCurr = 1 To UBound(FnArr)
    
        Call logMsg("004I Reading workbook...", FnArr(fnCurr))
        Set iWbk = Workbooks.Open(Filename:=FnArr(fnCurr), ReadOnly:=True)
        inpCtl = newShCtl()
        
        With inpCtl
        
            Set .Sh = iWbk.Worksheets(1)
            Set .Used = .Sh.UsedRange
            Call logMsg("005I Range", .Used.Address(xlA1))
            
            
            Select Case fnCurr
            
                Case 1
                Select Case True
                    Case Not regColNames(inpCtl, inpCtl.colMap, verMap)
                    Case Not nulSheet(inpCtl)
                    Case Else
                    oRows = movHead(inpCtl, conCtl)
                    conCtl.Sh.Columns(verMap("Time Canvassed")).NumberFormat = "h:mm AM/PM"
                    oRows = movData(inpCtl, conCtl)
                    End Select
                    
                Case Else
                jnkMap.RemoveAll
                Select Case True
                    Case Not regColNames(inpCtl, inpCtl.colMap, jnkMap)
                    Case Not verColNames(inpCtl, verMap)
                    Case Not nulSheet(inpCtl)
                    Case Else
                    oRows = movData(inpCtl, conCtl)
                    End Select
                    
                End Select
                
            End With
            
        iWbk.Close
        Next fnCurr
        
    Call logMsg("008I Output", "DataRows(" & (conCtl.RowLast - 1) & ")")
    conCtl.Sh.Cells.EntireColumn.AutoFit
    logCtl.Sh.Cells.EntireColumn.AutoFit
    
    'MsgBox "Conversion finished" _
        & vbCrLf & "Error(s): " & (outErr - 1) _
        & vbCrLf & "Converted line(s): " & (outRow - 1)
        
    OutWbk.Save
    ' add display output after combine, end when DONE , etc...
 '   OutWbk.Close
    CombSheetsRun = True
    
    'get rid of skeleton output book, return:
    Case Else
    OutWbk.Saved = True
    OutWbk.Close
    CombSheetsRun = True
    End Select
    
End Function

Function regColNames( _
    inpCtl As typShCtl, _
    addMap As Scripting.Dictionary, _
    verMap As Scripting.Dictionary) _
    As Boolean
''''''''''''''
Dim colIdx As Integer
Dim colName As String
'''''''''''''''''''''

With inpCtl

    addMap.CompareMode = TextCompare
    
    For colIdx = 1 To .Used.Columns.Count
        colName = Trim(.Sh.Cells(1, colIdx))
    
        Select Case True
            Case colName = ""
            Call logMsg("001E Column heading blank", .Sh.Parent.Name & " : " & .Sh.Name & " : " & colIdx)
            Exit Function
            
            Case addMap.Exists(colName)
            Call logMsg("002E Column heading duplicated", .Sh.Parent.Name & " : " & .Sh.Name & " : " * colName & " : " & colIdx)
            Exit Function
            
            Case Else
            addMap.Add colName, colIdx
            verMap.Add colName, colIdx
            End Select
            
        Next colIdx
        
    regColNames = True
    End With
    
End Function

Function verColNames( _
    inpCtl As typShCtl, _
    verMap As Scripting.Dictionary) _
    As Boolean
''''''''''''''
Dim colIdx As Integer
Dim colName As String
Dim verIdx As Integer
'''''''''''''''''''''

With inpCtl

    Select Case True
    
        Case verMap.Count <> .Used.Columns.Count
        Call logMsg("007E Column counts differ", "Expect:" & verMap.Count & " Received:" & .Used.Columns.Count)
        Exit Function
        
        Case Else
        For colIdx = 1 To .Used.Columns.Count
            colName = Trim(.Sh.Cells(1, colIdx))
            verIdx = IIf(verMap.Exists(colName), verMap(colName), 0)
        
            Select Case True
            
                Case verIdx = 0
                Call logMsg("001E Column heading unregistered", "[" & colName & "]")
                Exit Function
                
                Case verIdx <> colIdx
                Call logMsg("002E Column heading out of order", "[" & colName & "]")
                Exit Function
                
                Case Else
                'nop
                End Select
                
            Next colIdx
            
        End Select
        
    End With
        
verColNames = True
End Function

Function nulSheet(inpCtl As typShCtl) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''
With inpCtl

    Select Case True
        Case .Used.Rows.Count < 2
        Call logMsg("003W Spreadsheet empty", .Sh.Parent.Name)
        
        Case Else
        nulSheet = True
        End Select
        
    End With
    
End Function

Function movHead(inpCtl As typShCtl, conCtl As typShCtl) As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim colHeads
''''''''''''
colHeads = inpCtl.Used.Rows(1)
With conCtl.Sh
    .Range(.Cells(1, 1), .Cells(1, inpCtl.Used.Columns.Count)) = colHeads
    End With
conCtl.RowLast = conCtl.RowLast + 1
End Function

Function movData(inpCtl As typShCtl, conCtl As typShCtl) As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim iRng As Range
Dim oRng As Range
'''''''''''''''''
Set iRng = inpCtl.Used.Offset(1, 0).Resize(inpCtl.Used.Rows.Count - 1, inpCtl.Used.Columns.Count)
With conCtl
    Set oRng = _
    .Sh.Range(.Sh.Cells(.RowLast + 1, 1), .Sh.Cells(.RowLast + iRng.Rows.Count, iRng.Columns.Count))
    End With
Application.ScreenUpdating = False
oRng.Value = iRng.Value
Application.ScreenUpdating = True
conCtl.RowLast = conCtl.RowLast + iRng.Rows.Count
End Function

Sub logMsg(msgTxt As String, auxData)
'''''''''''''''''''''''''''''''''''''
With logCtl
    .RowLast = .RowLast + 1
    .Sh.Cells(.RowLast, .colMap("Time")) = Format(Now, "yyyy.mm.dd.hhmm")
    .Sh.Cells(.RowLast, .colMap("Code")) = Mid(msgTxt, 1, 3)
    .Sh.Cells(.RowLast, .colMap("!")) = Mid(msgTxt, 4, 1)
    .Sh.Cells(.RowLast, .colMap("Topic")) = Mid(msgTxt, 5)
    .Sh.Cells(.RowLast, .colMap("Detail")) = auxData
    End With
End Sub

Sub Freeze(myWsh As Worksheet)
''''''''''''''''''''''''''''''
Dim currWsh As Worksheet
''''''''''''''''''''''''
Set currWsh = ActiveWindow.ActiveSheet
myWsh.Activate
With ActiveWindow
    .SplitColumn = 0
    .SplitRow = 1
    .FreezePanes = True
    End With
currWsh.Activate
End Sub

Function newShCtl() As typShCtl
'''''''''''''''''''''''''''''''
'just returns new shCtl obj
newShCtl.colMap.CompareMode = TextCompare
End Function


