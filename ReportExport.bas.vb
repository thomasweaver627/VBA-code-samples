Public Function ReportExport(strSQL As String, strDir As String, strExportFile As String, strReportTitle As String)



    On Error GoTo Err_Stop

    Dim db As Database
    Dim objExcel As Object
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSheet As Object
    Dim strDir As String
    Dim strExportFileFull As String
    Dim rsExport As Recordset
    Dim aryCols() As String
    Dim aryRS As Variant
    Dim intLastRow As Long
    Dim intCol As Long
    Dim intRow As Long
    Dim intFieldCount As Long
    Dim i As Long
    Dim j As Long
    Dim col As Long
    Dim intShift As Integer
    Dim intRSCount As Long
    Dim oFS As FileSystemObject

    'Dim dteDateLastModified As Date
    'Dim strExistingFile As String
 
    ' additional code if you need this func to also read an existing file or previous export file
    ' get 'Date modified' info from existing file
    'dteDateLastModified = oFS.GetFile(strExistingFile).Datelastmodified

    ' If dteDateLastModified < Date Then
    '     MsgBox "Please ensure file has been updated." & vbNewLine & strExistingFile
    '     GoTo Err_Exit
    ' End If

    DoCmd.SetWarnings False
    Set db = CurrentDb
    Set oFS = New FileSystemObject
    Set objExcel = CreateObject("Excel.Application")

    Set rsExport = db.OpenRecordset(strSQL, dbOpenDynaset)

    'check that file args were passed
    If strDir = "" Or strExportFile = "" Then
        MsgBox "The output folder and/or file are not defined", vbOKOnly, "Report Export"
        Exit Function
    End If

    'check that directory exists
    If Not oFS.FolderExists(strDir) Then
        MsgBox "The report folder does not exist and could not be created: " & vbCrLf & strDir, vbOKOnly, "Report Export"
        Exit Function
    End If

    strExportFileFull = strDir & "\" & strExportFile & "_" & Format(Now, "YYYYMMDD_hhmmss") & ".xlsx"

    'Create and open Excel file
    With objExcel
        .DisplayAlerts = False ' True for testing
        .Workbooks.Add.SaveAs FileName:=strExportFileFull, ReadOnlyRecommended:=False
        .ActiveWorkbook.Sheets("Sheet1").Name = strExportFile
        Set xlWB = .Workbooks.Open(strExportFileFull)
        Set xlSheet = xlWB.Sheets(strExportFile)
    End With

    ' uncomment for testing
    ' objExcel.Visible = True

    ' allows func to work with any column count, can be made longer if needed
    ' split creates 0-based index array
    aryCols = Split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW,AX,AY,AZ", ",")

    ' gets record count
    rsExport.MoveLast
    intRSCount = rsExport.RecordCount
    rsExport.MoveFirst

    ' stores rsExport data as a 0-based 2D array (columns, rows)
    aryRS = rsExport.GetRows(intRSCount - 1)
    ' gets field/column count of array
    intFieldCount = UBound(aryRS, 1)

    'shifts down starting row if strReportTitle is passed
    If strReportTitle = "" Then
        intShift = 1
    Else
        intShift = 2
    End If

    'populate column/field names
    intCol = 0
    For j = LBound(aryRS, 1) To UBound(aryRS, 1)
        xlSheet.Range(aryCols(j) & intShift).Value = rsExport.Fields(intCol).Name
        intCol = intCol + 1
    Next j

    'populate the rest of the sheet
    For j = LBound(aryRS, 1) To UBound(aryRS, 1)
        intRow = intShift + 1
        For i = LBound(aryRS, 2) To UBound(aryRS, 2)
            xlSheet.Range(aryCols(j) & intRow).Value = aryRS(j, i)
            intRow = intRow + 1
        Next i
    Next j

    rsExport.Close
    Set rsExport = Nothing
        
    With xlSheet
        'Get Last Row
        intLastRow = .Range("A" & .Rows.Count).End(xlUp).Row

        'Format Report Columns
        For col = 0 To intFieldCount
            intRow = intShift + 1
            .Range(aryCols(col) & intRow, aryCols(col) & intLastRow).HorizontalAlignment = xlCenter
            If IsDate(.Range(aryCols(col) & intRow).Value) Then
                .Range(aryCols(col) & intRow, aryCols(col) & intLastRow).HorizontalAlignment = xlRight
                .Range(aryCols(col) & intRow, aryCols(col) & intLastRow).NumberFormat = "mm/dd/yyyy"
            ElseIf IsNumeric(.Range(aryCols(col) & intRow).Value) Then
                .Range(aryCols(col) & intRow, aryCols(col) & intLastRow).HorizontalAlignment = xlRight
                If .Range(aryCols(col) & intRow).Value > 9 Then
                    .Range(aryCols(col) & intRow, aryCols(col) & intLastRow).NumberFormat = "00000000-00" ' change as needed
                End If
            ElseIf Left(.Range(aryCols(col) & intRow).Value, 1) = "$" Then
                .Range(aryCols(col) & intRow, aryCols(col) & intLastRow).HorizontalAlignment = xlRight
                .Range(aryCols(col) & intRow, aryCols(col) & intLastRow).NumberFormat = "$#,##0.00"
            End If
        Next col

        ' use .TextToColumns to clean up messy string data
        ' If IsEmpty("D2") = False Then
        ' .Range("C2", "C" & intLastRow).TextToColumns
        ' End If

        'Format Report Borders
        .Range(aryCols(0) & intShift, aryCols(intFieldCount) & intLastRow).Borders(xlEdgeLeft).Weight = xlThin
        .Range(aryCols(0) & intShift, aryCols(intFieldCount) & intLastRow).Borders(xlEdgeRight).Weight = xlThin
        .Range(aryCols(0) & intShift, aryCols(intFieldCount) & intLastRow).Borders(xlEdgeBottom).Weight = xlThin
        .Range(aryCols(0) & intShift, aryCols(intFieldCount) & intLastRow).Borders(xlEdgeTop).Weight = xlThin
        .Range(aryCols(0) & intShift, aryCols(intFieldCount) & intLastRow).Borders(xlInsideVertical).Weight = xlThin
        .Range(aryCols(0) & intShift, aryCols(intFieldCount) & intLastRow).Borders(xlInsideHorizontal).Weight = xlThin

        'to insert totals at  bottom of spreadsheet
        ' .Range("B" & intLastRow + 2).Formula = "Total # of ______:"
        ' .Range("C" & intLastRow + 2).Formula = "=COUNTA(B2:B" & intLastRow & ")"

        ' Format Totals Fields
        ' With .Range("B" & intLastRow + 2, "C" & intLastRow + 2)
        '     .Font.Bold = True
        '     .Borders(xlEdgeTop).Weight = xlMedium
        '     .Borders(xlEdgeBottom).Weight = xlMedium
        ' End With

        'Hide Columns
        '.Range("M:N").EntireColumn.Hidden = True

        'Freeze Top Row
        .Range(aryCols(0) & intShift).Select
        objExcel.ActiveWindow.SplitRow = intShift
        objExcel.ActiveWindow.FreezePanes = True

        'Set Auto Filter and Auto Fit
        .AutoFilterMode = False
        .Range(aryCols(0) & intShift & ":" & aryCols(intFieldCount) & intLastRow).AutoFilter
        .Columns(aryCols(0) & ":" & aryCols(intFieldCount)).Select
        .Range(aryCols(0) & ":" & aryCols(intFieldCount)).Columns.AutoFit

        'Format Header/Fields
        .Range("A1:" & aryCols(intFieldCount) & intShift).Interior.Color = RGB(47, 96, 128)
        .Range("A1:" & aryCols(intFieldCount) & intShift).Font.Color = RGB(242, 242, 242)
        If strReportTitle <> "" Then 
            .Range("A1").Value = strReportTitle
            .Range("A1").Font.Size = 20
            .Rows("1:2").Font.Bold = True
            'moves cursor to first row
            .Range("A3").Select
        Else
            .Rows(1).Font.Bold = True
            .Range("A2").Select
        End If
        
        

        'Save Changes
        xlWB.Save
        xlWB.Close

    End With

    objExcel.Quit
    Set objExcel = Nothing
    
    FollowHyperlink strExportFileFull

Err_Exit:
    DoCmd.SetWarnings True
    'for testing
    ' objExcel.Quit
    ' Set objExcel = Nothing
    ' rsExport.Close
    ' Set rsExport = Nothing
    Exit Function

Err_Stop:
    MsgBox Err.Number & ": " & Err.Description
    Resume Err_Exit
    'Resume 'uncomment for testing

End Function