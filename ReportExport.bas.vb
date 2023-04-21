Public Function ReportExport(strSQL As String, strDir As String, strExportFile As String)


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
    'Dim oFS As Object
    'Dim dteDateLastModified As Date
    'Dim strExistingFile As String
    'Dim intRSCount As Long
    

        
        
    ' confirm import file exists
    If Dir(strDir) = "" Then
        MsgBox "Unable to locate directory! Please confirm directory is available: " & strDir
        GoTo Err_Exit
    End If
    
    ' additional code if you need this func to also read an existing file or previous export file
    ' create an instance of the MS Scripting Runtime FileSystemObject class
    'Set oFS = CreateObject("Scripting.FileSystemObject")

    ' get 'Date modified' info from existing file
    'dteDateLastModified = oFS.GetFile(strExistingFile).Datelastmodified

    ' If dteDateLastModified < Date Then
    '     MsgBox "Please ensure file has been updated." & vbNewLine & strExistingFile
    '     GoTo Err_Exit
    ' End If

    DoCmd.SetWarnings False
    
    Set db = CurrentDb

    Set rsExport = db.OpenRecordset(strSQL, dbOpenDynaset)

    strExportFileFull = strDir & strExportFile & "_" & Format(Now, "YYYYMMDD_hhmmss") & ".xlsx"


    'FORMAT EXCEL TEMPLATE
    Set objExcel = CreateObject("Excel.Application")
    
    With objExcel
        .DisplayAlerts = False ' True for testing
        Set xlWB = .Workbooks.Open(strExportFileFull)
        Set xlSheet = xlWB.Sheets(strExportFile)
    End With
    
    ' uncomment for testing
    ' objExcel.Visible = True
    
    ' allows func to work with any column count, can be made longer if needed
    ' split creates 0-based index array
    aryCols = Split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z", ",")
    
    ' if record count is needed, but it shouldn't be
    ' rsExport.MoveLast
    ' intRSCount = rsExport.RecordCount
    ' rsExport.MoveFirst

    ' stores rsExport as a 0-based 2D array (columns, rows)
    ' just the records, not field names
    aryRS = rsExport.GetRows
    ' gets column count of array
    intFieldCount = UBound(aryRS, 1)

    'populate column/field names
    For j = LBound(aryRS, 1) To UBound(aryRS, 1)
        intCol = 0
        xlSheet.Range(aryCols(j) & "1").Value = rsExport.Fields(intCol)
        intCol = intCol + 1
    Next j

    'populate the rest of the sheet
    For j = LBound(aryRS, 1) To UBound(aryRS, 1)
        intRow = 2
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
  
  
        'to insert totals at  bottom of spreadsheet
        ' .Range("B" & intLastRow + 2).Formula = "Total # of ______:"
        ' .Range("C" & intLastRow + 2).Formula = "=COUNTA(B2:B" & intLastRow & ")"


        'Format Report Columns
        For col = 0 To intFieldCount
            .Range(aryCols(col) & "2", aryCols(col) & intLastRow).HorizontalAlignment = xlCenter
            If IsDate(.Range(aryCols(col) & "2").Value) Then
                .Range(aryCols(col) & "2", aryCols(col) & intLastRow).HorizontalAlignment = xlRight
                .Range(aryCols(col) & "2", aryCols(col) & intLastRow).NumberFormat = "mm/dd/yyyy"
            Elseif IsNumeric(.Range(aryCols(col) & "2").Value) Then
                .Range(aryCols(col) & "2", aryCols(col) & intLastRow).HorizontalAlignment = xlRight
                .Range(aryCols(col) & "2", aryCols(col) & intLastRow).NumberFormat = "00000000-00" ' change as needed
            Elseif Left(.Range(aryCols(col) & "2").Value, 1) = "$" Then
                .Range(aryCols(col) & "2", aryCols(col) & intLastRow).HorizontalAlignment = xlRight
                .Range(aryCols(col) & "2", aryCols(col) & intLastRow).NumberFormat = "$#,##0.00"
            End If
        Next col


'       ' use .TextToColumns to clean up messy string data
'        If IsEmpty("D2") = False Then
'            .Range("C2", "C" & intLastRow).TextToColumns
'        End If

        'Format Report Borders
        .Range(aryCols(0) & "1", aryCols(intFieldCount) & intLastRow).Borders(xlEdgeLeft).Weight = xlThin
        .Range(aryCols(0) & "1", aryCols(intFieldCount) & intLastRow).Borders(xlEdgeRight).Weight = xlThin
        .Range(aryCols(0) & "1", aryCols(intFieldCount) & intLastRow).Borders(xlEdgeBottom).Weight = xlThin
        .Range(aryCols(0) & "1", aryCols(intFieldCount) & intLastRow).Borders(xlEdgeTop).Weight = xlThin
        .Range(aryCols(0) & "1", aryCols(intFieldCount) & intLastRow).Borders(xlInsideVertical).Weight = xlThin
        .Range(aryCols(0) & "1", aryCols(intFieldCount) & intLastRow).Borders(xlInsideHorizontal).Weight = xlThin
   
    
        ' Format Totals Fields
        ' With .Range("B" & intLastRow + 2, "C" & intLastRow + 2)
        '     .Font.Bold = True
        '     .Borders(xlEdgeTop).Weight = xlMedium
        '     .Borders(xlEdgeBottom).Weight = xlMedium
        ' End With
        
        
        'Format Column Headings
        .Rows("1:1").RowHeight = 31.5
        
        With .Range(aryCols(0) & "1", aryCols(intFieldCount) & "1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.6
            .PatternTintAndShade = 0
        End With
        With .Range(aryCols(0) & "1", aryCols(intFieldCount) & "1")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
        End With
        
        'Hide Columns
        '.Range("M:N").EntireColumn.Hidden = True
        
        'Freeze Top Row
        .Range("A1").Select
        objExcel.ActiveWindow.SplitRow = 1
        objExcel.ActiveWindow.FreezePanes = True
        


        'Set Auto Filter
            .AutoFilterMode = False
            .Range(aryCols(0) & ":" & aryCols(intFieldCount)).AutoFilter
        
            'Auto Fit Columns and Rows
            .Range(aryCols(0) & ":" & aryCols(intFieldCount)).ColumnWidth = 80
            .Columns(aryCols(0) & ":" & aryCols(intFieldCount)).EntireColumn.AutoFit
            .Rows("2:" & intLastRow).EntireRow.AutoFit
      
        
        'Configure Page Setup
        objExcel.PrintCommunication = False
    
        With .PageSetup
            .PrintTitleRows = "$1:$1"
            'customize main title as needed, in addition to report title
            .CenterHeader = "&""-,Bold""&12MAIN TITLE " & strExportFile & " Report"
            'TO-DO: add ability to get date/date range from query and add to report title
            '.CenterHeader = "&""-,Bold""&12MAIN TITLE " & strExportFile & " Report" & vbNewLine & Date
            .CenterHorizontally = True
            .Orientation = xlLandscape
        End With
    
        On Error Resume Next
        objExcel.PrintCommunication = True
        
        'Save Changes
        xlWB.Save
        xlWB.Close
            
    End With

    objExcel.Quit
    Set objExcel = Nothing
    
    

    FollowHyperlink strExportFileFull

Err_Exit:
    DoCmd.SetWarnings True
    objExcel.Quit
    Set objExcel = Nothing
    rsExport.Close
    Set rsExport = Nothing
    Exit Function

Err_Stop:
    MsgBox Err.Number & ": " & Err.Description
    Resume Err_Exit
    Resume

End Function