Attribute VB_Name = "Module1"
Const CTL_SH = "コントロール"
Const OUT_SH = "結果"
Const DSN_ADD = "E3"
Const SQL_ADD = "B3"

Const DSN_CONF_SRV_TYPE = 2
Const DSN_CONF_DSN_NAME = 3
Const DSN_CONF_HOST = 4
Const DSN_CONF_PORT = 5
Const DSN_CONF_DB_NAME = 6
Const DSN_CONF_USER = 7
Const DSN_CONF_PASS = 8

Function make_dsn_str_org()
    Dim show_name
    show_name = Worksheets(CTL_SH).Range(DSN_ADD).Value
    
    Dim srv_type
    srv_type = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_SRV_TYPE, False)
    Select Case srv_type
        Case "oracle"
            make_dsn_str_org = make_dsn_str_oracle
        Case "postgres"
            make_dsn_str_org = make_dsn_str_postgres
    End Select
    Debug.Print srv_type & " : " & make_dsn_str_org
End Function

Function make_dsn_str_oracle()
    Dim show_name
    show_name = Worksheets(CTL_SH).Range(DSN_ADD).Value
    
    Dim dsn_name, host, port, db_name, user, pass
    dsn_name = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_DSN_NAME, False)
    host = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_HOST, False)
    port = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_PORT, False)
    db_name = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_DB_NAME, False)
    user = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_USER, False)
    pass = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_PASS, False)
    
    make_dsn_str_oracle = "DSN=" & dsn_name & ";UID=" & user & ";PWD=" & pass & ";DBQ=" & host & ":" & port & "/" & db_name
End Function

Function make_dsn_str_postgres()
    Dim show_name
    show_name = Worksheets(CTL_SH).Range(DSN_ADD).Value
    
    Dim dsn_name, host, port, db_name, user, pass
    dsn_name = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_DSN_NAME, False)
    host = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_HOST, False)
    port = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_PORT, False)
    db_name = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_DB_NAME, False)
    user = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_USER, False)
    pass = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_PASS, False)
    
    make_dsn_str_postgres = "DSN=" & dsn_name & ";DATABASE=" & db_name & ";SERVER=" & host & ";PORT=" & port & ";UID=" & user & ";PWD=" & pass
End Function

Function make_dsn_str()
    make_dsn_str = "ODBC;" & make_dsn_str_org
End Function

Function get_sql_str()
    Dim base As Range
    Set base = ThisWorkbook.Worksheets(CTL_SH).Range("A1")
    
    get_sql_str = ""
    If base.Range("B3").Value <> "" Then
        get_sql_str = base.Range("B3").Value
        If base.Range("B4").Value <> "" Then
            Dim last_row
            last_row = base.Range("B3").End(xlDown).Row
            
            Dim r
            For r = 4 To last_row
                get_sql_str = get_sql_str & vbCrLf & base.Range("B" & r).Value
            
            Next
        End If
    End If
End Function

Sub set_title_vertical()
    Range(Range("A1"), Range("A1").End(xlToRight)).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = -90
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit
End Sub

Sub button_click()
    Dim dsn_str, sql
    dsn_str = make_dsn_str
    sql = get_sql_str
    Call select_sql(OUT_SH, dsn_str, sql)
End Sub

Sub select_sql(sh, dsn_str, sql)

    Worksheets(sh).Select
    
    
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If
    
    Cells.Delete Shift:=xlUp
        
    With ActiveSheet.QueryTables.Add(Connection:= _
        dsn_str _
        , Destination:=Range("A1"))
        .CommandText = sql
        .Name = "クエリ"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With

    If Worksheets(CTL_SH).CheckBox1.Value = True Then
        Call set_title_vertical
    End If
    
    Range("A1").Select
End Sub

Sub check_sql()
    Dim dsn_str, sql
    dsn_str = make_dsn_str_org
    sql = get_sql_str
    
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    On Error GoTo err_line
    
    con.Open dsn_str
       
    Set rs = con.Execute(sql)
    
    Exit Sub
    
err_line:
    
    MsgBox Err.Description, vbOKOnly, "Error"

End Sub

