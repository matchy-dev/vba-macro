Attribute VB_Name = "Module1"
Const CTL_SH = "コントロール"
Const OUT_SH = "結果"
Const TBL_SH = "テーブル一覧"
Const DSN_ADD = "E3"
Const SQL_ADD = "B3"

Const DSN_CONF_SRV_TYPE = 2
Const DSN_CONF_DSN_NAME = 3
Const DSN_CONF_HOST = 4
Const DSN_CONF_PORT = 5
Const DSN_CONF_DB_NAME = 6
Const DSN_CONF_USER = 7
Const DSN_CONF_PASS = 8

Function get_srv_type()
    Dim show_name
    show_name = Worksheets(CTL_SH).Range(DSN_ADD).Value
    get_srv_type = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_SRV_TYPE, False)
End Function

Function make_dsn_str_org()
    Dim show_name
    show_name = Worksheets(CTL_SH).Range(DSN_ADD).Value
    
    Dim srv_type
    srv_type = get_srv_type()
    
    Dim dsn_name, host, port, db_name, user, pass
    dsn_name = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_DSN_NAME, False)
    host = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_HOST, False)
    port = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_PORT, False)
    db_name = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_DB_NAME, False)
    user = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_USER, False)
    pass = WorksheetFunction.VLookup(show_name, Range("dsn_conf"), DSN_CONF_PASS, False)
    
    Select Case srv_type
        Case "oracle"
            make_dsn_str_org = "DSN=" & dsn_name & ";UID=" & user & ";PWD=" & pass & ";DBQ=" & host & ":" & port & "/" & db_name
        Case "postgres"
            make_dsn_str_org = "DSN=" & dsn_name & ";DATABASE=" & db_name & ";SERVER=" & host & ";PORT=" & port & ";UID=" & user & ";PWD=" & pass
        Case "sqlite"
            make_dsn_str_org = "DSN=" & dsn_name & ";Database=" & db_name
    End Select
    Debug.Print srv_type & " : " & make_dsn_str_org
End Function

Function make_dsn_str()
    make_dsn_str = "ODBC;" & make_dsn_str_org
End Function

Function get_sql_str()
    Dim base As Range
    Set base = ThisWorkbook.Worksheets(CTL_SH).Range(SQL_ADD)
    
    get_sql_str = get_sql_str_by_add(base)
End Function

Function get_sql_str_by_add(base As Range)
    get_sql_str_by_add = ""
    If base.Value <> "" Then
        get_sql_str_by_add = base.Value
        If base.Offset(1, 0).Value <> "" Then
            Dim st_row, last_row, row_num
            st_row = base.Row
            last_row = base.End(xlDown).Row
            row_num = last_row - st_row
            
            Dim r
            For r = 1 To row_num
                get_sql_str_by_add = get_sql_str_by_add & vbCrLf & base.Offset(r, 0).Value
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

Sub exec_this_sql()
    Dim dsn_str, sql
    dsn_str = make_dsn_str
    sql = get_sql_str_by_add(ActiveCell)
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

Sub get_table_list()
    Dim srv_type, sql
    srv_type = get_srv_type()
    Select Case srv_type
        Case "oracle"
            sql = "select table_name from user_tables order by 1"
        Case "postgres"
            sql = "select relname from pg_stat_user_tables order by 1"
        Case "sqlite"
            sql = "select name from sqlite_master where type='table' order by 1"
    End Select
    
    Call select_sql(TBL_SH, make_dsn_str(), sql)
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

