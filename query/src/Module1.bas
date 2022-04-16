Attribute VB_Name = "Module1"
Const CTL_SH = "コントロール"
Const OUT_SH = "結果"
Const TBL_SH = "テーブル一覧"
Const DSN_ADD = "E4"
Const SQL_ADD = "B4"

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
        Case "mysql"
            make_dsn_str_org = "DSN=" & dsn_name & ";SERVER=" & host & ";UID=" & user & ";PWD=" & pass & ";DATABASE=" & db_name & ";PORT=" & port
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

Function get_table_name()
    Dim table_name
    table_name = UCase(get_sql_str)
    table_name = Replace(table_name, vbTab, " ")
    table_name = Replace(table_name, vbCrLf, " ")
    table_name = Replace(table_name, vbLf, " ")
    
    Dim idx
    idx = InStr(1, table_name, " FROM ")
    table_name = Trim(Mid(table_name, idx + 5))
    
    idx = InStr(1, table_name, " ")
    If idx > 0 Then
        table_name = Left(table_name, idx - 1)
    End If
    
    get_table_name = table_name
End Function

Function get_sql_str_with_id(srv_type, in_sql)
    Dim sql
    sql = UCase(in_sql)
    sql = Replace(sql, vbTab, " ")
    sql = Replace(sql, vbCrLf, " ")
    sql = Replace(sql, vbLf, " ")
    
    Dim idx
    idx = InStr(1, sql, " FROM ")
    
    Dim head, tail
    head = Left(sql, idx - 1)
    tail = Mid(sql, idx)
    
    Dim add_col
    Select Case srv_type
        Case "oracle"
            idx = InStr(head, "*")
            If idx > 0 Then
                head = Replace(head, "*", get_table_name() & ".*")
            End If
            add_col = "ROWID"
        Case "postgres"
            add_col = "CAST(CTID AS TEXT) AS ROWID"
        Case "sqlite"
            add_col = "ROWID"
    End Select
    
    sql = head + ", " + add_col + " " + tail
    
    get_sql_str_with_id = sql
End Function

Function get_sel_rows()
    Dim sel_rows
    sel_rows = Selection.Address(False, False)
    Dim row_array
    row_array = Split(sel_rows, ",")
    Dim i, j, st_ed, out_row(), all_num
    all_num = 0
    For i = LBound(row_array) To UBound(row_array)
        st_ed = Split(row_array(i), ":")
        For j = st_ed(0) To st_ed(1)
            ReDim Preserve out_row(all_num)
            out_row(all_num) = j
            all_num = all_num + 1
        Next
    Next
    get_sel_rows = out_row
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
    If Worksheets(CTL_SH).CheckBox2.Value = True Then
        sql = get_sql_str_with_id(get_srv_type(), sql)
    End If
    
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

Sub insert_func()
    If MsgBox("登録しますか?", vbOKCancel, "insert") <> vbOK Then
        Exit Sub
    End If
    
    Dim table_name
    table_name = get_table_name()
    
    Dim ins_sql
    ins_sql = "INSERT INTO " & table_name & " ( "
    
    Dim last_col
    last_col = Range("A1").End(xlToRight).Column
    If UCase(Cells(1, last_col).Value) = "ROWID" Then
        last_col = last_col - 1
    End If
    
    Dim c
    For c = 1 To last_col
        If c > 1 Then
            ins_sql = ins_sql & ", "
        End If
        ins_sql = ins_sql & Cells(1, c).Value
    Next
    ins_sql = ins_sql & " ) VALUES ( "
    
    On Error GoTo err_line
    
    Dim con As New ADODB.Connection
    con.Open make_dsn_str_org()
    
    Dim sql
    For Each r In get_sel_rows()
        sql = ins_sql
        For c = 1 To last_col
            If c > 1 Then
                sql = sql & ", "
            End If
            sql = sql & "'" & Cells(r, c).Value & "'"
        Next
        sql = sql & " )"
        Debug.Print r & " " & sql
        Call con.Execute(sql)
    Next
    
    con.Close
    
    Call MsgBox("完了", vbOKOnly, "insert")
       
    Exit Sub
    
err_line:
    
    MsgBox Err.Description, vbOKOnly, "Error"
End Sub

Sub delete_func()
    If MsgBox("削除しますか?", vbOKCancel, "delete") <> vbOK Then
        Exit Sub
    End If
    
    Dim row_col
    row_col = Range("A1").End(xlToRight).Column
    If UCase(Cells(1, row_col).Value) <> "ROWID" Then
        Call MsgBox("「更新用IDを取得」をチェックして再検索してください", vbOKOnly, "")
        Exit Sub
    End If
    
    Dim table_name
    table_name = get_table_name()
    
    Dim srv_type, row_id_col_name
    srv_type = get_srv_type()
    Select Case srv_type
        Case "oracle"
            row_id_col_name = "ROWID"
        Case "postgres"
            row_id_col_name = "CTID"
        Case "sqlite"
            row_id_col_name = "ROWID"
    End Select
    
    On Error GoTo err_line
    
    Dim con As New ADODB.Connection
    con.Open make_dsn_str_org()
    
    Dim sql
    For Each r In get_sel_rows()
        sql = "DELETE FROM " & table_name & " WHERE " & row_id_col_name & " = '" & Cells(r, row_col).Value & "'"
        Debug.Print r & " " & sql
        Call con.Execute(sql)
    Next
    
    con.Close
    
    Call MsgBox("完了", vbOKOnly, "delete")
       
    Exit Sub
    
err_line:
    
    MsgBox Err.Description, vbOKOnly, "Error"
End Sub

Sub update_func()
    If MsgBox("更新しますか?", vbOKCancel, "update") <> vbOK Then
        Exit Sub
    End If
    
    Dim row_col
    row_col = Range("A1").End(xlToRight).Column
    If UCase(Cells(1, row_col).Value) <> "ROWID" Then
        Call MsgBox("「更新用IDを取得」をチェックして再検索してください", vbOKOnly, "")
        Exit Sub
    End If
    
    Dim table_name
    table_name = get_table_name()
    
    Dim srv_type, row_id_col_name
    srv_type = get_srv_type()
    Select Case srv_type
        Case "oracle"
            row_id_col_name = "ROWID"
        Case "postgres"
            row_id_col_name = "CTID"
        Case "sqlite"
            row_id_col_name = "ROWID"
    End Select
    
    On Error GoTo err_line
    
    Dim con As New ADODB.Connection
    con.Open make_dsn_str_org()
    
    Dim sql_head, sql
    sql_head = "UPDATE " & table_name & " SET "
    For Each r In get_sel_rows()
        sql = sql_head
        For c = 1 To row_col - 1
            If c > 1 Then
                sql = sql & ", "
            End If
            sql = sql & Cells(1, c).Value & " = '" & Cells(r, c).Value & "'"
        Next
        sql = sql & " WHERE " & row_id_col_name & " = '" & Cells(r, row_col).Value & "'"
        Debug.Print r & " " & sql
        Call con.Execute(sql)
    Next
    
    con.Close
    
    Call MsgBox("完了", vbOKOnly, "delete")
       
    Exit Sub
    
err_line:
    
    MsgBox Err.Description, vbOKOnly, "Error"
End Sub

Sub get_table_list()
    Dim srv_type, sql
    srv_type = get_srv_type()
    Select Case srv_type
        Case "oracle"
            sql = "select table_name from user_tables order by 1"
        Case "postgres"
            sql = "select relname from pg_stat_user_tables order by 1"
        Case "mysql"
            sql = "select table_name from information_schema.tables where table_schema = (select database()) and table_type = 'base table' order by 1"
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

