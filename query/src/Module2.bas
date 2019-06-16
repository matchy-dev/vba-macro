Attribute VB_Name = "Module2"
Sub add_command_bar(bar_type, menu_name, func_name)
    With Application.CommandBars(bar_type).Controls.Add()
        .Caption = menu_name
        .OnAction = func_name
    End With
End Sub

Sub del_command_bar(bar_type, menu_name)
    For Each bar In Application.CommandBars(bar_type).Controls
        If bar.Caption = menu_name Then
            Application.CommandBars(bar_type).Controls(menu_name).Delete
        End If
    Next
End Sub

Function get_def_data()
    get_def_data = Array( _
        Array("Cell", "Ç±ÇÃSQLÇé¿çs", "exec_this_sql"), _
        Array("Row", "insert", "insert_func"), _
        Array("Row", "delete", "delete_func"), _
        Array("Row", "update", "update_func") _
    )
End Function

Sub auto_open()
    Dim def_list
    def_list = get_def_data()
    
    Dim i
    For i = LBound(def_list) To UBound(def_list)
        Call add_command_bar(def_list(i)(0), def_list(i)(1), def_list(i)(2))
    Next
End Sub

Sub auto_close()
    Dim def_list
    def_list = get_def_data()
    For i = LBound(def_list) To UBound(def_list)
        Call del_command_bar(def_list(i)(0), def_list(i)(1))
    Next
End Sub
