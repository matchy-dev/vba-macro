Attribute VB_Name = "Module1"
Const DIR_ADD = "C3"

Sub export_code()
Attribute export_code.VB_ProcData.VB_Invoke_Func = "E\n14"
    out_dir = ThisWorkbook.Worksheets(1).Range(DIR_ADD).Value
    If Right(out_dir, 1) <> "\" Then
        out_dir = out_dir & "\"
    End If
    For Each component In ThisWorkbook.VBProject.VBComponents
        code_status = "NO_CODE"
        If component.codeModule.CountOfLines > 0 Then
            component.Export out_dir & component.Name & ".bas"
            code_status = "EXISTS"
        End If
        Debug.Print component.Type, component.Name, code_status
    Next
End Sub
