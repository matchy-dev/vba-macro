VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Call button_click
End Sub

Private Sub CommandButton2_Click()
    Call check_sql
End Sub

Private Sub CommandButton3_Click()
    Call get_table_list
End Sub
