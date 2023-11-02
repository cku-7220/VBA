VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MenuDzienne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btn_addMenu_Click()
    If CheckFields("MenuDzienne") = False Then
        Menu "Add"
    Else
        Forms![MenuDzienne]![lbl_msg].Caption = "Wprowadz dane do wszystkich pól."
    End If
End Sub

Private Sub btn_close_Click()
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSavePrompt
End Sub

Private Sub btn_deleteMenu_Click()
    If Forms![MenuDzienne]!lst_menus.Value <> "" Then
        deleteMenu Forms![MenuDzienne]!lst_menus.Value
    End If
End Sub

Private Sub btn_modifyMenu_Click()
    If CheckFields("MenuDzienne") = False Then
        Menu "Modify"
    Else
        Forms![MenuDzienne]![lbl_msg].Caption = "Wprowadz dane do wszystkich pól."
    End If
End Sub

Private Sub btn_main_Click()
    DoCmd.OpenForm "Main"
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSavePrompt
End Sub

Private Sub lst_Menus_Click()
    SelectMenu Forms![MenuDzienne]!lst_menus.Value
End Sub
