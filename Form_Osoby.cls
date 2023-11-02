VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Osoby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btn_addPerson_Click()
    If CheckFields("osoby") = False Then
        Person "Add"
    Else
        Forms![osoby]![lbl_msg].Caption = "Wprowadz dane do wszystkich pól."
    End If
End Sub

Private Sub btn_close_Click()
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSavePrompt
End Sub

Private Sub btn_main_Click()
    DoCmd.OpenForm "Main"
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSavePrompt
End Sub

Private Sub btn_modifyPerson_Click()
    If CheckFields("osoby") = False Then
        Person "Modify"
    Else
        Forms![osoby]![lbl_msg].Caption = "Wprowadz dane do wszystkich pól."
    End If
End Sub

Private Sub lst_persons_Click()
    SelectPerson Forms!osoby!lst_persons.Value
End Sub

Private Sub btn_deletePerson_Click()
    If Forms!osoby!lst_persons.Value <> "" Then
        DeletePerson Forms!osoby!lst_persons.Value
    End If
End Sub
