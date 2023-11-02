VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Produkty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btn_addProduct_Click()
    If CheckFields("produkty") = False Then
        Product "Add"
    Else
        Forms![produkty]![lbl_msg].Caption = "Wprowadz dane do wszystkich pól."
    End If
End Sub

Private Sub btn_close_Click()
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSavePrompt
End Sub

Private Sub btn_main_Click()
    DoCmd.OpenForm "Main"
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSavePrompt
End Sub

Private Sub btn_modifyProduct_Click()
    If CheckFields("produkty") = False Then
        Product "Modify"
    Else
        Forms![produkty]![lbl_msg].Caption = "Wprowadz dane do wszystkich pól."
    End If
End Sub

Private Sub lst_products_Click()
    SelectProduct Forms!produkty!lst_products.Value
End Sub

Private Sub btn_delete_Click()
    If Forms!produkty!lst_products.Value <> "" Then
        DeleteProduct Forms!produkty!lst_products.Value
    End If
End Sub
