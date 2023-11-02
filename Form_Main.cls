VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub btn_close_Click()
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSavePrompt
End Sub

Private Sub btn_DailyMenu_Click()
    DoCmd.OpenForm "MenuDzienne"
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSavePrompt
End Sub

Private Sub btn_persons_Click()
    DoCmd.OpenForm "osoby"
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSavePrompt
End Sub

Private Sub btn_products_Click()
    DoCmd.OpenForm "produkty"
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSavePrompt
End Sub

Private Sub btn_recipes_Click()
    DoCmd.OpenForm "przepisy"
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSavePrompt
End Sub
