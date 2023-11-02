VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Przepisy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btn_addRecipe_Click()
    If Forms!przepisy!txt_RecipeName <> "" And Forms!przepisy!txt_hyperlink <> "" Then
        Recipe "Add"
    Else
        Forms![przepisy]![lbl_msg].Caption = "Wprowadz dane do wszystkich pól."
        Forms!przepisy!txt_RecipeName.BorderColor = RGB(220, 0, 0)
        Forms!przepisy!txt_hyperlink.BorderColor = RGB(220, 0, 0)
    End If
End Sub

Private Sub btn_close_Click()
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSavePrompt
End Sub

Private Sub btn_deleteRecipe_Click()
    If Forms!przepisy!lst_recipes.Value <> "" Then
        deleteingridients
        DeleteRecipe Forms!przepisy!lst_recipes.Value
    End If
End Sub

Private Sub btn_deleteingridients_Click()
    deleteingridients
End Sub

Private Sub btn_main_Click()
    DoCmd.OpenForm "Main"
    DoCmd.Close ObjectType:=acForm, ObjectName:=Me.Name, Save:=acSavePrompt
End Sub

Private Sub btn_MofidyRecipe_Click()
    If Forms!przepisy!txt_RecipeName <> "" And Forms!przepisy!txt_hyperlink <> "" Then
        Recipe "Modify"
    Else
        Forms![przepisy]![lbl_msg].Caption = "Wprowadz dane do wszystkich pól."
        Forms!przepisy!txt_RecipeName.BorderColor = RGB(220, 0, 0)
        Forms!przepisy!txt_hyperlink.BorderColor = RGB(220, 0, 0)
    End If
End Sub

Private Sub lst_recipes_Click()
    SelectRecipe Forms!przepisy!lst_recipes.Value
End Sub

Private Sub btn_addingredient_Click()
    addingridient
End Sub

Private Sub btn_deleteingridient_Click()
    deleteingridient
End Sub


