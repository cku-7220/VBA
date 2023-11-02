Attribute VB_Name = "mRecipes"
Option Compare Database
Option Explicit

Sub Recipe(ByVal parameter As String)
Dim sql As String
Dim rs As DAO.Recordset
Dim isrecordexists As Boolean

    Set rs = CurrentDb.OpenRecordset("Przepisy")
    
    Do Until rs.EOF
        If rs.Fields("Nazwa_Przepisu") = Forms![przepisy]![txt_RecipeName] Then
            isrecordexists = True
        End If
        rs.MoveNext
    Loop
    
    Select Case parameter
        Case "Add"
            If isrecordexists = False Then
                sql = "INSERT INTO Przepisy ([Nazwa_Przepisu], [Lokalizacja_Przepisu]) " & _
                        "VALUES ('" & Forms![przepisy]![txt_RecipeName] & "', '" & Forms![przepisy]![txt_hyperlink] & "');"
                CurrentDb.Execute sql
                Forms![przepisy]![lbl_msg].Caption = "Wprowadzono nowy przepis do bazy." & vbNewLine & "Wprowadz dane i zatwierdü przyciskiem ""DODAJ""."
            Else
                Forms![przepisy]![lbl_msg].Caption = "Przepis juz istnieje."
            End If
        Case "Modify"
            If isrecordexists = True Then
                sql = "UPDATE Przepisy " & _
                        "SET [Nazwa_Przepisu] = '" & Forms![przepisy]![txt_RecipeName] & "', [Lokalizacja_Przepisu] = '" & Forms![przepisy]![txt_hyperlink] & "'" & _
                        "WHERE [Nazwa_Przepisu] = '" & Forms![przepisy]![txt_RecipeName] & "';"
                CurrentDb.Execute sql
                Forms![przepisy]![lbl_msg].Caption = "Zmodyfikowano dane przepisu." & vbNewLine & "Wprowadz dane i zatwierdü przyciskiem ""DODAJ""."
            Else
                Forms![przepisy]![lbl_msg].Caption = "Podany przepis nie istnieje."
            End If
    End Select
    Forms!przepisy!lst_recipes.Requery
    ClearFields "przepisy"
    ClearLists "przepisy"
    Forms![przepisy]![txt_RecipeName] = ""
    Forms![przepisy]![txt_hyperlink] = ""
    
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
End Sub

Sub DeleteRecipe(ID As Long)
    CurrentDb.Execute ("DELETE FROM przepisy WHERE ID = " & ID & ";")
    ClearFields "przepisy"
    Forms![przepisy]![txt_RecipeName] = ""
    Forms![przepisy]![txt_hyperlink] = ""
    Forms!przepisy!lst_recipes.Requery
    Forms![przepisy]![lbl_msg].Caption = "UsuniÍto przepis."
End Sub

Sub SelectRecipe(ID As Long)
Dim rs As DAO.Recordset
Dim sql As String

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM przepisy WHERE ID = " & ID & ";")
    
    Forms!przepisy!txt_RecipeName = rs.Fields("nazwa_przepisu")
    Forms!przepisy!txt_hyperlink = rs.Fields("Lokalizacja_Przepisu")
        
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
       
    sql = "SELECT 100 / SUM(Produkty.[Waga(g)_na_1JM] * Zestawienie_Produktow.Ilosc_Produktu) * SUM(Zestawienie_Produktow.Ilosc_Produktu * Produkty.kcal_na_100JM/100) as calories " & _
            "FROM Zestawienie_Produktow " & _
            "LEFT JOIN Produkty ON Zestawienie_Produktow.ID_Produktu = Produkty.ID " & _
            "WHERE Zestawienie_Produktow.ID_Przepisu = " & ID & ";"
            
            

    Set rs = CurrentDb.OpenRecordset(sql)
    Forms!przepisy!lbl_calories.Caption = Format(rs.Fields("calories"), "0") & " kcal"
    
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    
    RefreshIngridientsLists
End Sub

Sub RefreshIngridientsLists()
    Forms!przepisy!lst_availableproducts.Requery
    Forms!przepisy!lst_RecipeProducts.Requery
End Sub

Sub addingridient()
Dim sql As String
    
    ClearFields "przepisy"
    ClearLists "przepisy"
        
    If Forms!przepisy!lst_availableproducts.Value <> "" And Forms!przepisy!lst_recipes.Value <> "" And Forms!przepisy!txt_quantity.Value <> "" Then
        sql = "INSERT INTO Zestawienie_Produktow ([ID_Produktu], [Ilosc_Produktu], [ID_Przepisu]) " & _
                "VALUES (" & Forms!przepisy!lst_availableproducts.Value & ", " & Forms![przepisy]![txt_quantity] & ", " & Forms!przepisy!lst_recipes.Value & ");"
        CurrentDb.Execute sql
        Forms!przepisy!lst_availableproducts.Value = ""
        Forms![przepisy]![txt_quantity].Value = ""
        Forms!przepisy!lbl_msg.Caption = " "
        RefreshIngridientsLists
    Else
        Forms!przepisy!lbl_msg.Caption = "Zaznacz wymagane pola lub/oraz wprowadü dane."
        Forms!przepisy!lst_availableproducts.BorderColor = RGB(220, 0, 0)
        Forms!przepisy!lst_recipes.BorderColor = RGB(220, 0, 0)
        Forms!przepisy!txt_quantity.BorderColor = RGB(220, 0, 0)
    End If
End Sub

Sub deleteingridient()
Dim sql As String

    ClearFields "przepisy"
    ClearLists "przepisy"
    
    If Forms!przepisy!lst_RecipeProducts.Value <> "" Then
        sql = "DELETE FROM Zestawienie_Produktow WHERE ID_Przepisu = " & Forms!przepisy!lst_recipes.Value & " AND ID_Produktu = " & Forms!przepisy!lst_RecipeProducts.Value & ";"
        CurrentDb.Execute sql
        Forms!przepisy!lst_RecipeProducts.Value = ""
        Forms!przepisy!lbl_msg.Caption = " "
        RefreshIngridientsLists
    Else
        Forms!przepisy!lbl_msg.Caption = "Zaznacz sk≥adnik."
        Forms!przepisy!lst_RecipeProducts.BorderColor = RGB(220, 0, 0)
    End If
    
End Sub

Sub deleteingridients()
Dim sql As String

    If Forms!przepisy!lst_RecipeProducts.ListCount > 1 Then
        ClearFields "przepisy"
        ClearLists "przepisy"
        
        If MsgBox("Czy chcesz usunπÊ wszsytkie sk≥adniki z listy produktÛw dla przepisu?", vbYesNo, "Are you sure?") = vbYes Then
            sql = "DELETE FROM Zestawienie_Produktow WHERE ID_Przepisu = " & Forms!przepisy!lst_recipes.Value & ";"
            CurrentDb.Execute sql
            Forms!przepisy!lst_RecipeProducts.Value = ""
            Forms!przepisy!lbl_msg.Caption = " "
            RefreshIngridientsLists
        Else
            Forms!przepisy!lbl_msg.Caption = "Zaznacz sk≥adnik."
            Forms!przepisy!lst_RecipeProducts.BorderColor = RGB(220, 0, 0)
        End If
    End If
End Sub


