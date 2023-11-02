Attribute VB_Name = "mMenu"
Option Compare Database
Option Explicit

Sub Menu(ByVal parameter As String)
Dim sql As String
Dim rs As DAO.Recordset
Dim isrecordexists As Boolean

On Error GoTo errhandling
    Set rs = CurrentDb.OpenRecordset("Szablony_Dzienne")
    
    Do Until rs.EOF
        If rs.Fields("Nazwa_Szablonu_Dziennego") = Forms![MenuDzienne]![txt_MenuName] Then
            isrecordexists = True
        End If
        rs.MoveNext
    Loop
    
    Select Case parameter
        Case "Add"
            If isrecordexists = False Then
                sql = "INSERT INTO Szablony_Dzienne ([Nazwa_Szablonu_Dziennego]) " & _
                        "VALUES ('" & Forms![MenuDzienne]![txt_MenuName] & "');"
                CurrentDb.Execute sql
                Forms![MenuDzienne]![lbl_msg].Caption = "Wprowadzono nowe menu do bazy." & vbNewLine & "Wprowadz dane i zatwierdŸ przyciskiem ""DODAJ""."
            Else
                Forms![MenuDzienne]![lbl_msg].Caption = "Przepis juz istnieje."
            End If
    End Select
    Forms!MenuDzienne!lst_menus.Requery
    ClearFields "MenuDzienne"
    ClearLists "MenuDzienne"
    Forms![MenuDzienne]![txt_MenuName] = ""
    
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If

errhandling:
        
    If Not rs Is Nothing Then
    rs.Close
    Set rs = Nothing
    End If
    
    If Err.Number = 3061 Or Err.Number = 3075 Then
        Forms![produkty]![lbl_msg].Caption = "Pola oznaczone na czerwono wymagaj¹ wartoœci numerycznych."
        iscontrolnumeric "produkty"
    End If
End Sub

Sub deleteMenu(ID As Long)
    CurrentDb.Execute ("DELETE FROM Szablony_Dzienne WHERE ID = " & ID & ";")
    ClearFields "MenuDzienne"
    Forms![MenuDzienne]![txt_MenuName] = ""
    Forms!MenuDzienne!lst_menus.Requery
    Forms![MenuDzienne]![lbl_msg].Caption = "Usuniêto menu."
End Sub

Sub SelectMenu(ID As Long)
Dim rs As DAO.Recordset
Dim sql As String

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM Szablony_Dzienne WHERE ID = " & ID & ";")
    
    Forms!MenuDzienne!txt_MenuName = rs.Fields("Nazwa_Szablonu_Dziennego")
        
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
     
End Sub

'Sub RefreshIngridientsLists()
'    Forms!MenuDzienne!lst_availableproducts.Requery
'    Forms!MenuDzienne!lst_MenuProducts.Requery
'End Sub
'
'Sub addingridient()
'Dim sql As String
'
'    ClearFields "Szablony_Dzienne"
'    ClearLists "Szablony_Dzienne"
'
'    If Forms!MenuDzienne!lst_availableproducts.Value <> "" And Forms!MenuDzienne!lst_menus.Value <> "" And Forms!MenuDzienne!txt_quantity.Value <> "" Then
'        sql = "INSERT INTO Zestawienie_Produktow ([ID_Produktu], [Ilosc_Produktu], [ID_Przepisu]) " & _
'                "VALUES (" & Forms!MenuDzienne!lst_availableproducts.Value & ", " & Forms![MenuDzienne]![txt_quantity] & ", " & Forms!MenuDzienne!lst_menus.Value & ");"
'        CurrentDb.Execute sql
'        Forms!MenuDzienne!lst_availableproducts.Value = ""
'        Forms![MenuDzienne]![txt_quantity].Value = ""
'        Forms!MenuDzienne!lbl_msg.Caption = " "
'        RefreshIngridientsLists
'    Else
'        Forms!MenuDzienne!lbl_msg.Caption = "Zaznacz wymagane pola lub/oraz wprowadŸ dane."
'        Forms!MenuDzienne!lst_availableproducts.BorderColor = RGB(220, 0, 0)
'        Forms!MenuDzienne!lst_menus.BorderColor = RGB(220, 0, 0)
'        Forms!MenuDzienne!txt_quantity.BorderColor = RGB(220, 0, 0)
'    End If
'End Sub
'
'Sub deleteingridient()
'Dim sql As String
'
'    ClearFields "Szablony_Dzienne"
'    ClearLists "Szablony_Dzienne"
'
'    If Forms!MenuDzienne!lst_MenuProducts.Value <> "" Then
'        sql = "DELETE FROM Zestawienie_Produktow WHERE ID_Przepisu = " & Forms!MenuDzienne!lst_menus.Value & " AND ID_Produktu = " & Forms!MenuDzienne!lst_MenuProducts.Value & ";"
'        CurrentDb.Execute sql
'        Forms!MenuDzienne!lst_MenuProducts.Value = ""
'        Forms!MenuDzienne!lbl_msg.Caption = " "
'        RefreshIngridientsLists
'    Else
'        Forms!MenuDzienne!lbl_msg.Caption = "Zaznacz sk³adnik."
'        Forms!MenuDzienne!lst_MenuProducts.BorderColor = RGB(220, 0, 0)
'    End If
'
'End Sub
'
'Sub deleteingridients()
'Dim sql As String
'
'    If Forms!MenuDzienne!lst_MenuProducts.ListCount > 1 Then
'        ClearFields "Szablony_Dzienne"
'        ClearLists "Szablony_Dzienne"
'
'        If MsgBox("Czy chcesz usun¹æ wszsytkie sk³adniki z listy produktów dla przepisu?", vbYesNo, "Are you sure?") = vbYes Then
'            sql = "DELETE FROM Zestawienie_Produktow WHERE ID_Przepisu = " & Forms!MenuDzienne!lst_menus.Value & ";"
'            CurrentDb.Execute sql
'            Forms!MenuDzienne!lst_MenuProducts.Value = ""
'            Forms!MenuDzienne!lbl_msg.Caption = " "
'            RefreshIngridientsLists
'        Else
'            Forms!MenuDzienne!lbl_msg.Caption = "Zaznacz sk³adnik."
'            Forms!MenuDzienne!lst_MenuProducts.BorderColor = RGB(220, 0, 0)
'        End If
'    End If
'End Sub




