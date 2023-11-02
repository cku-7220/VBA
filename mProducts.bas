Attribute VB_Name = "mProducts"
Option Compare Database
Option Explicit

Sub Product(ByVal parameter As String)
Dim sql As String
Dim rs As DAO.Recordset
Dim isrecordexists As Boolean

On Error GoTo errhandling

    Set rs = CurrentDb.OpenRecordset("Produkty")
    
    Do Until rs.EOF
        If rs.Fields("Nazwa_Produktu") = Forms![produkty]![txt_ProductName] Then
            isrecordexists = True
        End If
        rs.MoveNext
    Loop
    
    Select Case parameter
        Case "Add"
            If isrecordexists = False Then
                sql = "INSERT INTO Produkty ([Nazwa_Produktu], [kcal_na_100JM], [Weglowodany], [Bialka], [Tluszcz], [JM_ID], [Waga(g)_na_1JM]) " & _
                        "VALUES ('" & Forms![produkty]![txt_ProductName] & "', " & dots(Forms![produkty]![txt_kcal]) & ", " & dots(Forms![produkty]![txt_Carbohydrates]) & ", " & dots(Forms![produkty]![txt_proteins]) & ", " & dots(Forms![produkty]![txt_fat]) & ", " & Forms![produkty]![txt_um] & ", " & Forms![produkty]![txt_weight] & ");"
                CurrentDb.Execute sql
                Forms![produkty]![lbl_msg].Caption = "Wprowadzono nowy produkt do bazy." & vbNewLine & "Wprowadz dane i zatwierdü przyciskiem ""DODAJ""."
            Else
                Forms![produkty]![lbl_msg].Caption = "Produkt juz istnieje."
            End If
        Case "Modify"
            If isrecordexists = True Then
                sql = "UPDATE Produkty " & _
                        "SET [kcal_na_100JM] = " & dots(Forms![produkty]![txt_kcal]) & ", [Weglowodany] = " & dots(Forms![produkty]![txt_Carbohydrates]) & ", [Bialka] = " & dots(Forms![produkty]![txt_proteins]) & ", [Tluszcz] = " & dots(Forms![produkty]![txt_fat]) & ", [JM_ID] = " & Forms![produkty]![txt_um] & ", [Waga(g)_na_1JM] = " & dots(Forms![produkty]![txt_weight]) & " " & _
                        "WHERE [Nazwa_Produktu] = '" & Forms![produkty]![txt_ProductName] & "';"
                CurrentDb.Execute sql
                Forms![produkty]![lbl_msg].Caption = "Zmodyfikowano dane produktu." & vbNewLine & "Wprowadz dane i zatwierdü przyciskiem ""DODAJ""."
            Else
                Forms![produkty]![lbl_msg].Caption = "Podany produkt nie istnieje."
            End If
    End Select
    Forms!produkty!lst_products.Requery
    Forms!produkty!txt_ProductName = ""
    Forms!produkty!txt_kcal = ""
    Forms!produkty!txt_Carbohydrates = ""
    Forms!produkty!txt_proteins = ""
    Forms!produkty!txt_fat = ""
    Forms!produkty!txt_um = ""
    Forms!produkty!txt_weight = ""
    ClearFields "produkty"
    ClearLists "produkty"
    Forms!produkty!txt_ProductName.SetFocus
    
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
        Forms![produkty]![lbl_msg].Caption = "Pola oznaczone na czerwono wymagajπ wartoúci numerycznych."
        iscontrolnumeric "produkty"
    End If

End Sub

Sub SelectProduct(ID As Long)
Dim rs As DAO.Recordset

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM produkty WHERE ID = " & ID & ";")
    
    Forms!produkty!txt_ProductName = rs.Fields("nazwa_produktu")
    Forms!produkty!txt_kcal = rs.Fields("kcal_na_100JM")
    Forms!produkty!txt_Carbohydrates = rs.Fields("Weglowodany")
    Forms!produkty!txt_proteins = rs.Fields("Bialka")
    Forms!produkty!txt_fat = rs.Fields("Tluszcz")
    Forms!produkty!txt_um = rs.Fields("JM_ID")
    Forms!produkty!txt_weight = rs.Fields("Waga(g)_na_1JM")
    
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
End Sub

Sub DeleteProduct(ID As Long)
    CurrentDb.Execute ("DELETE FROM produkty WHERE ID = " & ID & ";")
    ClearFields "produkty"
    Forms!produkty!lst_products.Requery
    Forms!produkty!txt_ProductName = ""
    Forms!produkty!txt_kcal = ""
    Forms!produkty!txt_Carbohydrates = ""
    Forms!produkty!txt_proteins = ""
    Forms!produkty!txt_fat = ""
    Forms!produkty!txt_um = ""
    Forms!produkty!txt_weight = ""
    Forms![produkty]![lbl_msg].Caption = "UsuniÍto dane produktu."
End Sub

