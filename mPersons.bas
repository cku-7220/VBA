Attribute VB_Name = "mPersons"
Option Compare Database
Option Explicit

Sub Person(ByVal parameter As String)
Dim sql As String
Dim rs As DAO.Recordset
Dim isrecordexists As Boolean

On Error GoTo errhandling

    Set rs = CurrentDb.OpenRecordset("Osoby")
    
    Do Until rs.EOF
        If rs.Fields("Imie_Osoby") = Forms![osoby]![txt_Name] Then
            isrecordexists = True
        End If
        rs.MoveNext
    Loop
    
    Select Case parameter
        Case "Add"
            If isrecordexists = False Then
                sql = "INSERT INTO Osoby ([Imie_Osoby], [Dzienne_Zapotrzebowanie_Na_Kal]) " & _
                        "VALUES ('" & Forms![osoby]![txt_Name] & "', " & dots(Forms![osoby]![txt_kcal]) & ");"
                CurrentDb.Execute sql
                Forms![osoby]![lbl_msg].Caption = "Wprowadzono now¹ osobê do bazy." & vbNewLine & "Wprowadz dane i zatwierdŸ przyciskiem ""DODAJ""."
            Else
                Forms![osoby]![lbl_msg].Caption = "Osoba juz istnieje."
            End If
        Case "Modify"
            If isrecordexists = True Then
                sql = "UPDATE osoby " & _
                        "SET [Dzienne_Zapotrzebowanie_Na_Kal] = " & dots(Forms![osoby]![txt_kcal]) & " " & _
                        "WHERE [Imie_Osoby] = '" & Forms![osoby]![txt_Name] & "';"
                CurrentDb.Execute sql
                Forms![osoby]![lbl_msg].Caption = "Zmodyfikowano dane osoby." & vbNewLine & "Wprowadz dane i zatwierdŸ przyciskiem ""DODAJ""."
            Else
                Forms![osoby]![lbl_msg].Caption = "Podana osoba nie istnieje."
            End If
    End Select
    Forms!osoby!lst_persons.Requery
    Forms!osoby!txt_Name = ""
    Forms!osoby!txt_kcal = ""
    ClearFields "osoby"
    ClearLists "osoby"
    Forms!osoby!txt_Name.SetFocus
    
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
        Forms![osoby]![lbl_msg].Caption = "Pola oznaczone na czerwono wymagaj¹ wartoœci numerycznych."
        iscontrolnumeric "osoby"
    End If

End Sub

Sub SelectPerson(ID As Long)
Dim rs As DAO.Recordset

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM osoby WHERE ID = " & ID & ";")
    
    Forms!osoby!txt_Name = rs.Fields("Imie_Osoby")
    Forms!osoby!txt_kcal = rs.Fields("Dzienne_Zapotrzebowanie_Na_Kal")
        
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
End Sub

Sub DeletePerson(ID As Long)
    CurrentDb.Execute ("DELETE FROM osoby WHERE ID = " & ID & ";")
    ClearFields "osoby"
    Forms!osoby!lst_persons.Requery
    Forms!osoby!txt_Name = ""
    Forms!osoby!txt_kcal = ""
    Forms![osoby]![lbl_msg].Caption = "Usuniêto dane osoby."
End Sub

