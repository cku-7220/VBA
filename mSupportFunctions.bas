Attribute VB_Name = "mSupportFunctions"
Option Compare Database
Option Explicit

Function dots(EntryValue)
    dots = Replace(EntryValue, ",", ".")
End Function

Sub iscontrolnumeric(ByVal form_name As String)
Dim txt As Control
    
    For Each txt In Forms(form_name).Controls
        If TypeOf txt Is TextBox Then
            Select Case IsNumeric(txt)
                Case False
                    txt.BorderColor = RGB(220, 0, 0)
                Case True
                    txt.BorderColor = Default
            End Select
        End If
    Next txt
End Sub

Function CheckFields(form_name As String) As Boolean
    Dim txt As Control
    
    For Each txt In Forms(form_name).Controls
        If TypeOf txt Is TextBox Then
            If txt <> "" Then
                txt.BorderColor = Default
            Else
                txt.BorderColor = RGB(220, 0, 0)
                CheckFields = True
            End If
        End If
    Next txt
End Function

Sub ClearFields(form_name As String)
    Dim txt As Control
    
    For Each txt In Forms(form_name).Controls
        If TypeOf txt Is TextBox Then
            txt.BorderColor = Default
        End If
    Next txt
    Forms(form_name).Refresh
End Sub

Sub ClearLists(form_name As String)
    Dim lbl As Control
    
    For Each lbl In Forms(form_name).Controls
        If TypeOf lbl Is ListBox Then
            lbl.BorderColor = Default
        End If
    Next lbl
    Forms(form_name).Refresh
End Sub
