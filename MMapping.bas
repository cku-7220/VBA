Attribute VB_Name = "mMapping"
Option Explicit

Private mcnSQLServer                        As ADODB.Connection
Dim gbStatus                                As Boolean

Sub AllCCDownload(ByVal gsCAP_REPORT_TITLE As String, ByVal gbStatus As Boolean)
    Dim sQuery                              As String
    Dim iError                              As Integer
    Dim i                                   As Long
    Dim lItem                               As Long
    Dim lRowOffset                          As Long
    Dim rngStartCell                        As Range
    Dim lColumnOffset                       As Long
    Dim sLi                                 As Object
    Dim sCommand                            As String
    Dim lRow                                As Long
    Dim rsCCList                            As ADODB.Recordset
    
    Application.ScreenUpdating = False

    On Error GoTo ErrHandler
    
    If gbMSG_RUAdd = True Then
        Exit Sub
    End If
    
    With fMapping.ListView2.ListItems
        If Not .Count = 0 Then
            .Clear
        End If
    End With
    
    If gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_SK Then
    
        If gbStatus = True Then
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry"
                
        ElseIf gbStatus = False Then
            If gsCAP_REPORT_TITLE_Level_2 = gsCAP_REPORT_TITLE_SKwoJO Then
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry"
                
            ElseIf gsCAP_REPORT_TITLE_Level_2 = "" Then
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry" & _
                    IIf(Len(fMapping.txtNumber1.Value) = 0, "", " AND [CC].[ColumnName] LIKE '" & fMapping.txtNumber1.Value & "' ") & _
                    IIf(Len(fMapping.txtDesc1.Value) = 0, "", " AND [CC].[ColumnName] LIKE '" & fMapping.txtDesc1.Value & "' ") & _
                    IIf(Len(fMapping.txtNumber2.Value) = 0, "", " AND [CRC].[ColumnName] LIKE '" & fMapping.txtNumber2.Value & "' ") & _
                    IIf(Len(fMapping.txtDesc2.Value) = 0, "", " AND [CRC].[ColumnName] LIKE '" & fMapping.txtDesc2.Value & "' ") & _
                    IIf(Len(fMapping.txtNumber3.Value) = 0, "", " AND [CRC].[ColumnName] LIKE '" & fMapping.txtNumber3.Value & "' ") & _
                    IIf(Len(fMapping.txtDesc3.Value) = 0, "", " AND [CRC].[ColumnName] LIKE '" & fMapping.txtDesc3.Value & "' ") & _
                    IIf(Len(fMapping.txtEGCode.Value) = 0, "", " AND [CS].[ColumnName] LIKE '" & fMapping.txtEGCode.Value & "' ") & _
                    IIf(Len(fMapping.txtCrop.Value) = 0, "", " AND [CRC].[ColumnName] LIKE '" & fMapping.txtCrop.Value & "' ") & _
                    IIf(fMapping.cbMap.Text = "Wszystkie", "", IIf(fMapping.cbMap.Text = "Niezmapowane", "AND [CRC].[ColumnName] IS NULL", "AND [CRC].[ColumnName] IS NOT NULL")) & _
                    " ORDER BY [CC].[ColumnName] "
            End If
            'Debug.Print sQuery
        End If
    ElseIf gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_RC Then
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry" & _
            IIf(Len(fMapping.txtNumber1.Value) = 0, "", " AND [RC].[ColumnName] LIKE '" & fMapping.txtNumber1.Value & "' ") & _
            IIf(Len(fMapping.txtDesc1.Value) = 0, "", " AND [RC].[ColumnName] LIKE '" & fMapping.txtDesc1.Value & "' ") & _
            IIf(Len(fMapping.txtEGCode.Value) = 0, "", " AND [CS].[ColumnName] LIKE '" & fMapping.txtEGCode.Value & "' ") & _
            " ORDER BY [ColumnName] "
    ElseIf gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_RU Then
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry" & _
            IIf(Len(fMapping.txtNumber1.Value) = 0, "", " AND [RU].[ColumnName] LIKE '" & fMapping.txtNumber1.Value & "' ") & _
            IIf(Len(fMapping.txtDesc1.Value) = 0, "", " AND [RU].[ColumnName] LIKE '" & fMapping.txtDesc1.Value & "' ") & _
            IIf(Len(fMapping.txtNumber2.Value) = 0, "", " AND [RU].[ColumnName] LIKE '" & fMapping.txtNumber2.Value & "' ") & _
            IIf(Len(fMapping.txtDesc2.Value) = 0, "", " AND [RU].[ColumnName] LIKE '" & fMapping.txtDesc2.Value & "' ") & _
            IIf(Len(fMapping.txtEGCode.Value) = 0, "", " AND [CS].[ColumnName] LIKE '" & fMapping.txtEGCode.Value & "' ") & _
            IIf(Len(fMapping.txtCrop.Value) = 0, "", " AND [RU].[ColumnName] LIKE '" & fMapping.txtCrop.Value & "' ")
    ElseIf gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_SKwoJO Then
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry"
            'Debug.Print sQuery
    End If
    
    'Debug.Print Now()
    'Debug.Print sQuery

    If sQuery <> "" Then
        Set mcnSQLServer = New ADODB.Connection
        mcnSQLServer.CursorLocation = adUseClient
        mcnSQLServer.ConnectionString = gsCONNECTION_STRING_ODYSSEY
        mcnSQLServer.Open
        Set rsCCList = New ADODB.Recordset
        rsCCList.Open sQuery, mcnSQLServer, adOpenDynamic, adLockOptimistic

        With fMapping.ListView2
            .AllowColumnReorder = False
            .CheckBoxes = True
            .FullRowSelect = True
            .MultiSelect = False
            .View = 3
            .Gridlines = True
            With .ListItems
                
            End With
            With .ColumnHeaders
                .Clear
                .Add Text:="  ", Width:=20
                For lItem = 1 To rsCCList.Fields.Count
                    If lItem < rsCCList.Fields.Count Then
                        .Add , , rsCCList.Fields(lItem - 1).name, Width:=130
                    ElseIf lItem = rsCCList.Fields.Count Then
                        .Add , , rsCCList.Fields(lItem - 1).name, Width:=5
                    End If
                Next lItem
            End With
            
            With rsCCList
                If .BOF And .EOF Then
                    MsgBox "Zbiór jest pusty."
                    Exit Sub
                End If
            End With
            
            lRowOffset = 0
            lColumnOffset = rsCCList.Fields.Count
    
            rsCCList.MoveFirst
            With rngStartCell
                While Not rsCCList.EOF
                    ReDim arr(CLng(rsCCList.RecordCount), lColumnOffset)
                    For lItem = 0 To lColumnOffset - 1
                        If lItem = 0 Then
                            Set sLi = fMapping.ListView2.ListItems.Add()
                            If IsNull(rsCCList.Fields(lItem).Value) Then
                                sLi.ListSubItems.Add , , "-"
                                arr(lRowOffset, lItem) = "-"
                            Else
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                'Set sLI = ListView2.ListItems.Add(, , rsCCList.Fields(lItem).Value)
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                                'ActiveSheet.Cells(lRowOffset + 1, lItem + 1) = rsCCList.Fields(lItem).Value
                            End If
                            GoTo next_iteration
                        End If
                        
                        Select Case rsCCList.Fields(lItem).Type
                        Case 1
                            ' SQL char type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 2
                            ' SQL int type
                            If IsNull(rsCCList.Fields(lItem).Value) Then
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            Else
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            End If
                        Case 3
                            ' SQL float type
                            If IsNull(rsCCList.Fields(lItem).Value) Then
                                sLi.ListSubItems.Add , , "-"
                                arr(lRowOffset, lItem) = "-"
                            Else
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            End If
                        Case 6
                            ' SQL image type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 7
                            ' SQL bit type
                            If IsNull(rsCCList.Fields(lItem).Value) Then
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            Else
                                sLi.ListSubItems.Add , , CInt(rsCCList.Fields(lItem).Value)
                                arr(lRowOffset, lItem) = CInt(rsCCList.Fields(lItem).Value)
                            End If
                        Case 9
                            ' SQL Date/time
                            If IsNull(rsCCList.Fields(lItem).Value) Then
                                sLi.ListSubItems.Add , , "-"
                                arr(lRowOffset, lItem) = "-"
                            Else
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            End If
                        Case 10
                            ' SQL uniqueidentifier type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case Else
                            ' Other
                            If IsNull(rsCCList.Fields(lItem).Value) Then
                                sLi.ListSubItems.Add , , "-"
                                arr(lRowOffset, lItem) = "-"
                                'ActiveSheet.Cells(lRowOffset + 1, lItem + 1) = "-"
                            Else
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                                'ActiveSheet.Cells(lRowOffset + 1, lItem + 1) = rsCCList.Fields(lItem).Value
                            End If
                        End Select
next_iteration:
                    Next lItem
                    If lRowOffset = 1000 Then GoTo SubExit
                    lRowOffset = lRowOffset + 1
                    rsCCList.MoveNext
                Wend
            End With
        End With
               
SubExit:
        On Error Resume Next
        rsCCList.MoveFirst
        If Not rsCCList.EOF Then
            fMapping.TextBox1.Visible = False
            fMapping.ListView2.Visible = True
            gbStatus = False
            'Debug.Print rsCCList.RecordCount 'cEdlQuery.NumRows
            'MsgBox "Zapytanie zwróci³o " & lRowOffset & " wierszy.", vbInformation + vbOKOnly, gsAPP_NAME
        End If
    End If
        Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    On Error Resume Next
    Debug.Print Err.Description
    fMapping.TextBox2.Visible = True
    fMapping.ListView2.Visible = False
    fMapping.TextBox2.MultiLine = True
    fMapping.lblResult.Caption = "Wyst¹pi³ b³¹d"
    fMapping.TextBox2.Value = Err.Description & " w " & Err.Source & " o numerze " & Err.Number & "."
    Resume SubExit
End Sub

Sub CC_download(ByVal gsCAP_REPORT_TITLE As String)
    Dim sQuery                              As String
    Dim iError                              As Integer
    Dim i                                   As Long
    Dim lItem                               As Long
    Dim lRowOffset                          As Long
    Dim rngStartCell                        As Range
    Dim lColumnOffset                       As Long
    Dim sLi                                 As Variant
    Dim sCommand                            As String
    Dim lRow                                As Long
    Dim rsCCList                            As ADODB.Recordset

    Application.ScreenUpdating = False

    On Error GoTo ErrHandler
            
    With fMapping
        .ListView1.Visible = True
        .TextBox1.Visible = False
    End With
    
    With fMapping.ListView1.ListItems
        If Not .Count = 0 Then
            .Clear
        End If
    End With
    '
    If gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_SK Then
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry"
    '
    ElseIf gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_RC Then
    
        Dim sRespCentreCode                       As String

        For lItem = 1 To fMapping.ListView2.ListItems.Count
            If fMapping.ListView2.ListItems(lItem).Checked Then
                gbStatus = True
                sRespCentreCode = fMapping.ListView2.ListItems(lItem).SubItems(1)
                Exit For
            End If
        Next lItem
    
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry"
    '
    ElseIf gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_RU Then
    
        Dim sRespUnitCode                       As String

        For lItem = 1 To fMapping.ListView2.ListItems.Count
            If fMapping.ListView2.ListItems(lItem).Checked Then
                gbStatus = True
                sRespUnitCode = fMapping.ListView2.ListItems(lItem).SubItems(1)
                Exit For
            End If
        Next lItem
        
'        If gbStatus = False Then
'            MsgBox "Wybierz pozycjê, któr¹ chcez dodaæ."
'            Exit Sub
'        End If
    
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry"
    End If

    If sQuery <> "" Then
        Set mcnSQLServer = New ADODB.Connection
        mcnSQLServer.CursorLocation = adUseClient
        mcnSQLServer.ConnectionString = gsCONNECTION_STRING_ODYSSEY
        mcnSQLServer.Open
        Set rsCCList = New ADODB.Recordset
        rsCCList.Open sQuery, mcnSQLServer, adOpenDynamic, adLockOptimistic
    
        With fMapping.ListView1
            .AllowColumnReorder = False
            .CheckBoxes = True
            .FullRowSelect = True
            .MultiSelect = False
            .View = 3
            .Gridlines = True
            With .ColumnHeaders
                .Clear
                .Add Text:="  ", Width:=20
                For i = 1 To rsCCList.Fields.Count
                    .Add , , rsCCList.Fields(i - 1).name, Width:=130
                Next i
            End With
    
            lRowOffset = 0
            lColumnOffset = rsCCList.Fields.Count
            
            With rsCCList
                If .BOF And .EOF Then
                    MsgBox "Zbiór jest pusty."
                    Exit Sub
                End If
            End With
            
            rsCCList.MoveFirst
            With rngStartCell
                While Not rsCCList.EOF
                    ReDim arr(CLng(rsCCList.RecordCount), lColumnOffset)
                    For lItem = 0 To lColumnOffset - 1
                        If lItem = 0 Then
                            Set sLi = fMapping.ListView1.ListItems.Add()
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            'Set sLI = ListView2.ListItems.Add(, , rsCCList.Fields(lItem).Value)
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            GoTo next_iteration
                        End If
                        
                        Select Case rsCCList.Fields(lItem).Type
                        Case 1
                            ' SQL char type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 2
                            ' SQL int type
                            If rsCCList.Fields(lItem).Value = Null Then
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            Else
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            End If
                        Case 3
                            ' SQL float type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 6
                            ' SQL image type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 7
                            ' SQL bit type
                            If rsCCList.Fields(lItem).Value = Null Then
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            Else
                                sLi.ListSubItems.Add , , CInt(rsCCList.Fields(lItem).Value)
                                arr(lRowOffset, lItem) = CInt(rsCCList.Fields(lItem).Value)
                            End If
                        Case 9
                            ' SQL Date/time
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 10
                            ' SQL uniqueidentifier type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case Else
                            ' Other
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        End Select
next_iteration:
                    Next lItem
                    If lRowOffset > 1000 Then GoTo SubExit
                    lRowOffset = lRowOffset + 1
                    rsCCList.MoveNext
                Wend
            End With
        End With
       
SubExit:
        On Error Resume Next
        rsCCList.MoveFirst
        If Not rsCCList.EOF Then
            fMapping.TextBox1.Visible = False
            fMapping.ListView1.Visible = True
            'Debug.Print rsCCList.RecordCount 'cEdlQuery.NumRows
            'MsgBox "Zapytanie zwróci³o " & lRowOffset & " wierszy.", vbInformation + vbOKOnly, gsAPP_NAME
        End If
    End If
        Application.ScreenUpdating = True

    Exit Sub
    
ErrHandler:
    On Error Resume Next
    Debug.Print Err.Description
    fMapping.TextBox2.Visible = True
    fMapping.ListView2.Visible = False
    fMapping.TextBox2.MultiLine = True
    fMapping.lblResult.Caption = "Wyst¹pi³ b³¹d"
    fMapping.TextBox2.Value = Err.Description & " w " & Err.Source & " o numerze " & Err.Number & "."
    Resume SubExit
End Sub

Sub ActiveDBDownload_plus()  '(ByVal gsCAP_REPORT_TITLE As String, ByVal gbStatus As Boolean)
    Dim sQuery                              As String
    Dim iError                              As Integer
    Dim i                                   As Long
    Dim lItem                               As Long
    Dim lRowOffset                          As Long
    Dim rngStartCell                        As Range
    Dim lColumnOffset                       As Long
    Dim sLi                                 As Object
    Dim sCommand                            As String
    Dim lRow                                As Long
    Dim rsCCList                            As ADODB.Recordset

    Application.ScreenUpdating = False

    On Error GoTo ErrHandler
    With fMapping_plus.lvDBNumber.ListItems
        If Not .Count = 0 Then
            .Clear
        End If
    End With

    If gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_SK Then
    Dim sRUCode As String
    sRUCode = Left(fMapping_plus.cbRUItem.Text, InStr(1, fMapping_plus.cbRUItem.Text, "-") - 2)
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry"
    ElseIf gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_RC Then
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry"
    ElseIf gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_RU Then
    Dim sRCCode As String
    sRCCode = Left(fMapping_plus.cbRCItem.Text, InStr(1, fMapping_plus.cbRCItem.Text, "-") - 2)
    
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry"


    End If

    If sQuery <> "" Then
        Set mcnSQLServer = New ADODB.Connection
        mcnSQLServer.CursorLocation = adUseClient
        mcnSQLServer.ConnectionString = gsCONNECTION_STRING_ODYSSEY
        mcnSQLServer.Open
        Set rsCCList = New ADODB.Recordset
        rsCCList.Open sQuery, mcnSQLServer, adOpenDynamic, adLockOptimistic

        With fMapping_plus.lvDBNumber
            .AllowColumnReorder = False
            .CheckBoxes = True
            .FullRowSelect = True
            .MultiSelect = False
            .View = 3
            .Gridlines = True
            With .ListItems

            End With
            With .ColumnHeaders
                .Clear
                .Add Text:="  ", Width:=20
                For lItem = 1 To rsCCList.Fields.Count
                    If lItem < rsCCList.Fields.Count Then

                        .Add , , rsCCList.Fields(lItem - 1).name, Width:=95
                    ElseIf lItem = rsCCList.Fields.Count Then
                        .Add , , rsCCList.Fields(lItem - 1).name, Width:=5
                    End If
                Next lItem
            End With

            With rsCCList
                If .BOF And .EOF Then
                    MsgBox "Zbiór jest pusty."
                    Exit Sub
                End If
            End With

            lRowOffset = 0
            lColumnOffset = rsCCList.Fields.Count

            rsCCList.MoveFirst
            With rngStartCell
                While Not rsCCList.EOF
                    ReDim arr(CLng(rsCCList.RecordCount), lColumnOffset)
                    For lItem = 0 To lColumnOffset - 1
                        If lItem = 0 Then
                            Set sLi = fMapping_plus.lvDBNumber.ListItems.Add()
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            'Set sLI = ListView1.ListItems.Add(, , rsCCList.Fields(lItem).Value)
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            'ActiveSheet.Cells(lRowOffset + 1, lItem + 1) = rsCCList.Fields(lItem).Value
                            GoTo next_iteration
                        End If

                        Select Case rsCCList.Fields(lItem).Type
                        Case 1
                            ' SQL char type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 2
                            ' SQL int type
                            If IsNull(rsCCList.Fields(lItem).Value) Then
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            Else
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            End If
                        Case 3
                            ' SQL float type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 6
                            ' SQL image type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 7
                            ' SQL bit type
                            If IsNull(rsCCList.Fields(lItem).Value) Then
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            Else
                                sLi.ListSubItems.Add , , CInt(rsCCList.Fields(lItem).Value)
                                arr(lRowOffset, lItem) = CInt(rsCCList.Fields(lItem).Value)
                            End If
                        Case 9
                            ' SQL Date/time
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 10
                            ' SQL uniqueidentifier type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case Else
                            ' Other
                            If IsNull(rsCCList.Fields(lItem).Value) Then
                                sLi.ListSubItems.Add , , "-"
                                arr(lRowOffset, lItem) = "-"
                                'ActiveSheet.Cells(lRowOffset + 1, lItem + 1) = "-"
                            Else
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                                'ActiveSheet.Cells(lRowOffset + 1, lItem + 1) = rsCCList.Fields(lItem).Value
                            End If
                        End Select
next_iteration:
                    Next lItem
                    If lRowOffset = 1000 Then GoTo SubExit
                    lRowOffset = lRowOffset + 1
                    rsCCList.MoveNext
                Wend
            End With
        End With

SubExit:
        On Error Resume Next
        rsCCList.MoveFirst
        If Not rsCCList.EOF Then
            fMapping_plus.lvDBNumber.Visible = True
            gbStatus = False

        End If
    End If
    
    For i = 1 To fMapping_plus.lvDBNumber.ListItems.Count
        fMapping_plus.lvDBNumber.ListItems.Item(i).Checked = True
    Next i
    
    Application.ScreenUpdating = True
    
    Exit Sub

ErrHandler:
    On Error Resume Next
    Debug.Print Err.Description
    Resume SubExit
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ActiveDBDownload()  '(ByVal gsCAP_REPORT_TITLE As String, ByVal gbStatus As Boolean)
    Dim sQuery                              As String
    Dim iError                              As Integer
    Dim i                                   As Long
    Dim lItem                               As Long
    Dim lRowOffset                          As Long
    Dim rngStartCell                        As Range
    Dim lColumnOffset                       As Long
    Dim sLi                                 As Object
    Dim sCommand                            As String
    Dim lRow                                As Long
    Dim rsCCList                            As ADODB.Recordset

    Application.ScreenUpdating = False

    On Error GoTo ErrHandler
    With fMapping_ItemAdd.lvDBNumber.ListItems
        If Not .Count = 0 Then
            .Clear
        End If
    End With

    If gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_SK Then
    Dim sRUCode As String
    sRUCode = Left(fMapping_ItemAdd.cbRUItem.Text, InStr(1, fMapping_ItemAdd.cbRUItem.Text, "-") - 2)
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry"
    ElseIf gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_RC Then
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry"
    ElseIf gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_RU Then
    Dim sRCCode As String
    sRCCode = Left(fMapping_ItemAdd.cbRCItem.Text, InStr(1, fMapping_ItemAdd.cbRCItem.Text, "-") - 2)
    
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry"


    End If

    If sQuery <> "" Then
        Set mcnSQLServer = New ADODB.Connection
        mcnSQLServer.CursorLocation = adUseClient
        mcnSQLServer.ConnectionString = gsCONNECTION_STRING_ODYSSEY
        mcnSQLServer.Open
        Set rsCCList = New ADODB.Recordset
        rsCCList.Open sQuery, mcnSQLServer, adOpenDynamic, adLockOptimistic

        With fMapping_ItemAdd.lvDBNumber
            .AllowColumnReorder = False
            .CheckBoxes = True
            .FullRowSelect = True
            .MultiSelect = False
            .View = 3
            .Gridlines = True
            With .ListItems

            End With
            With .ColumnHeaders
                .Clear
                .Add Text:="  ", Width:=20
                For lItem = 1 To rsCCList.Fields.Count
                    If lItem < rsCCList.Fields.Count Then

                        .Add , , rsCCList.Fields(lItem - 1).name, Width:=95
                    ElseIf lItem = rsCCList.Fields.Count Then
                        .Add , , rsCCList.Fields(lItem - 1).name, Width:=5
                    End If
                Next lItem
            End With

            With rsCCList
                If .BOF And .EOF Then
                    MsgBox "Zbiór jest pusty."
                    Exit Sub
                End If
            End With

            lRowOffset = 0
            lColumnOffset = rsCCList.Fields.Count

            rsCCList.MoveFirst
            With rngStartCell
                While Not rsCCList.EOF
                    ReDim arr(CLng(rsCCList.RecordCount), lColumnOffset)
                    For lItem = 0 To lColumnOffset - 1
                        If lItem = 0 Then
                            Set sLi = fMapping_ItemAdd.lvDBNumber.ListItems.Add()
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            'Set sLI = ListView1.ListItems.Add(, , rsCCList.Fields(lItem).Value)
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            'ActiveSheet.Cells(lRowOffset + 1, lItem + 1) = rsCCList.Fields(lItem).Value
                            GoTo next_iteration
                        End If

                        Select Case rsCCList.Fields(lItem).Type
                        Case 1
                            ' SQL char type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 2
                            ' SQL int type
                            If IsNull(rsCCList.Fields(lItem).Value) Then
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            Else
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            End If
                        Case 3
                            ' SQL float type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 6
                            ' SQL image type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 7
                            ' SQL bit type
                            If IsNull(rsCCList.Fields(lItem).Value) Then
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            Else
                                sLi.ListSubItems.Add , , CInt(rsCCList.Fields(lItem).Value)
                                arr(lRowOffset, lItem) = CInt(rsCCList.Fields(lItem).Value)
                            End If
                        Case 9
                            ' SQL Date/time
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 10
                            ' SQL uniqueidentifier type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case Else
                            ' Other
                            If IsNull(rsCCList.Fields(lItem).Value) Then
                                sLi.ListSubItems.Add , , "-"
                                arr(lRowOffset, lItem) = "-"
                                'ActiveSheet.Cells(lRowOffset + 1, lItem + 1) = "-"
                            Else
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                                'ActiveSheet.Cells(lRowOffset + 1, lItem + 1) = rsCCList.Fields(lItem).Value
                            End If
                        End Select
next_iteration:
                    Next lItem
                    If lRowOffset = 1000 Then GoTo SubExit
                    lRowOffset = lRowOffset + 1
                    rsCCList.MoveNext
                Wend
            End With
        End With

SubExit:
        On Error Resume Next
        rsCCList.MoveFirst
        If Not rsCCList.EOF Then
            'fMapping_ItemAdd.TextBox1.Visible = False
            fMapping_ItemAdd.lvDBNumber.Visible = True
            gbStatus = False
            'Debug.Print rsCCList.RecordCount 'cEdlQuery.NumRows
            'MsgBox "Zapytanie zwróci³o " & lRowOffset & " wierszy.", vbInformation + vbOKOnly, gsAPP_NAME
        End If
    End If
    
    For i = 1 To fMapping_ItemAdd.lvDBNumber.ListItems.Count
        fMapping_ItemAdd.lvDBNumber.ListItems.Item(i).Checked = True
    Next i
    
    Application.ScreenUpdating = True
    
    Exit Sub

ErrHandler:
    On Error Resume Next
    Debug.Print Err.Description
'    fMapping_ItemAdd.TextBox2.Visible = True
'    fMapping_ItemAdd.ListView1.Visible = False
'    fMapping_ItemAdd.TextBox2.MultiLine = True
'    fMapping_ItemAdd.lblResult.Caption = "Wyst¹pi³ b³¹d"
'    fMapping_ItemAdd.TextBox2.Value = Err.Description & " w " & Err.Source & " o numerze " & Err.Number & "."
    Resume SubExit
End Sub



Sub LogTab()  '(ByVal gsCAP_REPORT_TITLE As String, ByVal gbStatus As Boolean)
    Dim sQuery                              As String
    Dim iError                              As Integer
    Dim i                                   As Long
    Dim lItem                               As Long
    Dim lRowOffset                          As Long
    Dim rngStartCell                        As Range
    Dim lColumnOffset                       As Long
    Dim sLi                                 As Object
    Dim sCommand                            As String
    Dim lRow                                As Long
    Dim rsCCList                            As ADODB.Recordset

    Application.ScreenUpdating = False

    On Error GoTo ErrHandler
    With fDataProces.ListView1.ListItems
        If Not .Count = 0 Then
            .Clear
        End If
    End With

    sQuery = "Select top 500 [ColumnName], [ColumnName], [ColumnName], [ColumnName] FROM [DB].[dbo].[Table] ORDER BY [ColumnName] DESC"

    'Debug.Print Now()
    'Debug.Print sQuery

    If sQuery <> "" Then
        Set mcnSQLServer = New ADODB.Connection
        mcnSQLServer.CursorLocation = adUseClient
        mcnSQLServer.ConnectionString = gsCONNECTION_STRING_ODYSSEY
        mcnSQLServer.Open
        Set rsCCList = New ADODB.Recordset
        rsCCList.Open sQuery, mcnSQLServer, adOpenDynamic, adLockOptimistic

        With fDataProces.ListView1
            .AllowColumnReorder = False
            .CheckBoxes = False
            .FullRowSelect = True
            .MultiSelect = False
            .View = 3
            .Gridlines = True
            With .ListItems

            End With
            With .ColumnHeaders
                .Clear
                .Add Text:="  ", Width:=0
                For lItem = 1 To rsCCList.Fields.Count
                    Select Case lItem
                        Case 1
                            .Add , , rsCCList.Fields(lItem - 1).name, Width:=33
                        Case 2
                            .Add , , rsCCList.Fields(lItem - 1).name, Width:=95
                        Case 3
                            .Add , , rsCCList.Fields(lItem - 1).name, Width:=95
                        Case 4
                            .Add , , rsCCList.Fields(lItem - 1).name, Width:=65
                    End Select
                Next lItem
            End With

            With rsCCList
                If .BOF And .EOF Then
                    MsgBox "Zbiór jest pusty."
                    Exit Sub
                End If
            End With

            lRowOffset = 0
            lColumnOffset = rsCCList.Fields.Count

            rsCCList.MoveFirst
            With rngStartCell
                While Not rsCCList.EOF
                    ReDim arr(CLng(rsCCList.RecordCount), lColumnOffset)
                    For lItem = 0 To lColumnOffset - 1
                        If lItem = 0 Then
                            Set sLi = fDataProces.ListView1.ListItems.Add()
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            GoTo next_iteration
                        End If

                        Select Case rsCCList.Fields(lItem).Type
                        Case 1
                            ' SQL char type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 2
                            ' SQL int type
                            If IsNull(rsCCList.Fields(lItem).Value) Then
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            Else
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            End If
                        Case 3
                            ' SQL float type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 6
                            ' SQL image type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 7
                            ' SQL bit type
                            If IsNull(rsCCList.Fields(lItem).Value) Then
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            Else
                                sLi.ListSubItems.Add , , CInt(rsCCList.Fields(lItem).Value)
                                arr(lRowOffset, lItem) = CInt(rsCCList.Fields(lItem).Value)
                            End If
                        Case 9
                            ' SQL Date/time
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case 10
                            ' SQL uniqueidentifier type
                            sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                            arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                        Case Else
                            ' Other
                            If IsNull(rsCCList.Fields(lItem).Value) Then
                                sLi.ListSubItems.Add , , "-"
                                arr(lRowOffset, lItem) = "-"
                            Else
                                sLi.ListSubItems.Add , , rsCCList.Fields(lItem).Value
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
                            End If
                        End Select
next_iteration:
                    Next lItem
                    If lRowOffset = 1000 Then GoTo SubExit
                    lRowOffset = lRowOffset + 1
                    rsCCList.MoveNext
                Wend
            End With
        End With

SubExit:
        On Error Resume Next
        rsCCList.MoveFirst
        If Not rsCCList.EOF Then
            fDataProces.ListView1.Visible = True
            gbStatus = False
            'MsgBox "Zapytanie zwróci³o " & lRowOffset & " wierszy.", vbInformation + vbOKOnly, gsAPP_NAME
        End If
    End If
        Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    On Error Resume Next
    Debug.Print Err.Description
    Resume SubExit
End Sub

Function DependenciesCheck(ByVal gsCAP_REPORT_TITLE As String) As Boolean
    Dim sQuery                              As String
    Dim iError                              As Integer
    Dim i                                   As Long
    Dim lItem                               As Long
    Dim lRowOffset                          As Long
    Dim rngStartCell                        As Range
    Dim lColumnOffset                       As Long
    Dim sLi                                 As Variant
    Dim sCommand                            As String
    Dim lRow                                As Long
    Dim rsCCList                            As ADODB.Recordset
    Dim iCompanyID                          As Integer


    Application.ScreenUpdating = False
    
    DependenciesCheck = False

    On Error GoTo ErrHandler
            
    '
    If gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_SK Then
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry"
    '
    ElseIf gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_RC Then
    
        Dim sRespCentreCode                         As String

        For lItem = 1 To fMapping.ListView2.ListItems.Count
            If fMapping.ListView2.ListItems(lItem).Checked Then
                gbStatus = True
                sRespCentreCode = fMapping.ListView2.ListItems(lItem).SubItems(1)
                iCompanyID = GetCompanyID(fMapping.ListView2.ListItems(lItem).SubItems(3))
                Exit For
            End If
        Next lItem
    
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry"
    '
    ElseIf gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_RU Then
    
        Dim sRespUnitCode                           As String
        
        For lItem = 1 To fMapping.ListView2.ListItems.Count
            If fMapping.ListView2.ListItems(lItem).Checked Then
                gbStatus = True
                sRespUnitCode = fMapping.ListView2.ListItems(lItem).SubItems(1)
                iCompanyID = GetCompanyID(fMapping.ListView2.ListItems(lItem).SubItems(5))
                Exit For
            End If
        Next lItem
    
            sQuery = "Querry " & _
                "Querry " & _
                "Querry " & _
                "Querry"
                
    ElseIf gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_SG Then
        
    End If

    If sQuery <> "" Then
        Set mcnSQLServer = New ADODB.Connection
        mcnSQLServer.CursorLocation = adUseClient
        mcnSQLServer.ConnectionString = gsCONNECTION_STRING_ODYSSEY
        mcnSQLServer.Open
        Set rsCCList = New ADODB.Recordset
        rsCCList.Open sQuery, mcnSQLServer, adOpenDynamic, adLockOptimistic
        
        If CLng(rsCCList.RecordCount) > 0 Then
            DependenciesCheck = True
            
            If gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_SK Then

            ElseIf gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_RC Then
                MsgBox "Wskazany obiekt posiada" & CLng(rsCCList.RecordCount) & " przypisanych podobiektów." & Chr(10) & "Nie mo¿e zostaæ usuniêty.", vbInformation + vbOKOnly, gsAPP_NAME
            ElseIf gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_RU Then
                MsgBox "Wskazany obiekt " & CLng(rsCCList.RecordCount) & " przypisanych podobiektów." & Chr(10) & "Nie mo¿e zostaæ usuniêta.", vbInformation + vbOKOnly, gsAPP_NAME
            ElseIf gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_SG Then
                MsgBox "Wskazany obiekt " & CLng(rsCCList.RecordCount) & " przypisanych podobiektów." & Chr(10) & "Nie mo¿e zostaæ usuniêta.", vbInformation + vbOKOnly, gsAPP_NAME
            End If
            
        ElseIf CLng(rsCCList.RecordCount) = 0 Then
            DependenciesCheck = False
        End If
        
SubExit:
        On Error Resume Next
        rsCCList.MoveFirst
        If Not rsCCList.EOF Then
        End If
    End If
       Application.ScreenUpdating = True

    Exit Function
    
ErrHandler:
    On Error Resume Next
    Debug.Print Err.Description
    fMapping.TextBox2.Visible = True
    fMapping.ListView2.Visible = False
    fMapping.TextBox2.MultiLine = True
    fMapping.lblResult.Caption = "Wyst¹pi³ b³¹d"
    fMapping.TextBox2.Value = Err.Description & " w " & Err.Source & " o numerze " & Err.Number & "."
    Resume SubExit
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function fnADOComboboxSetRS(cmb As ComboBox, strSQL1, strSQL2 As String)
    Dim rs As ADODB.Recordset
    Dim errErrors                       As ADODB.Errors
    Dim errErrorItem                    As ADODB.error
    Dim objCommand                      As ADODB.Command

    Const sSource                          As String = "fnADOComboboxSetRS()"
    
    On Error GoTo fnADOComboboxSetRS_Error
    
    If gsCAP_REPORT_TITLE = gsCAP_REPORT_TITLE_SK And InStr(1, strSQL2, "[RespUnitCode]", 1) <> 0 Then
        'cmb.AddItem ""
        cmb.AddItem vbNullString
    End If
    
    If strSQL1 <> "" Then
        Set rs = fnADOSelectCommon(strSQL1, adLockReadOnly, adOpenForwardOnly)
        If rs.Fields.Count > 1 Then
            cmb.Value = rs.Fields(0) & " - " & rs.Fields(1) '& " - " & rs.Fields(2)
        Else
            rs.MoveFirst
            cmb.Value = rs.Fields(0)
        End If
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    
    Set rs = fnADOSelectCommon(strSQL2, adLockReadOnly, adOpenForwardOnly)
       
    If Not rs Is Nothing Then
         If Not (rs.EOF And rs.BOF) Then
             rs.MoveFirst
             Do Until rs.EOF
                If rs.Fields.Count > 1 Then
                    cmb.AddItem rs.Fields(0) & " - " & rs.Fields(1) '& " - " & rs.Fields(2)
                Else
                    cmb.AddItem rs.Fields(0)
                End If
                rs.MoveNext
             Loop
         End If
     End If

fnADOComboboxSetRS_Exit:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    Exit Function

fnADOComboboxSetRS_Error:
    Set objCommand = New ADODB.Command
    If objCommand.ActiveConnection.Errors.Count > 0 Then
        Set errErrors = objCommand.ActiveConnection.Errors
        For Each errErrorItem In errErrors
            If errErrorItem.Number = -2147217873 Then
                MsgBox "Zmiana nie mo¿e zostaæ wprowadzona, poniewa¿ doprowadzi³aby do powstania duplikatów.", vbCritical + vbOKOnly, gsAPP_NAME
                Exit For
            Else
                MsgBox gsERR_LEAD_INFO_GENERAL & errErrorItem.Number & " " & errErrorItem.Description, vbCritical + vbOKOnly, gsAPP_NAME
            End If
        Next errErrorItem
    Else
        'MsgBox gsERR_LEAD_INFO_GENERAL & Err.Source & " (procedure " & sSource & " in module " & msMODULE & ") " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
    End If
End Function

Private Function fnADOSelectCommon(ByVal sQuery As String, adLockReadOnly, adOpenForwardOnly)
    Set mcnSQLServer = New ADODB.Connection
    mcnSQLServer.CursorLocation = adUseClient
    mcnSQLServer.ConnectionString = gsCONNECTION_STRING_ODYSSEY
    mcnSQLServer.Open
    Set fnADOSelectCommon = New ADODB.Recordset
    fnADOSelectCommon.Open sQuery, mcnSQLServer, adOpenDynamic, adLockOptimistic
End Function

 Function GetCompanyID(CompanyNumber As String) As Long
 'GetCompanyID_from_[TFR_Companies]
    Dim sSelectQuery As String
    Dim sConnect As String
    Dim mcnSQLServer As ADODB.Connection
    Dim rsData As ADODB.Recordset
    
    Set mcnSQLServer = New ADODB.Connection
        
    sConnect = gsCONNECTION_STRING_ODYSSEY
    sSelectQuery = "SELECT [ColumnName] FROM [dbo].[Table] WHERE ColumnName = '" & CompanyNumber & "';"
    
    mcnSQLServer.ConnectionString = sConnect
    mcnSQLServer.Open
    Set rsData = mcnSQLServer.Execute(sSelectQuery)
    
    If Not rsData.EOF Then
        GetCompanyID = rsData.Fields(0).Value
    End If
    rsData.Close
    
    Set mcnSQLServer = Nothing
    Set rsData = Nothing
End Function


Function GetCCID_from_TFR_EGCostCentresRespUnitMap(lCCID As Long, sRespUnitID As String) As Long
'GetRUnitID_from_TFR_EGCostCentresRespUnitMap
    Dim sSelectQuery As String
    Dim sConnect As String
    Dim mcnSQLServer As ADODB.Connection
    Dim rsData As ADODB.Recordset
    
    Set mcnSQLServer = New ADODB.Connection
        
    sConnect = gsCONNECTION_STRING_ODYSSEY
    sSelectQuery = "SELECT [CostCentreID] FROM [TFR].[dbo].[TFR_EGCostCentresRespUnitMap] WHERE [CostCentreID] = '" & lCCID & "';"
    'Debug.Print sSelectQuery
    
    mcnSQLServer.ConnectionString = sConnect
    mcnSQLServer.Open
    Set rsData = mcnSQLServer.Execute(sSelectQuery)
    
    If Not rsData.EOF Then
        GetCCID_from_TFR_EGCostCentresRespUnitMap = rsData.Fields(0).Value
    End If
    rsData.Close
    
    Set mcnSQLServer = Nothing
    Set rsData = Nothing
End Function

Function GetRespUnitID(sRespUnitCode, iCompanyID) As Long
'GetRespUnitID_from_[TFR_RespCentreUnits_vw]
Dim sQuery As String
Dim sConnect As String
Dim mcnSQLServer As ADODB.Connection
Dim rsData As ADODB.Recordset
    
    If iCompanyID = 0 Then Exit Function

'    sRespUnitCode = Left(Me.cbRU.Text, InStr(1, Me.cbRU.Text, "-") - 2)
    sQuery = "Select [RespUnitID] FROM [TFR].[dbo].[TFR_RespCentreUnits_vw] where RespUnitCode = '" & sRespUnitCode & "' And CompanyID = '" & iCompanyID & "'"
    'Debug.Print sQuery
    
    
    Set mcnSQLServer = New ADODB.Connection
        
    sConnect = gsCONNECTION_STRING_ODYSSEY
    
    mcnSQLServer.ConnectionString = sConnect
    mcnSQLServer.Open
    Set rsData = mcnSQLServer.Execute(sQuery)
    If Not rsData.EOF Then
        GetRespUnitID = rsData.Fields(0).Value
    End If

    rsData.Close
    
    Set mcnSQLServer = Nothing
    Set rsData = Nothing
    
End Function

Function GetRUnitID(ICCID As Long) As Long
'GetRUnitID_from_TFR_EGCostCentresRespUnitMap
    Dim sSelectQuery As String
    Dim sConnect As String
    Dim mcnSQLServer As ADODB.Connection
    Dim rsData As ADODB.Recordset
    
    Set mcnSQLServer = New ADODB.Connection
        
    sConnect = gsCONNECTION_STRING_ODYSSEY
    sSelectQuery = "SELECT [RespUnitID] FROM [TFR].[dbo].[TFR_EGCostCentresRespUnitMap] WHERE [CostCentreID] = '" & ICCID & "';"
    'Debug.Print sSelectQuery
    
    mcnSQLServer.ConnectionString = sConnect
    mcnSQLServer.Open
    Set rsData = mcnSQLServer.Execute(sSelectQuery)
    
    If Not rsData.EOF Then
        GetRUnitID = rsData.Fields(0).Value
    End If
    rsData.Close
    
    Set mcnSQLServer = Nothing
    Set rsData = Nothing
End Function

Function GetRCentreID(sCentreCode As String, iCompanyID As Integer) As Long
'GetRCentreID_from_[TFR_RespCentres]
    Dim sSelectQuery As String
    Dim sConnect As String
    Dim mcnSQLServer As ADODB.Connection
    Dim rsData As ADODB.Recordset
    
    Set mcnSQLServer = New ADODB.Connection
        
    sConnect = gsCONNECTION_STRING_ODYSSEY
    sSelectQuery = "SELECT [RespCentreID] FROM [TFR].[dbo].[TFR_RespCentres] WHERE [RespCentreCode] = '" & sCentreCode & "' AND[CompanyID] = '" & iCompanyID & "';"
    
    mcnSQLServer.ConnectionString = sConnect
    mcnSQLServer.Open
    Set rsData = mcnSQLServer.Execute(sSelectQuery)
    
    If Not rsData.EOF Then
        GetRCentreID = rsData.Fields(0).Value
    End If
    rsData.Close
    
    Set mcnSQLServer = Nothing
    Set rsData = Nothing
End Function

Function GetRespCentreCode(sRespUnitCode, iCompanyID) As String
'GetRespCentreCode_from_[TFR_RespCentreUnits_vw]
Dim sQuery As String
Dim sConnect As String
Dim mcnSQLServer As ADODB.Connection
Dim rsData As ADODB.Recordset
    
    If iCompanyID = 0 Then Exit Function

    sQuery = "Select [ColumnName], [ColumnName] FROM [DB].[dbo].[DB_vw] where [ColumnName] = '" & sRespUnitCode & "';" ' And [CompanyID] = '" & iCompanyID & "'"
    'Debug.Print sQuery
    
    
    Set mcnSQLServer = New ADODB.Connection
        
    sConnect = gsCONNECTION_STRING_ODYSSEY
    
    mcnSQLServer.ConnectionString = sConnect
    mcnSQLServer.Open
    Set rsData = mcnSQLServer.Execute(sQuery)
    If Not rsData.EOF Then
        GetRespCentreCode = rsData.Fields(0).Value & " - " & rsData.Fields(1).Value
    End If

    rsData.Close
    
    Set mcnSQLServer = Nothing
    Set rsData = Nothing
    
End Function




