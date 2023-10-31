VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fFDMEE 
   Caption         =   "UserForm1"
   ClientHeight    =   4560
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7272
   OleObjectBlob   =   "fFDMEE.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fFDMEE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef GUID As Byte) As Long

Private Sub cmdClose_Click()
    Unload Me
    fMetaHandler.Show
End Sub

Private Sub cmdImport_Click()
    
    Dim sSelectQuery As String
    Dim sConnect As String
    Dim mcnSQLServer As ADODB.Connection
    Dim bItem As Byte
    Dim sSourceFile As String
    Dim iFirstChar As Integer
    Dim iSecondChar As Integer
    Dim iLength As Integer
    Dim sPartName As String
    Dim sTableGUID As String
    
    On Error GoTo error
    
    Set mcnSQLServer = New ADODB.Connection
        
    sConnect = gsCONNECTION_STRING_ODYSSEY
    mcnSQLServer.ConnectionString = sConnect
    'mcnSQLServer.ConnectionTimeout = 0
    mcnSQLServer.CommandTimeout = 0
    mcnSQLServer.Open
            
    With Me
        For bItem = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(bItem).Checked = True Then
                sSourceFile = Me.txtPath.Text & "\" & ListView1.ListItems(bItem).SubItems(1)
                
                If InStr(1, ListView1.ListItems(bItem).SubItems(1), "PolandPROD", 1) <> 0 Then
                    sPartName = "PolandPROD"
                ElseIf InStr(1, ListView1.ListItems(bItem).SubItems(1), "PolandTRAD", 1) <> 0 Then
                    sPartName = "PolandTRAD"
                Else
                    MsgBox "Podany plik '" & ListView1.ListItems(bItem).SubItems(1) & "' nie zosta³ rozpoznany jako plik Ÿród³owy."
                    GoTo nextFile
                End If
                
                sTableGUID = GetGUID


                sSelectQuery = "CREATE TABLE FDM_MAPS_Temp_" & sTableGUID & " ([Edytuj pozycje noty] VARCHAR(20), " & _
                        "[ColumnName] NVARCHAR(75), " & _
                        "[ColumnName] NVARCHAR(75), " & _
                        "[ColumnName] NVARCHAR(75), " & _
                        "[ColumnName] NVARCHAR(75), " & _
                        "[ColumnName] NVARCHAR(75), " & _
                        "[ColumnName] NVARCHAR(75), " & _
                        "[ColumnName] NVARCHAR(75), " & _
                        "[ColumnName] NVARCHAR(75), " & _
                        "[ColumnName] NVARCHAR(75), " & _
                        "[ColumnName] NVARCHAR(75), " & _
                        "[ColumnName] NVARCHAR(75), " & _
                        "[ColumnName] NVARCHAR(75), " & _
                        "[ColumnName] NVARCHAR(75), " & _
                        "[ColumnName] NVARCHAR(75), " & _
                        "[ColumnName] NVARCHAR(50), " & _
                        "[ColumnName] NVARCHAR(50)) "
                'Debug.Print sSelectQuery
                mcnSQLServer.Execute (sSelectQuery)
                
' Uzupelniamy kolumny tymczasowe 'Table' danymi z pliku .csv
                sSelectQuery = "BULK INSERT [DB].[dbo].[Table_" & sTableGUID & "] " & _
                        "FROM '" & sSourceFile & "' " & _
                        "WITH (FORMAT = 'CSV', FIRSTROW = 2, FIELDTERMINATOR = ';', " & _
                        "ROWTERMINATOR = '\n', CHECK_CONSTRAINTS, KEEPIDENTITY)"
                'Debug.Print sSelectQuery
                mcnSQLServer.Execute (sSelectQuery)
                
' Usuwamy niepotrzebne kolumny
                sSelectQuery = "ALTER TABLE [DB].[dbo].[Table_" & sTableGUID & "] " & _
                        "DROP COLUMN [Kwota], [Kwota Ÿród³owa], [Edytuj pozycje noty];"
                'Debug.Print sSelectQuery
                mcnSQLServer.Execute (sSelectQuery)
                
' Dodajemy potrzebne kolumny
                sSelectQuery = "ALTER TABLE [DB].[dbo].[Table_" & sTableGUID & "] " & _
                        "ADD [PartName] [nvarchar](20), " & _
                        "[PeriodKey] [date], " & _
                        "[PeriodKeyYear] [int];"
                'Debug.Print sSelectQuery
                mcnSQLServer.Execute (sSelectQuery)
    
' Uzupelniamy dodane kolumny danymi
                sSelectQuery = "UPDATE [DB].[dbo].[Table_" & sTableGUID & "] " & _
                        "Set [PartName]  = '" & sPartName & "', " & _
                        "[PeriodKey] = '" & Format(Me.dtpDateFrom.Text, "YYYY-MM-DD") & "', " & _
                        "[PeriodKeyYear] = '" & Year(Me.dtpDateFrom.Text) & "';"
                'Debug.Print sSelectQuery
                mcnSQLServer.Execute (sSelectQuery)
                
' Obcinamy [Account] do 6 pierwszych znaków
                sSelectQuery = "Update [DB].[dbo].[Table_" & sTableGUID & "] " & _
                        "set [Account] = Left([Account], 6)"
                'Debug.Print sSelectQuery
                mcnSQLServer.Execute (sSelectQuery)
                
' Zmieniamy rozmiar kolumny [Account] z 75 na 6 - zgodne z 'Table'
                sSelectQuery = "ALTER TABLE [DB].[dbo].[Table_" & sTableGUID & "] " & _
                        "ALTER COLUMN [Account] NVARCHAR(6)"
                'Debug.Print sSelectQuery
                mcnSQLServer.Execute (sSelectQuery)
                
' Kopiujemy do tabeli 'Table'

            'Remove mappings in [dbo].[FDM_Maps] for newly imported PartNames and PeriodKeys
                sSelectQuery = "DELETE [DB].[dbo].[Table] FROM [DB].[dbo].[Table] AS CT INNER JOIN [DB].[dbo].[Table_Temp_" & sTableGUID & "] AS UP ON [CT].[ColumnName] = [UP].[ColumnName] AND [CT].[PeriodKey] = [UP].[PeriodKey] AND [UP].[UD1] NOT LIKE '%QTY';"
                'Debug.Print sSelectQuery
                mcnSQLServer.Execute (sSelectQuery)
                
            'Copy new mappings from [dbo].[FDM_Maps_Upload] to [dbo].[FDM_Maps]
                sSelectQuery = "INSERT INTO [DB].[dbo].[Table] SELECT " & _
                        " [ColumnName],  " & _
                        " [ColumnName], " & _
                        " [ColumnName], " & _
                        " [ColumnName] , " & _
                        " [ColumnName] , " & _
                        " [ColumnName] , " & _
                        " [ColumnName] , " & _
                        " [ColumnName] , " & _
                        " [ColumnName] , " & _
                        " [ColumnName] , " & _
                        " [ColumnName] , " & _
                        " [ColumnName] , " & _
                        " [ColumnName] , " & _
                        " [ColumnName] , " & _
                        " [ColumnName] , " & _
                        " [ColumnName] , " & _
                        " [ColumnName] " & _
                        "FROM [DB].[dbo].[Table_" & sTableGUID & "] WHERE [DB].[dbo].[Table_" & sTableGUID & "].[ColumnName] NOT LIKE '%QTY';"
                'Debug.Print sSelectQuery
                mcnSQLServer.Execute (sSelectQuery)
                
' Usuwamy tabele tymczasowa 'Table_Temp'
                sSelectQuery = "DROP TABLE [DB].[dbo].[Table_" & sTableGUID & "];"
                'Debug.Print sSelectQuery
                mcnSQLServer.Execute (sSelectQuery)
            End If
nextFile:
        Next bItem
    End With
    
    MsgBox "Operacja zakoñczona powodzeniem."

    End

error:
    
    MsgBox "Operacja zakoñczona niepowodzeniem."
    sSelectQuery = "DROP TABLE [DB].[dbo].[Table_" & sTableGUID & "]; "
    'Debug.Print sSelectQuery
    mcnSQLServer.Execute (sSelectQuery)
    
    Dim ADOError As ADODB.error
    For Each ADOError In mcnSQLServer.Errors
      MsgBox "ADO Error: " & ADOError.Description & " Native Error: " & _
        ADOError.NativeError & " SQL State: " & _
        ADOError.SqlState & "Source: " & _
        ADOError.Source, vbCritical, "Error Number: " & ADOError.Number
    Next ADOError
    
    Set mcnSQLServer = Nothing
    End
End Sub

Private Sub dtpDateFrom_DropButtonClick()
    Set Calendar1 = New cCalendar
    fFDMEE.dtpDateFrom.SetFocus
    fFDMEE.Tag = fFDMEE.dtpDateFrom.Value
    fCalendar.Show
    fFDMEE.dtpDateFrom.Value = fFDMEE.Tag
End Sub

Private Sub txtPath_DropButtonClick()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Wybierz folder z plikami Ÿród³owymi"
        .Filters.Clear
        '.Filters.Add ".CSV", "*.CSV", 1
        .AllowMultiSelect = False
        .InitialView = msoFileDialogViewDetails
        If Me.txtPath.Text <> "" Then
            .InitialFileName = Me.txtPath.Text
        End If
        If .Show Then
            txtPath.Text = .SelectedItems(1)
            LoopAllFilesInAFolder (Me.txtPath.Text)
        End If
    End With
    
    FilePathString_Set Me.txtPath.Text, gsREG_FOLDER_FULL_PATH_SOURCE
End Sub

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler
    
    With Me
        .Height = 258
        .Width = 375.75
        .Caption = "FDMEE"
        .Frame1.Caption = "Opcje importu plików FDMEE"
        .cmdClose.Caption = "Zamknij"
        .cmdClose.Visible = True
        .cmdImport.Caption = "Importuj"
        .cmdImport.Visible = True
        .cmdImport.SetFocus
        .lblPath.Caption = "Wybierz folder plików Ÿród³owych"
        .lblLV.Caption = "Wybierz plik"
        With .txtPath
            .Text = FilePathString_Get(gsREG_FOLDER_FULL_PATH_SOURCE)
            .Enabled = True
            .DropButtonStyle = fmDropButtonStyleEllipsis
            .ShowDropButtonWhen = fmShowDropButtonWhenAlways
        End With
        With .dtpDateFrom
            .TextAlign = fmTextAlignLeft
            .Enabled = True
            .DropButtonStyle = fmDropButtonStyleEllipsis
            .ShowDropButtonWhen = fmShowDropButtonWhenAlways
        End With
        .lblDateFrom.Caption = "Data raportowania"
        .dtpDateFrom = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "yyyy-mm-dd")
        '.dtpDateFrom.Value = Format(DateSerial(IIf(Month(Now()) > 1, Year(Now()), Year(Now()) - 1), 1, 1), "yyyy-mm-dd")
    End With
    
    LoopAllFilesInAFolder (Me.txtPath.Text)
ErrHandler:
End Sub

Sub LoopAllFilesInAFolder(folderselected)
    Dim fileName As String
    Dim sLi As ListItem
    Dim i As Integer
    fileName = Dir(folderselected & "\")
    
    ListView1.ListItems.Clear
    
    While fileName <> ""
        
        With fFDMEE.ListView1
            .AllowColumnReorder = False
            .CheckBoxes = True
            .FullRowSelect = True
            .MultiSelect = False
            .View = 3
            .Gridlines = True
            
            With .ColumnHeaders
                .Clear
                .Add Text:="  ", Width:=20
                .Add , , "Nazwa pliku", Width:=200
            End With
        End With
                        
        Set sLi = fFDMEE.ListView1.ListItems.Add()
        sLi.ListSubItems.Add , , fileName
        
        fileName = Dir
    Wend
    
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(i).Checked = True
    Next i
End Sub

Private Function FilePathString_Set(Optional ByVal sFilePathStringValue As String, Optional ByVal sRegKey As String) As Boolean
    SaveSetting gsREG_APP, gsREG_SECTION, sRegKey, sFilePathStringValue
End Function

Private Function FilePathString_Exists(Optional sFilePathStringName As String) As Boolean
    Dim FilePathString As String
    
    FilePathString = GetSetting(gsREG_APP, gsREG_SECTION, sFilePathStringName, "")
    If FilePathString <> "" Then
        FilePathString_Exists = True
    Else
        FilePathString_Exists = False
    End If
End Function

Private Function FilePathString_Get(Optional ByVal sRegKey As String) As String

    If Not FilePathString_Exists(sRegKey) Then
        FilePathString_Set 0, sRegKey
    End If

    FilePathString_Get = GetSetting(gsREG_APP, gsREG_SECTION, sRegKey, "")
End Function

Private Function LastAppear(rCell As String, rChar As String)
Dim i As Integer
Dim rLen As Integer
    
    rLen = Len(rCell)
    For i = rLen To 1 Step -1
        If Mid(rCell, i, 1) = rChar Then
            LastAppear = i
            Exit Function
        End If
    Next i
End Function

Private Function GetGUID() As String
    Dim ID(0 To 15) As Byte
    Dim N As Long
    Dim GUID As String
    Dim Res As Long
    
    Res = CoCreateGuid(ID(0))
    For N = 0 To 15
    GUID = GUID & IIf(ID(N) < 16, "0", "") & Hex$(ID(N))
    Next N
    GetGUID = GUID
End Function

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

    End If


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
                                arr(lRowOffset, lItem) = rsCCList.Fields(lItem).Value
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
            fMapping.TextBox1.Visible = False
            fMapping.ListView2.Visible = True
            gbStatus = False

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
    fMapping.lblResult.Caption = "Wystąpił błąd"
    fMapping.TextBox2.Value = Err.Description & " w " & Err.Source & " o numerze " & Err.Number & "."
    Resume SubExit
End Sub
