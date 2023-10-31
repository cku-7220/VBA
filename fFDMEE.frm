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
    'CSV_Import - zmieniamy podejœcie - tworzymy tabelê przejœciow¹ dla plików csv, kór¹ póŸniej modelujemy, usuwaj¹c wartoœci i dodaj¹c kolumny zale¿nie od nazwy pliku (prod, trad), oraz kolumny z datami.
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

' Tworzymy tabelê tymczasow¹ 'FDM_MAPS_Temp'
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
                
' Uzupe³niamy kolumnê tymczasow¹ 'FDM_MAPS_Temp' danymi z pliku .csv
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
    
' Uzupe³niamy dodane kolumny danymi
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
                
' Zmieniamy rozmiar kolumny [Account] z 75 na 6 - zgodne z FDM_Maps
                sSelectQuery = "ALTER TABLE [DB].[dbo].[Table_" & sTableGUID & "] " & _
                        "ALTER COLUMN [Account] NVARCHAR(6)"
                'Debug.Print sSelectQuery
                mcnSQLServer.Execute (sSelectQuery)
                
' Kopiujemy do tabeli FDM_Maps

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
                
' Usuwamy tabelê tymczasow¹ 'FDM_MAPS_Temp'
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
