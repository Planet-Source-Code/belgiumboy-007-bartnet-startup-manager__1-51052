Attribute VB_Name = "StartUp"
''#####################################################''
''##                                                 ##''
''##  Created By BelgiumBoy_007                      ##''
''##                                                 ##''
''##  Visit BartNet @ www.bartnet.be for more Codes  ##''
''##                                                 ##''
''##  Copyright 2003 BartNet Corp.                   ##''
''##                                                 ##''
''#####################################################''

Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public nid As NOTIFYICONDATA

Public Enum StartWindowState
    START_HIDDEN = 0
    START_NORMAL = 4
    START_MINIMIZED = 2
    START_MAXIMIZED = 3
End Enum

Private IniFile As New clsIniFile

Private Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hdcDest&, ByVal x&, ByVal Y&, ByVal Flags&) As Long

Private FileInfo As typSHFILEINFO

Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400

Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

#If Win32 Then

Public Declare Function ShellExecute2 Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else

Public Declare Function ShellExecute2 Lib "shell.dll" (ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
#End If
Public Const SW_SHOWNORMAL = 1

Sub Main()
'    App.Title = "BartNet Startup Manager"
    App.Title = ""
    
    With nid
        .cbSize = Len(nid)
        .hWnd = frmMain.hWnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = frmMain.Icon
        .szTip = "BartNet Startup Manager" & vbNullChar
    End With
        
    Shell_NotifyIcon NIM_ADD, nid
    
'    Load frmMenu
    frmMenu.Left = Screen.Width * 2
    frmMenu.Show
    frmMenu.Visible = False
    
    If IniFile.ReadFrom("Other", "RunBefore", "NO") = "YES" Then
        If IniFile.ReadFrom("Other", "Alert", "1") = 1 Then
            CheckProgList
        End If
    End If
    
    ApplySettings
    SaveProgList
End Sub

Public Sub LoadSettings()
    With frmOptions
        .ckSecurityCreation.Value = IniFile.ReadFrom("Security", "Creation", "0")
        .ckSecurityDeletion.Value = IniFile.ReadFrom("Security", "Deletion", "0")
        .ckSecurityModification.Value = IniFile.ReadFrom("Security", "Modification", "0")
        
        .ckListColumnHeaders.Value = IniFile.ReadFrom("ListView", "ColumnHeaders", "1")
        .cboListBorder.Text = IniFile.ReadFrom("ListView", "BorderStyle", "Fixed Single")
        .cboListAppearance.Text = IniFile.ReadFrom("ListView", "Appearance", "3D")
        
        .ckOtherRun.Value = IniFile.ReadFrom("Other", "Run", "1")
        .ckOtherAlert.Value = IniFile.ReadFrom("Other", "Alert", "1")
        .txtOtherSpeed.Text = IniFile.ReadFrom("Other", "Speed", "5000")
        
        If .txtOtherSpeed.Text = "NONE" Then .txtOtherSpeed.Text = ""
        
        .cboTabStyle.Text = IniFile.ReadFrom("TabStrip", "Style", "Tabs")
    End With
End Sub

Public Sub ApplySettings()
    If IniFile.ReadFrom("ListView", "ColumnHeaders", "1") = 1 Then frmMain.ListView.HideColumnHeaders = False Else frmMain.ListView.HideColumnHeaders = True

    If IniFile.ReadFrom("ListView", "BorderStyle", "Fixed Single") = "Fixed Single" Then frmMain.ListView.BorderStyle = ccFixedSingle Else frmMain.ListView.BorderStyle = ccNone
    If IniFile.ReadFrom("ListView", "Appearance", "3D") = "3D" Then frmMain.ListView.Appearance = cc3D Else frmMain.ListView.Appearance = ccFlat
    
    If IniFile.ReadFrom("Other", "Run", "1") = 1 Then SetKeyValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "BartNet Startup Manager", App.Path & "\" & App.EXEName & ".exe", 1 Else DeleteValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "BartNet Startup Manager"
    
    Dim tmp As Integer
    
    If IniFile.ReadFrom("Other", "Speed", "5000") = "NONE" Then
        frmMenu.Timer.Enabled = False
    Else
        tmp = IniFile.ReadFrom("Other", "Speed", "5000")
        frmMenu.Timer.Interval = tmp
        
        If frmMenu.Visible = True Then frmMenu.Timer.Enabled = True
    End If
    
    If IniFile.ReadFrom("TabStrip", "Style", "Tabs") = "Tabs" Then frmMain.TabStrip.Style = tabTabs Else frmMain.TabStrip.Style = tabButtons
    
    IniFile.WriteTo "Other", "RunBefore", "YES"
End Sub

Public Function ExtractIcon(FileName As String, AddtoImageList As ImageList, PictureBox As PictureBox, PixelsXY As Integer) As Long
    Dim SmallIcon As Long
    Dim NewImage As ListImage
    Dim IconIndex As Integer
    
    If PixelsXY = 16 Then
        SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_SMALLICON)
    Else
        SmallIcon = SHGetFileInfo(FileName, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
    End If
    
    If SmallIcon <> 0 Then
      With PictureBox
        .Height = 15 * PixelsXY
        .Width = 15 * PixelsXY
        .ScaleHeight = 15 * PixelsXY
        .ScaleWidth = 15 * PixelsXY
        .Picture = LoadPicture("")
        .AutoRedraw = True
        
        SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, PictureBox.hdc, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      
      IconIndex = AddtoImageList.ListImages.Count + 1
      Set NewImage = AddtoImageList.ListImages.Add(IconIndex, , PictureBox.Image)
      ExtractIcon = IconIndex
    End If
End Function

Public Function FileExists(ByVal sFileName As String) As Boolean
    Dim i As Integer
    
    On Error GoTo NotFound
    
    i = GetAttr(sFileName)
    FileExists = True
    
    Exit Function
    
NotFound:
    FileExists = False
End Function

Public Sub SaveProgList()
    frmMain.Visible = False
    frmMain.TabStrip.Tabs.Item(1).Selected = True
    frmMain.TabStrip_Click
    
    Dim i As Integer
    
    i = 1
    
    Do Until i = frmMain.ListView.ListItems.Count + 1
        IniFile.WriteTo "CurrentUser", frmMain.ListView.ListItems.Item(i).Text, frmMain.ListView.ListItems.Item(i).SubItems(1)
        
        i = i + 1
    Loop
    
    frmMain.TabStrip.Tabs.Item(2).Selected = True
    frmMain.TabStrip_Click
    
    i = 1
    
    Do Until i = frmMain.ListView.ListItems.Count + 1
        IniFile.WriteTo "AllUsers", frmMain.ListView.ListItems.Item(i).Text, frmMain.ListView.ListItems.Item(i).SubItems(1)
        
        i = i + 1
    Loop
End Sub

Public Sub CheckProgList()
    frmMain.Visible = False
    frmMain.TabStrip.Tabs.Item(1).Selected = True
    frmMain.TabStrip_Click
    
    Dim i As Integer
    Dim blNew As Boolean
    Dim intNew As Integer
    
    blNew = False
    
    i = 1
    
    Do Until i = frmMain.ListView.ListItems.Count + 1
        If IniFile.ReadFrom("CurrentUser", frmMain.ListView.ListItems.Item(i).Text, "DEFAULT") = "DEFAULT" Then
            frmMenu.List.AddItem frmMain.ListView.ListItems.Item(i).Text
            blNew = True
            intNew = intNew + 1
        End If
        
        i = i + 1
    Loop
    
    frmMain.TabStrip.Tabs.Item(2).Selected = True
    frmMain.TabStrip_Click
    
    i = 1
    
    Do Until i = frmMain.ListView.ListItems.Count + 1
        If IniFile.ReadFrom("AllUsers", frmMain.ListView.ListItems.Item(i).Text, "DEFAULT") = "DEFAULT" Then
            frmMenu.List.AddItem frmMain.ListView.ListItems.Item(i).Text
            blNew = True
            intNew = intNew + 1
        End If
        
        i = i + 1
    Loop
    
    If blNew = True Then
        If intNew = 1 Then
            MsgBox "BartNet Startup Manager has found a new registry key.", vbOKOnly + vbInformation, "Warning"
        Else
            MsgBox "BartNet Startup Manager has found " & intNew & " new registry keys.", vbOKOnly + vbInformation, "Warning"
        End If
    End If
End Sub
