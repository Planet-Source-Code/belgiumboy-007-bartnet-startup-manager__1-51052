VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List 
      Height          =   1230
      Left            =   2880
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2280
      Top             =   1080
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      Height          =   375
      Left            =   840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   5
      Left            =   480
      Top             =   480
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   6
      Left            =   480
      Picture         =   "frmMenu.frx":0000
      Top             =   840
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   8
      Left            =   480
      Picture         =   "frmMenu.frx":0342
      Top             =   1560
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   7
      Left            =   480
      Picture         =   "frmMenu.frx":0684
      Top             =   1200
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   3
      Left            =   120
      Picture         =   "frmMenu.frx":09C6
      Top             =   1200
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   4
      Left            =   120
      Picture         =   "frmMenu.frx":0D08
      Top             =   1560
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   2
      Left            =   120
      Picture         =   "frmMenu.frx":104A
      Top             =   840
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   1
      Left            =   120
      Top             =   480
      Width           =   240
   End
   Begin VB.Image img 
      Height          =   240
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCreateKey 
         Caption         =   "Create Key"
      End
      Begin VB.Menu mnuDeleteKey 
         Caption         =   "Delete Key"
      End
      Begin VB.Menu mnuModifyKey 
         Caption         =   "Modify Key"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''#####################################################''
''##                                                 ##''
''##  Created By BelgiumBoy_007                      ##''
''##                                                 ##''
''##  Visit BartNet @ www.bartnet.be for more Codes  ##''
''##                                                 ##''
''##  Copyright 2003 BartNet Corp.                   ##''
''##                                                 ##''
''#####################################################''

Dim pnt As PaintEffects
Dim MyFont As Long
Dim OldFont As Long
Dim wlOldProc As Long
Public LastIndex As Long
Private Caps(2 To 8) As String

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long

Dim RegKey() As KeyInfo
Dim isKeyEmpty As Boolean

Private Type KeyInfo
    Name As String
    Value As String
    PLocation As String
End Type

Private IniFile As New clsIniFile

Public Function MsgProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim MeasureInfo As MEASUREITEMSTRUCT
    Dim DrawInfo As DRAWITEMSTRUCT
    Dim IsSep As Boolean
    Dim hBr As Long, hOldBr As Long
    Dim hPEN As Long, hOldPen As Long
    Dim lTextColor As Long
    Dim iRectOffset As Integer
    
    If wMsg = WM_DRAWITEM Then
        If wParam = 0 Then
            Call CopyMem(DrawInfo, ByVal lParam, LenB(DrawInfo))
            IsSep = IsSeparator(DrawInfo.itemID)
            
            MyFont = SendMessage(Me.hWnd, WM_GETFONT, 0&, 0&)
            OldFont = SelectObject(DrawInfo.hdc, MyFont)
            If (DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED Then
                hBr = CreateSolidBrush( _
                GetSysColor(COLOR_HIGHLIGHT))
                lTextColor = GetSysColor(COLOR_MENUTEXT)
            Else
                hBr = CreateSolidBrush(GetSysColor(COLOR_MENU))
                hPEN = GetPen(1, GetSysColor(COLOR_MENU))
                lTextColor = GetSysColor(COLOR_MENUTEXT)
            End If
            QuickGDI.TargethDC = DrawInfo.hdc
            

            hOldBr = SelectObject(DrawInfo.hdc, hBr)
            hOldPen = SelectObject(DrawInfo.hdc, hPEN)
            With DrawInfo.rcItem
                If (DrawInfo.itemState And ODS_SELECTED) <> ODS_SELECTED Then
                    QuickGDI.DrawRect .Left, .Top, 22, .Bottom
                End If
                
                iRectOffset = IIf(img(DrawInfo.itemID).Picture.Handle <> 0 _
                    , 23, 0)
                If Not IsSep Then
                    
                    QuickGDI.DrawRect .Left + iRectOffset, .Top, .Right, .Bottom
                    
                    DrawFilledRect DrawInfo.hdc, .Left, .Top, .Right, .Bottom, vbWhite
                    
                    If (DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED Then
                        DrawFilledRect1 .Left, .Top, .Right, .Bottom
                    End If
                    
                    'Print the item's text
                    '(held in the Caps() array)
                    hPrint .Left + 30, .Top + 3, Caps(DrawInfo.itemID), lTextColor
                End If
            End With
            Call SelectObject(DrawInfo.hdc, hOldBr)
            Call SelectObject(DrawInfo.hdc, hOldPen)
            Call DeleteObject(hBr)
            Call DeleteObject(hPEN)
            With DrawInfo
                If img(DrawInfo.itemID).Picture.Handle <> 0 Then
                    
                    If (DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED Then
                        Dim i As Long
                        Dim e As Long
                        Dim a As Long
                        
                        Picture2.Cls
                        
                        pnt.PaintTransparentStdPic Picture2.hdc, 0, 0, _
                            16, 16, img(DrawInfo.itemID).Picture, 0, 0, vbMagenta
                        
                        For i = 0 To 16
                            For e = 0 To 16
                                a = GetPixel(Picture2.hdc, i, e)
                                If a <> vbMagenta Then SetPixel Picture2.hdc, i, e, RGB(158, 158, 165)
                            Next
                        Next
                    
                        Picture2.Refresh
                        
                                pnt.PaintTransparentDC .hdc, _
                            5, .rcItem.Top + 6, _
                            16, 16, Picture2.hdc, _
                            0, 0, vbMagenta
                                                    
                        pnt.PaintTransparentStdPic .hdc, _
                            3, .rcItem.Top + 4, _
                            16, 16, img(DrawInfo.itemID).Picture, _
                            0, 0, vbMagenta
                    Else
                        DrawFilledRect .hdc, 0, .rcItem.Top, 23, .rcItem.Bottom, RGB(241, 240, 242)
                        
                        pnt.PaintTransparentStdPic .hdc, _
                            4, .rcItem.Top + 5, _
                            16, 16, img(DrawInfo.itemID).Picture, _
                            0, 0, vbMagenta
                    End If
                    
                End If
                If IsSep Then
                    Dim pt As POINTAPI
                    DrawFilledRect .hdc, .rcItem.Left, .rcItem.Top, .rcItem.Right, .rcItem.Bottom, vbWhite
                    DrawFilledRect .hdc, 0, .rcItem.Top, 23, .rcItem.Bottom, RGB(241, 240, 242)
                    MoveToEx .hdc, .rcItem.Left + 25, .rcItem.Top + 2, pt
                    LineTo .hdc, .rcItem.Right, .rcItem.Top + 2
                End If
            End With
        End If
        MsgProc = False
        Exit Function
        
    ElseIf wMsg = WM_MEASUREITEM Then
        Call CopyMem(MeasureInfo, ByVal lParam, Len(MeasureInfo))
        IsSep = IsSeparator(MeasureInfo.itemID)
        MeasureInfo.itemWidth = 150
        MeasureInfo.itemHeight = IIf(IsSep, 5, 22)
        Call CopyMem(ByVal lParam, MeasureInfo, Len(MeasureInfo))
        MsgProc = False
        Exit Function
    ElseIf wMsg = WM_MENUSELECT Then
        
    End If
    
    MsgProc = CallWindowProc(wlOldProc, hWnd, wMsg, wParam, lParam)
End Function

Public Function IsSeparator(ByVal IID As Integer) As Boolean
    Dim mii As MENUITEMINFO
    mii.cbSize = Len(mii)
    mii.fMask = MIIM_TYPE
    mii.wID = IID
    GetMenuItemInfo GetMenu(hWnd), IID, False, mii
    IsSeparator = ((mii.fType And MFT_SEPARATOR) = MFT_SEPARATOR)
End Function

Private Sub Form_Load()
    Set pnt = New PaintEffects
    
    Caps(2) = "Open"
    Caps(3) = "Options"
    Caps(4) = "Exit"
    
    Caps(6) = "Create Key"
    Caps(7) = "Delete Key"
    Caps(8) = "Modify Key"
    
    If wlOldProc <> 0 Then Exit Sub
    
    Dim i As Integer
    
    MenuItems.MenuForm = Me

    MenuItems.SubMenu = 0
    For i = 0 To 5
        MenuItems.MenuID = i
        OwnerDrawMenu (i + 2)
    Next
    
    MenuItems.SubMenu = 1
    For i = 0 To 5
        MenuItems.MenuID = i
        OwnerDrawMenu (i + 2)
    Next
        
    wlOldProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf OwnMenuProc)
End Sub

Private Sub mnuCreateKey_Click()
    frmMain.cmdCreate_Click
End Sub

Private Sub mnuDeleteKey_Click()
    frmMain.cmdDelete_Click
End Sub

Private Sub mnuExit_Click()
    Shell_NotifyIcon NIM_DELETE, nid
    
    Unload frmMain
    Unload frmOptions
    Unload Me
End Sub

Private Sub mnuModifyKey_Click()
    frmMain.cmdModify_Click
End Sub

Private Sub mnuOpen_Click()
    If IniFile.ReadFrom("Other", "Speed", "5000") <> "NONE" Then Timer.Enabled = True
    frmMain.Show
    frmMain.TabStrip_Click
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show
End Sub

Private Sub Timer_Timer()
Debug.Print "TIMER REACHED" & vbCrLf
    With frmMain
        Dim i As Integer
        Dim Count As Integer
        Dim tmp As String
        Dim clsReg As clsRegistry
        Dim vKeyArray() As Variant
        Dim vValueArray() As Variant
        Dim errOccur As Boolean
        Dim isKeyOK As Boolean
        Dim DoneSomething As Boolean
        
        DoneSomething = False
        
        Set clsReg = New clsRegistry
        
        If .TabStrip.SelectedItem.Index = 1 Then
            tmp = "HKEY_CURRENT_USER"
            clsReg.hKey = HKEY_CURRENT_USER
        Else
            tmp = "HKEY_LOCAL_MACHINE"
            clsReg.hKey = HKEY_LOCAL_MACHINE
        End If
Redo:
        Count = .ListView.ListItems.Count + 1
        i = 1
        
        Do Until i = Count
            If .KeyExists(.ListView.ListItems.Item(i).Text, tmp) = False Then
                .ListView.ListItems.Remove (i)
                DoneSomething = True
                GoTo Redo
            End If
            
            i = i + 1
            
            DoEvents
        Loop
        
On Error GoTo errReport
        errOccur = False
       
        clsReg.SubKey = "Software\Microsoft\Windows\CurrentVersion\Run"
    
        If Not clsReg.Reg_OpenKey Then GoTo errHandler

        isKeyOK = clsReg.GetAllValues(vKeyArray, vValueArray)
        
        If isKeyOK Then
            isKeyEmpty = False
           
            Dim a As Integer
            
            If Not isKeyEmpty Then
                ReDim RegKey(UBound(vKeyArray))
                For a = LBound(RegKey) To UBound(RegKey)
                    RegKey(a).Name = vKeyArray(a)
                    RegKey(a).Value = vValueArray(a)
                    RegKey(a).PLocation = ""
                Next a
            End If
        Else
            isKeyEmpty = True
        End If
        
        Dim b As Integer
        Dim Item As ListItem
        Dim tmpPath As String
        Dim Found As Boolean
                
        If isKeyEmpty Then
            ListView.ListItems.Add , , KEY_EMPTY
            DoneSomething = True
        Else
            For b = LBound(RegKey) To UBound(RegKey)
                i = 1
                Found = False
                Count = .ListView.ListItems.Count + 1
                
                Do Until i = Count
                    If UCase(RegKey(b).Name) = UCase(.ListView.ListItems.Item(i).Text) Then
                        Found = True
                    End If
                    
                    i = i + 1
                    
                    DoEvents
                Loop
                
                If Found = False Then
                    Set Item = .ListView.ListItems.Add(, , RegKey(b).Name, "Default", "Default")
                    DoneSomething = True
                            
                    Item.SubItems(1) = RegKey(Item.Index - 1).Value
                
                    tmpPath = RegKey(Item.Index - 1).Value
                    
                    If Mid(tmpPath, 1, 1) = """" Then
                        Dim c As Integer
                        
                        c = 2
                        
                        Do Until c = Len(tmpPath)
                            If Mid(tmpPath, c, 1) = """" Then
                                Exit Do
                            Else
                                c = c + 1
                            End If
                        Loop
                        
                        tmpPath = Mid(tmpPath, 2, c - 2)
                    End If
                
                    If UCase(Mid(tmpPath, 1, 8)) = "RUNDLL32" Then
                        tmpPath = "c:\windows\system32\rundll32.exe"
                    End If
                    
                    tmpPath = Replace(tmpPath, "\\", "\")
    
                    If FileExists(tmpPath) = True Then
                        Item.Icon = ExtractIcon(tmpPath, .ImageList, .picTemp, 16)
                        Item.SmallIcon = ExtractIcon(tmpPath, .ImageList, .picTemp, 16)
                    End If
                End If
            Next b
            
            If DoneSomething = True Then .ListView.ListItems.Item(1).Selected = True
        End If
    
CleanUp:
        On Error Resume Next
        clsReg.Reg_CloseKey
        Set clsReg = Nothing
'        If errOccur Then End'
        Exit Sub

errHandler:
        On Error GoTo errReport
        Err.Raise vbObjectError, , "Error working/opening registry key" & vbCr & vbCr & "Ending program."
errReport:
        MsgBox Err.Number & " - " & Err.Description
        errOccur = True
        GoTo CleanUp
    End With
End Sub
