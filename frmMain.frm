VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "BartNet Startup Manager"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picTemp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   2520
      ScaleHeight     =   360
      ScaleWidth      =   1560
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "Modify Key"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Key"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Key"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin ComctlLib.ListView ListView 
      Height          =   4815
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "ImageList"
      SmallIcons      =   "ImageList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Data"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.TabStrip TabStrip 
      Height          =   5775
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10186
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Current User"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Current User"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "All Users"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "All Users"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList 
      Left            =   3120
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5C12
            Key             =   "Default"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
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

Private Type KeyInfo
    Name As String
    Value As String
    PLocation As String
End Type

Dim clsReg As clsRegistry
Dim RegKey() As KeyInfo
Dim isKeyEmpty As Boolean

Private IniFile As New clsIniFile

Public Sub Assign_Array(vKey As Variant, vValue As Variant)
    Dim i As Integer
    
    If Not isKeyEmpty Then
        ReDim RegKey(UBound(vKey))
        For i = LBound(RegKey) To UBound(RegKey)
            RegKey(i).Name = vKey(i)
            RegKey(i).Value = vValue(i)
            RegKey(i).PLocation = ""
        Next i
    End If
End Sub

Public Sub Display_Keys()
    Dim i As Integer
    Dim Item As ListItem
    Dim tmpPath As String
    
    ListView.ListItems.Clear
    
    If isKeyEmpty Then
        ListView.ListItems.Add , , KEY_EMPTY
    Else
        For i = LBound(RegKey) To UBound(RegKey)
            Set Item = ListView.ListItems.Add(, , RegKey(i).Name, "Default", "Default")
                        
            Item.SubItems(1) = RegKey(Item.Index - 1).Value
            
            tmpPath = RegKey(Item.Index - 1).Value
            
            If Mid(tmpPath, 1, 1) = """" Then
                Dim a As Integer
                
                a = 2
                
                Do Until a = Len(tmpPath)
                    If Mid(tmpPath, a, 1) = """" Then
                        Exit Do
                    Else
                        a = a + 1
                    End If
                Loop
                
                tmpPath = Mid(tmpPath, 2, a - 2)
            End If
            
            If UCase(Mid(tmpPath, 1, 8)) = "RUNDLL32" Then
                tmpPath = "c:\windows\system32\rundll32.exe"
            End If
            
            tmpPath = Replace(tmpPath, "\\", "\")

            If FileExists(tmpPath) = True Then
                Item.Icon = ExtractIcon(tmpPath, ImageList, picTemp, 16)
                Item.SmallIcon = ExtractIcon(tmpPath, ImageList, picTemp, 16)
'            Else
'                Item.Icon = "Default"
'                Item.SmallIcon = "Default"
            End If
        Next i
        ListView.ListItems.Item(1).Selected = True
        
        If IniFile.ReadFrom("Other", "Alert", "1") = 1 Then
'            Dim tmp As Integer
'            Dim tmpCat As String
'
'            tmp = 1
'
'            If TabStrip.SelectedItem.Index = 1 Then
'                tmpCat = "CurrentUser"
'            Else
'                tmpCat = "AllUsers"
'            End If
'
'            Do Until tmp = ListView.ListItems.Count + 1
'                If IniFile.ReadFrom(tmpCat, frmMain.ListView.ListItems.Item(tmp).Text, "DEFAULT") = "DEFAULT" Then
'                    ListView.ListItems.Item(tmp).Ghosted = True
'                End If
'
'                tmp = tmp + 1
'            Loop
            Dim tmp1 As Integer
            Dim tmp2 As Integer
            
            tmp1 = 1
            
            Do Until tmp1 = ListView.ListItems.Count + 1
                tmp2 = 0
                
                Do Until tmp2 = frmMenu.List.ListCount
                    If ListView.ListItems.Item(tmp1).Text = frmMenu.List.List(tmp2) Then ListView.ListItems.Item(tmp1).Ghosted = True
                    
                    tmp2 = tmp2 + 1
                Loop
                
                tmp1 = tmp1 + 1
            Loop
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Me.Hide
    frmMenu.Timer.Enabled = False
End Sub

Public Sub cmdCreate_Click()
    If IniFile.ReadFrom("Security", "Creation", "0") = 1 Then
        With frmCreate
            If TabStrip.SelectedItem.Index = 1 Then
                .txtData.Tag = &H80000001
            Else
                .txtData.Tag = &H80000002
            End If
            
            .Show vbModal, Me
        End With
    Else
        MsgBox "You cannot Create registry keys.", vbOKOnly + vbCritical, "Error"
    End If
End Sub

Public Sub cmdDelete_Click()
    If IniFile.ReadFrom("Security", "Deletion", "0") = 1 Then
        If MsgBox("Are you sure you want to Delete the selected key ?", vbYesNo + vbQuestion, "Delete Key") = vbYes Then
            Dim tmp As Long
            
            If TabStrip.SelectedItem.Index = 1 Then
                tmp = &H80000001
            Else
                tmp = &H80000002
            End If
        
            DeleteValue tmp, "Software\Microsoft\Windows\CurrentVersion\Run", ListView.SelectedItem.Text
            
            TabStrip_Click
        End If
    Else
        MsgBox "You cannot Delete registry keys.", vbOKOnly + vbCritical, "Error"
    End If
End Sub

Public Sub cmdModify_Click()
    If IniFile.ReadFrom("Security", "Modification", "0") = 1 Then
        With frmModify
            .txtName.Text = ListView.SelectedItem.Text
            .txtName.Tag = ListView.SelectedItem.Text
            .txtData.Text = ListView.SelectedItem.SubItems(1)
            
            If TabStrip.SelectedItem.Index = 1 Then
                .txtData.Tag = &H80000001
            Else
                .txtData.Tag = &H80000002
            End If
            
            .Show vbModal, Me
        End With
    Else
        MsgBox "You cannot Modify registry keys.", vbOKOnly + vbCritical, "Error"
    End If
End Sub

Private Sub Form_Load()
    ListView.ColumnHeaders.Item(1).Width = 1440
    
    TabStrip_Click
End Sub

Public Function KeyExists(ByVal strKey As String, ByVal strLoc As String) As Boolean
On Error GoTo errReport
    Dim vKeyArray() As Variant
    Dim vValueArray() As Variant
    Dim errOccur As Boolean
    Dim isKeyOK As Boolean
    
    errOccur = False
    
    Set clsReg = New clsRegistry
    
    If strLoc = "HKEY_CURRENT_USER" Then
        clsReg.hKey = HKEY_CURRENT_USER
    Else
        clsReg.hKey = HKEY_LOCAL_MACHINE
    End If
    
    clsReg.SubKey = "Software\Microsoft\Windows\CurrentVersion\Run"
    
    If Not clsReg.Reg_OpenKey Then GoTo errHandler

    isKeyOK = clsReg.GetAllValues(vKeyArray, vValueArray)
        
    If isKeyOK Then
        isKeyEmpty = False
        
        Dim i As Integer
        
        If Not isKeyEmpty Then
            ReDim RegKey(UBound(vKeyArray))
            For i = LBound(RegKey) To UBound(RegKey)
                RegKey(i).Name = vKeyArray(i)
                RegKey(i).Value = vValueArray(i)
                RegKey(i).PLocation = ""
            Next i
        End If
    Else
        isKeyEmpty = True
    End If
    
    Dim a As Integer
    
    If isKeyEmpty Then
        KeyExists = False
        Exit Function
    Else
        For i = LBound(RegKey) To UBound(RegKey)
            If UCase(RegKey(i).Name) = UCase(strKey) Then
                KeyExists = True
                Exit Function
            End If
        Next i
        
        KeyExists = False
        Exit Function
    End If
    
CleanUp:
    On Error Resume Next
    clsReg.Reg_CloseKey
    Set clsReg = Nothing
'    If errOccur Then End'
    Exit Function

errHandler:
    On Error GoTo errReport
    Err.Raise vbObjectError, , "Error working/opening registry key" & vbCr & vbCr & "Ending program."
errReport:
    MsgBox Err.Number & " - " & Err.Description
    errOccur = True
    GoTo CleanUp
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim Result, Action As Long

    Action = x / Screen.TwipsPerPixelX

    If Action = WM_RBUTTONUP Then
        Result = SetForegroundWindow(Me.hWnd)
        PopupMenu frmMenu.mnuFile
    Else
    
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = 1
        Me.Hide
        frmMenu.Timer.Enabled = False
    Else
'        End'
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        If Me.Width >= 8500 Then
            TabStrip.Width = Me.ScaleWidth - 240
            ListView.Width = Me.ScaleWidth - 480
            cmdClose.Left = Me.ScaleWidth - 240 - 1215
            ListView.ColumnHeaders.Item(2).Width = ListView.Width - ListView.ColumnHeaders.Item(1).Width - 660
        Else
            Me.Width = 8500
        End If
        
        If Me.Height >= 6000 Then
            TabStrip.Height = Me.ScaleHeight - 240
            ListView.Height = Me.ScaleHeight - 1200
            
            cmdCreate.Top = Me.ScaleHeight - 620
            cmdDelete.Top = cmdCreate.Top
            cmdModify.Top = cmdCreate.Top
            cmdClose.Top = cmdCreate.Top
        Else
            Me.Height = 6000
        End If
    End If
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    Dim fso As New FileSystemObject
'    Dim strm As TextStream
'    Dim File As File
'    Dim vKeyArray() As Variant
'    Dim vValueArray() As Variant
'    Dim errOccur As Boolean
'    Dim isKeyOK As Boolean
'    Dim a As Integer
'    Dim tmpPath As String
'
'    If fso.FileExists(App.Path & "\BNCUSF") = True Then fso.DeleteFile App.Path & "\BNCUSF", True
'    If fso.FileExists(App.Path & "\BNAUSF") = True Then fso.DeleteFile App.Path & "\BNAUSF", True
'
'    Set strm = fso.CreateTextFile(App.Path & "\BNCUSF", True)
'
'On Error GoTo errReport
'    errOccur = False
'
'    Set clsReg = New clsRegistry
'
'    clsReg.hKey = HKEY_CURRENT_USER
'    clsReg.SubKey = "Software\Microsoft\Windows\CurrentVersion\Run"
'
'    If Not clsReg.Reg_OpenKey Then GoTo errHandler
'
'    isKeyOK = clsReg.GetAllValues(vKeyArray, vValueArray)
'
'    If isKeyOK Then
'        isKeyEmpty = False
'
'        Dim i As Integer
'
'        If Not isKeyEmpty Then
'            ReDim RegKey(UBound(vKeyArray))
'            For i = LBound(RegKey) To UBound(RegKey)
'                RegKey(i).Name = vKeyArray(i)
'                RegKey(i).Value = vValueArray(i)
'                RegKey(i).PLocation = ""
'            Next i
'        End If
'    Else
'        isKeyEmpty = True
'    End If
'
'    If isKeyEmpty Then
'        strm.WriteLine "NONE"
'        strm.WriteLine "NONE"
'    Else
'        For a = LBound(RegKey) To UBound(RegKey)
'            strm.WriteLine RegKey(a).Name
'            strm.WriteLine RegKey(a).Value
'        Next a
'    End If
'
'    strm.Close
'
'    Set strm = fso.CreateTextFile(App.Path & "\BNAUSF", True)
'
'    errOccur = False
'
'    Set clsReg = New clsRegistry
'
'    clsReg.hKey = HKEY_CURRENT_USER
'    clsReg.SubKey = "Software\Microsoft\Windows\CurrentVersion\Run"
'
'    If Not clsReg.Reg_OpenKey Then GoTo errHandler
'
'    isKeyOK = clsReg.GetAllValues(vKeyArray, vValueArray)
'
'    If isKeyOK Then
'        isKeyEmpty = False
'
'        i = 0
'
'        If Not isKeyEmpty Then
'            ReDim RegKey(UBound(vKeyArray))
'            For i = LBound(RegKey) To UBound(RegKey)
'                RegKey(i).Name = vKeyArray(i)
'                RegKey(i).Value = vValueArray(i)
'                RegKey(i).PLocation = ""
'            Next i
'        End If
'    Else
'        isKeyEmpty = True
'    End If
'
'    If isKeyEmpty Then
'        strm.WriteLine "NONE"
'        strm.WriteLine "NONE"
'    Else
'        For a = LBound(RegKey) To UBound(RegKey)
'            strm.WriteLine RegKey(a).Name
'            strm.WriteLine RegKey(a).Value
'        Next a
'    End If
'
'    strm.Close
'
'    Set File = fso.GetFile(App.Path & "\BNCUSF")
'    File.Attributes = System + Hidden
'    Set File = fso.GetFile(App.Path & "\BNAUSF")
'    File.Attributes = System + Hidden
'
'Cleanup:
'    On Error Resume Next
'    clsReg.Reg_CloseKey
'    Set clsReg = Nothing
'    If errOccur Then End
'    Exit Sub
'
'errHandler:
'    On Error GoTo errReport
'    Err.Raise vbObjectError, , "Error working/opening registry key" & vbCr & vbCr & "Ending program."
'errReport:
'    MsgBox Err.Number & " - " & Err.Description
'    errOccur = True
'    GoTo Cleanup
'End Sub

Private Sub ListView_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    If ListView.SortKey = ColumnHeader.Index - 1 Then
        If ListView.SortOrder = lvwAscending Then ListView.SortOrder = lvwDescending Else ListView.SortOrder = lvwAscending
    Else
        ListView.SortKey = ColumnHeader.Index - 1
        ListView.SortOrder = lvwAscending
    End If
    
    ListView.Sorted = True
End Sub

Private Sub ListView_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then cmdDelete_Click
End Sub

Private Sub ListView_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then PopupMenu frmMenu.mnuEdit
End Sub

Public Sub TabStrip_Click()
On Error GoTo errReport
    Dim vKeyArray() As Variant
    Dim vValueArray() As Variant
    Dim errOccur As Boolean
    Dim isKeyOK As Boolean
    
    errOccur = False
    
    Set clsReg = New clsRegistry
    
    If TabStrip.SelectedItem.Index = 1 Then
        clsReg.hKey = HKEY_CURRENT_USER
    Else
        clsReg.hKey = HKEY_LOCAL_MACHINE
    End If
    
    clsReg.SubKey = "Software\Microsoft\Windows\CurrentVersion\Run"
    
    If Not clsReg.Reg_OpenKey Then GoTo errHandler

    isKeyOK = clsReg.GetAllValues(vKeyArray, vValueArray)
        
    If isKeyOK Then
        isKeyEmpty = False
        Call Assign_Array(vKeyArray, vValueArray)
        Call Display_Keys
    Else
        isKeyEmpty = True
        Call Display_Keys
    End If
    
CleanUp:
    On Error Resume Next
    clsReg.Reg_CloseKey
    Set clsReg = Nothing
'    If errOccur Then End'
    Exit Sub

errHandler:
    On Error GoTo errReport
    Err.Raise vbObjectError, , "Error working/opening registry key" & vbCr & vbCr & "Ending program."
errReport:
    MsgBox Err.Number & " - " & Err.Description
    errOccur = True
    GoTo CleanUp
End Sub
