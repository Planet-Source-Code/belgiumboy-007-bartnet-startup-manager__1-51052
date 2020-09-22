VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2400
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRestoreDefaults 
      Caption         =   "Restore Defaults"
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check For Update"
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh Options"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Options"
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   "Other"
      Height          =   1575
      Left            =   120
      TabIndex        =   21
      Top             =   1800
      Width           =   3975
      Begin VB.TextBox txtOtherSpeed 
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   1080
         Width           =   1400
      End
      Begin VB.CheckBox ckOtherAlert 
         Caption         =   "Alert Me When New Keys Have Been Created"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   3615
      End
      Begin VB.CheckBox ckOtherRun 
         Caption         =   "Run This Program When Windows Starts"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "Refresh Speed (miliseconds) :"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1100
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Security"
      Height          =   1575
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox ckSecurityModification 
         Caption         =   "Enable Key Modification"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox ckSecurityDeletion 
         Caption         =   "Enable Key Deletion"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   2055
      End
      Begin VB.CheckBox ckSecurityCreation 
         Caption         =   "Enable Key Creation"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ListView"
      Height          =   1575
      Left            =   2760
      TabIndex        =   17
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox cboListAppearance 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cboListBorder 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox ckListColumnHeaders 
         Caption         =   "Show Column Headers"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Appearance :"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1115
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "BorderStyle :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   755
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "TabStrip"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   2175
      Begin VB.ComboBox cboTabStyle 
         Height          =   315
         ItemData        =   "frmOptions.frx":0000
         Left            =   720
         List            =   "frmOptions.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Style :"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   270
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmOptions"
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

Private IniFile As New clsIniFile

Private Sub Change()
    cmdSave.Enabled = True
    cmdRefresh.Enabled = True
End Sub

Private Sub cboListAppearance_Click()
    Change
End Sub

Private Sub cboListBorder_Click()
    Change
End Sub

Private Sub cboTabStyle_Click()
    Change
End Sub

Private Sub ckListColumnHeaders_Click()
    Change
End Sub

Private Sub ckOtherAlert_Click()
    Change
End Sub

Private Sub ckOtherRun_Click()
    Change
End Sub

Private Sub ckSecurityCreation_Click()
    Change
End Sub

Private Sub ckSecurityDeletion_Click()
    Change
End Sub

Private Sub ckSecurityModification_Click()
    Change
End Sub

Private Sub cmdClose_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    LoadSettings
    
    cmdSave.Enabled = False
    cmdRefresh.Enabled = False
End Sub

Private Sub cmdRestoreDefaults_Click()
    ckSecurityCreation.Value = 0
    ckSecurityDeletion.Value = 0
    ckSecurityModification.Value = 0
    
    ckListColumnHeaders.Value = 1
    cboListBorder.Text = "Fixed Single"
    cboListAppearance.Text = "3D"
    
    ckOtherRun.Value = 1
    ckOtherAlert.Value = 1
    txtOtherSpeed = "5000"
    
    cboTabStyle.Text = "Tabs"
    
    cmdSave_Click
End Sub

Private Sub cmdSave_Click()
    Dim tmp As String
    
    If Trim(txtOtherSpeed.Text) = "" Then tmp = "NONE" Else tmp = txtOtherSpeed.Text
    
    IniFile.WriteTo "Security", "Creation", ckSecurityCreation.Value
    IniFile.WriteTo "Security", "Deletion", ckSecurityDeletion.Value
    IniFile.WriteTo "Security", "Modification", ckSecurityModification.Value
    
    IniFile.WriteTo "ListView", "ColumnHeaders", ckListColumnHeaders.Value
    IniFile.WriteTo "ListView", "BorderStyle", cboListBorder.Text
    IniFile.WriteTo "ListView", "Appearance", cboListAppearance.Text
    
    IniFile.WriteTo "Other", "Run", ckOtherRun.Value
    IniFile.WriteTo "Other", "Alert", ckOtherAlert.Value
    IniFile.WriteTo "Other", "Speed", tmp
    
    IniFile.WriteTo "TabStrip", "Style", cboTabStyle.Text
    
    cmdSave.Enabled = False
    cmdRefresh.Enabled = False
    
    ApplySettings
End Sub

Private Sub Command1_Click()
    frmUpdate.Show vbModal, Me
End Sub

Private Sub Form_Load()
    cboTabStyle.AddItem "Tabs"
    cboTabStyle.AddItem "Buttons"
    
    cboListBorder.AddItem "None"
    cboListBorder.AddItem "Fixed Single"
    
    cboListAppearance.AddItem "Flat"
    cboListAppearance.AddItem "3D"
    
    LoadSettings
    
    cmdSave.Enabled = False
    cmdRefresh.Enabled = False
End Sub

Private Sub txtOtherSpeed_Change()
    Change
End Sub
