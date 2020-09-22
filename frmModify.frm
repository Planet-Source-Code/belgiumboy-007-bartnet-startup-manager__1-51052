VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmModify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Registry Key"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4695
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "Key Data :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Key Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmModify"
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

Private Sub cmdCancel_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtName.Text) = "" Or Trim(txtData.Text) = "" Then
        MsgBox "Please enter a Name and Data for the key.", vbOKOnly + vbCritical, "Error"
    Else
        If UCase(txtName.Tag) <> UCase(txtName.Text) Then
            Dim tmp As String
            
            If txtData.Tag = &H80000001 Then tmp = "HKEY_CURRENT_USER" Else tmp = "HKEY_LOCAL_MACHINE"
            
            If frmMain.KeyExists(txtName.Text, tmp) = True Then
                MsgBox "There is already a key with the Name you specified.", vbOKOnly + vbCritical, "Error"
                Exit Sub
            End If
        End If
        
        DeleteValue txtData.Tag, "Software\Microsoft\Windows\CurrentVersion\Run", txtName.Tag
        SetKeyValue txtData.Tag, "Software\Microsoft\Windows\CurrentVersion\Run", txtName.Text, txtData.Text, 1
    
        frmMain.TabStrip_Click
    
        cmdCancel_Click
    End If
End Sub
