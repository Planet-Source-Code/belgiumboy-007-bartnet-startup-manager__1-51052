VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check For Update"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   240
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   360
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   720
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text 
      Height          =   1695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   360
      Width           =   3735
   End
   Begin ComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Min             =   1e-4
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Progress :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Status :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmUpdate"
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

Private NewVersion As String
Private FileSize As String
Private FileName As String
Private URL As String
Private FileLocation As String

Private Sub cmdCancel_Click()
    If MsgBox("Are you sure you want to Cancel ?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
        Me.Hide
        Unload Me
    End If
End Sub

Private Sub cmdStart_Click()
    Text.Text = Text.Text & "Connecting to www.bartnet.be" & vbCrLf

    cmdStart.Enabled = False
    Label1.Enabled = True
    Text.Enabled = True
    Text.Text = ""
    Label2.Enabled = True
    ProgressBar.Enabled = True
    
    ProgressBar.Min = 0
    ProgressBar.Max = 12
    ProgressBar.Value = 0
    
    Timer.Enabled = True
End Sub

Private Sub Form_Load()
    Label1.Enabled = False
    Text.Enabled = False
    Label2.Enabled = False
    ProgressBar.Enabled = False
End Sub

Private Sub DownloadFile(ByVal strFile As String, ByVal strDest As String)
    Dim myData() As Byte
    
    If FileExists(strDest) = True Then Kill strDest
    
    myData() = Inet.OpenURL(strFile, icByteArray)
    
    Open strDest For Binary Access Write As #1
    Put #1, , myData()
    Close #1
    
    Inet.Cancel
End Sub

Private Sub Timer_Timer()
    Timer.Enabled = False
    
'    DownloadFile "http://www.bartnet.be/Other Files/Programs/BartNet Startup Manager/FILELOCATION.BartNet", App.Path & "\THISISTHETESTRESULT.TXT"
    Text.Text = Text.Text & "Downloading information File #1" & vbCrLf
    DownloadFile "http://www.bartnet.be/Other Files/Programs/BartNet Startup Manager/NEWVERSION.BartNet", App.Path & "\NEWVERSION.BartNet"
    ProgressBar.Value = ProgressBar.Value + 1
    Text.Text = Text.Text & "Downloading information File #2" & vbCrLf
    DownloadFile "http://www.bartnet.be/Other Files/Programs/BartNet Startup Manager/FILESIZE.BartNet", App.Path & "\FILESIZE.BartNet"
    ProgressBar.Value = ProgressBar.Value + 1
    Text.Text = Text.Text & "Downloading information File #3" & vbCrLf
    DownloadFile "http://www.bartnet.be/Other Files/Programs/BartNet Startup Manager/FILENAME.BartNet", App.Path & "\FILENAME.BartNet"
    ProgressBar.Value = ProgressBar.Value + 1
    Text.Text = Text.Text & "Downloading information File #4" & vbCrLf
    DownloadFile "http://www.bartnet.be/Other Files/Programs/BartNet Startup Manager/URL.BartNet", App.Path & "\URL.BartNet"
    ProgressBar.Value = ProgressBar.Value + 1
    Text.Text = Text.Text & "Downloading information File #5" & vbCrLf
    DownloadFile "http://www.bartnet.be/Other Files/Programs/BartNet Startup Manager/FILELOCATION.BartNet", App.Path & "\FILELOCATION.BartNet"
    ProgressBar.Value = ProgressBar.Value + 1
    Text.Text = Text.Text & "Processing information" & vbCrLf
    
    Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    
    Dim Lenght1, Lenght2, Lenght3, Lenght4, Lenght5 As Long
    Dim Counter As Long
    Dim bContent() As Byte

    Lenght1 = FileLen(App.Path & "\NEWVERSION.BartNet")
    Lenght2 = FileLen(App.Path & "\FILESIZE.BartNet")
    Lenght3 = FileLen(App.Path & "\FILENAME.BartNet")
    Lenght4 = FileLen(App.Path & "\URL.BartNet")
    Lenght5 = FileLen(App.Path & "\FILELOCATION.BartNet")
       
    Open App.Path & "\NEWVERSION.BartNet" For Binary Access Read As #1
    Open App.Path & "\FILESIZE.BartNet" For Binary Access Read As #2
    Open App.Path & "\FILENAME.BartNet" For Binary Access Read As #3
    Open App.Path & "\URL.BartNet" For Binary Access Read As #4
    Open App.Path & "\FILELOCATION.BartNet" For Binary Access Read As #5
    
    ProgressBar.Value = ProgressBar.Value + 1
    
    ReDim bContent(Lenght1): NewVersion = Space$(Lenght1)
    
    For Counter = 1 To Lenght1
        Get #1, Counter, bContent(Counter)
        Mid$(NewVersion, Counter, 1) = Chr$(bContent(Counter))
    Next
    
    ProgressBar.Value = ProgressBar.Value + 1
    
    ReDim bContent(Lenght2): FileSize = Space$(Lenght2)
    
    For Counter = 1 To Lenght2
        Get #2, Counter, bContent(Counter)
        Mid$(FileSize, Counter, 1) = Chr$(bContent(Counter))
    Next
    
    ProgressBar.Value = ProgressBar.Value + 1
    
    ReDim bContent(Lenght3): FileName = Space$(Lenght3)
    
    For Counter = 1 To Lenght3
        Get #3, Counter, bContent(Counter)
        Mid$(FileName, Counter, 1) = Chr$(bContent(Counter))
    Next
    
    ProgressBar.Value = ProgressBar.Value + 1
    
    ReDim bContent(Lenght4): URL = Space$(Lenght4)
    
    For Counter = 1 To Lenght4
        Get #4, Counter, bContent(Counter)
        Mid$(URL, Counter, 1) = Chr$(bContent(Counter))
    Next
    
    ProgressBar.Value = ProgressBar.Value + 1
    
    ReDim bContent(Lenght5): FileLocation = Space$(Lenght5)
    
    For Counter = 1 To Lenght5
        Get #5, Counter, bContent(Counter)
        Mid$(FileLocation, Counter, 1) = Chr$(bContent(Counter))
    Next
      
    ProgressBar.Value = ProgressBar.Value + 1
    
    Close #1
    Close #2
    Close #3
    Close #4
    Close #5
    
    ProgressBar.Value = ProgressBar.Value + 1
    
    If NewVersion <> "1.0.0" Then
        If MsgBox("There is a new version of BartNet Startup Manager available for downloading.  This program can download the setup File for you automatically.  File details follow : " & vbCrLf & "Version : " & NewVersion & vbCrLf & "Name : " & FileName & vbCrLf & "Size : " & FileSize & vbCrLf & vbCrLf & "Do you want BartNet Startup Manager to download the new version automatically ?", vbYesNo + vbQuestion, "Results") = vbYes Then
            With frmDownload
                Me.Hide
                
                .Show
                
                .DownloadFile FileLocation, App.Path & "\" & FileName
                
                Unload Me
            End With
        Else
            If MsgBox("Would you like to be taken to the website where the new version can be downloaded ?", vbYesNo + vbQuestion, "Results") = vbYes Then ShellExecute2 Me.hWnd, vbNullString, URL, vbNullString, "c:\", SW_SHOWNORMAL
            
            Me.Hide
            Unload Me
        End If
    Else
        MsgBox "You are currently using the newest version of BartNet Startup Manager", vbOKOnly + vbInformation, "Results"
        
        Me.Hide
        Unload Me
    End If
    
'    MsgBox "Details :" & vbCrLf & vbCrLf & _
'        "Version : " & NewVersion & vbCrLf & _
'        "Size : " & FileSize & vbCrLf & _
'        "Name : " & FileName & vbCrLf & _
'        "URL : " & URL & vbCrLf & _
'        "Location : " & FileLocation, vbOKOnly + vbInformation, "Results"
End Sub
