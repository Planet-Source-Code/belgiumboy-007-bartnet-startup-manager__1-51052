VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDownload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Downloading new version"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   ControlBox      =   0   'False
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5655
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4080
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2220
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   615
      Left            =   480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      _Version        =   393216
      FullWidth       =   313
      FullHeight      =   41
   End
   Begin VB.Label StatusLabel 
      Caption         =   "StatusLabel"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   5235
   End
   Begin VB.Label EstimatedTimeLeft 
      Caption         =   "Estimated time left:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label SourceLabel 
      Caption         =   "SourceLabel"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5250
   End
   Begin VB.Label TimeLabel 
      Caption         =   "TimeLabel"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   3765
   End
End
Attribute VB_Name = "frmDownload"
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

Private CancelSearch As Boolean

Private Declare Function GetDiskFreeSpace Lib "kernel32" _
                         Alias "GetDiskFreeSpaceA" _
                         (ByVal lpRootPathName As String, _
                          lpSectorsPerCluster As Long, _
                          lpBytesPerSector As Long, _
                          lpNumberOfFreeClusters As Long, _
                          lpTotalNumberOfClusters As Long) As Long

Public Function DownloadFile(strURL As String, strDestination As String, Optional UserName As String = Empty, Optional Password As String = Empty) As Boolean
    Const CHUNK_SIZE As Long = 1024
    Const ROLLBACK As Long = 4096

    Dim bData() As Byte
    Dim blnResume As Boolean
    Dim intFile As Integer
    Dim lngBytesReceived As Long
    Dim lngFileLength As Long
    Dim lngX
    Dim sglLastTime As Single
    Dim sglRate As Single
    Dim sglTime As Single
    Dim strFile As String
    Dim strHeader As String
    Dim strHost As String
    
On Local Error GoTo InternetErrorHandler
    
    CancelSearch = False

    strFile = ReturnFileOrFolder(strDestination, True)
    strHost = ReturnFileOrFolder(strURL, True, True)

    SourceLabel = Empty
    TimeLabel = Empty
    ToLabel = Empty
    RateLabel = Empty

    With Animation1
        .AutoPlay = True
        .Open App.Path & "\DOWNLD2.AVI"
    End With

    Show

    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

StartDownload:

    If blnResume Then
        StatusLabel = "Resuming download..."
        lngBytesReceived = lngBytesReceived - ROLLBACK
        If lngBytesReceived < 0 Then lngBytesReceived = 0
    Else
        StatusLabel = "Getting file information..."
    End If

    DoEvents
    
    With Inet1
        .URL = strURL
        .UserName = UserName
        .Password = Password
    
        .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
    
    
        While .StillExecuting
            DoEvents
            If CancelSearch Then GoTo ExitDownload
        Wend

        StatusLabel = "Saving:"
        SourceLabel = FitText(SourceLabel, strHost & " from " & .RemoteHost)

        strHeader = .GetHeader
    End With
    
    Select Case Mid(strHeader, 10, 3)
        Case "200"
            If blnResume Then
                Kill strDestination
                If MsgBox("The server is unable to resume this download." & vbCr & vbCr & "Do you want to continue anyway?", vbExclamation + vbYesNo, "Unable to Resume Download") = vbYes Then
                    blnResume = False
                Else
                    CancelSearch = True
                    GoTo ExitDownload
                End If
            End If
        Case "206"
        Case "204"
            MsgBox "Nothing to download!", vbInformation, "No Content"
            CancelSearch = True
            GoTo ExitDownload
        Case "401"
            MsgBox "Authorization failed!", vbCritical, "Unauthorized"
            CancelSearch = True
            GoTo ExitDownload
        Case "404"
            MsgBox "The file, " & """" & Inet1.URL & """" & " was not found!", vbCritical, "File Not Found"
            CancelSearch = True
            GoTo ExitDownload
        Case vbCrLf
            MsgBox "Cannot establish connection." & vbCr & vbCr & "Check your Internet connection and try again.", vbExclamation, "Cannot Establish Connection"
            CancelSearch = True
            GoTo ExitDownload
        Case Else
            strHeader = Left(strHeader, InStr(strHeader, vbCr))
            If strHeader = Empty Then strHeader = "<nothing>"
            MsgBox "The server returned the following response:" & vbCr & vbCr & strHeader, vbCritical, "Error Downloading File"
            CancelSearch = True
            GoTo ExitDownload
    End Select

    If blnResume = False Then
        sglLastTime = Timer - 1
        strHeader = Inet1.GetHeader("Content-Length")
        lngFileLength = Val(strHeader)
        If lngFileLength = 0 Then
            GoTo ExitDownload
        End If
    End If

    If Mid(strDestination, 2, 2) = ":\" Then
        If DiskFreeSpace(Left(strDestination, InStr(strDestination, "\"))) < lngFileLength Then
            MsgBox "There is not enough free space on disk for this file." & vbCr & vbCr & "Please free up some disk space and try again.", vbCritical, "Insufficient Disk Space"
            GoTo ExitDownload
        End If
    End If

    With ProgressBar
        .Value = 0
        .Max = lngFileLength
    End With

    DoEvents
    
    If blnResume = False Then lngBytesReceived = 0

On Local Error GoTo FileErrorHandler

    strHeader = ReturnFileOrFolder(strDestination, False)
    If Dir(strHeader, vbDirectory) = Empty Then
        MkDir strHeader
    End If

    intFile = FreeFile()

    Open strDestination For Binary Access Write As #intFile

    If blnResume Then Seek #intFile, lngBytesReceived + 1
    Do
        bData = Inet1.GetChunk(CHUNK_SIZE, icByteArray)
        Put #intFile, , bData
        If CancelSearch Then Exit Do
        lngBytesReceived = lngBytesReceived + UBound(bData, 1) + 1
        sglRate = lngBytesReceived / (Timer - sglLastTime)
        sglTime = (lngFileLength - lngBytesReceived) / sglRate
        TimeLabel = FormatTime(sglTime) & " (" & FormatFileSize(lngBytesReceived) & " of " & FormatFileSize(lngFileLength) & " copied)"
        RateLabel = FormatFileSize(sglRate, "###.0") & "/Sec"
        ProgressBar.Value = lngBytesReceived
        Me.Caption = Format((lngBytesReceived / lngFileLength), "##0%") & " of " & strFile & " Completed"
    Loop While UBound(bData, 1) > 0

    Close #intFile

ExitDownload:

    If lngBytesReceived = lngFileLength Then
        StatusLabel = "Download completed!"
        DownloadFile = True
        If MsgBox("The File was downlaoded successfully, would you like to open it now ?", vbYesNo + vbQuestion, "Results") = vbYes Then
            ShellExecute2 Me.hWnd, vbNullString, strDestination, vbNullString, "c:\", SW_SHOWNORMAL
        Else
            MsgBox "The File has been saved to """ & strDestination & """.", vbOKOnly + vbInformation, "Results"
        End If
    Else
        If Dir(strDestination) = Empty Then
            CancelSearch = True
        Else
            If CancelSearch = False Then
                If MsgBox("The connection with the server was reset." & vbCr & vbCr & "Click ""Retry"" to resume downloading the file." & vbCr & "(Approximate time remaining: " & FormatTime(sglTime) & ")" & vbCr & vbCr & "Click ""Cancel"" to cancel downloading the file.", vbExclamation + vbRetryCancel, "Download Incomplete") = vbRetry Then
                    blnResume = True
                    GoTo StartDownload
                End If
            End If
        End If
        If Not Dir(strDestination) = Empty Then Kill strDestination
        DownloadFile = False
    End If

CleanUp:

    Animation1.Close

    Inet1.Cancel

    Unload Me

    Exit Function

InternetErrorHandler:
    
    If Err.Number = 9 Then Resume Next
    MsgBox "Error: " & Err.Description & " occurred.", vbCritical, "Error Downloading File"
    Err.Clear
    GoTo ExitDownload
    
FileErrorHandler:

    MsgBox "Cannot write file to disk." & vbCr & vbCr & "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error Downloading File"
    CancelSearch = True
    Err.Clear
    GoTo ExitDownload
End Function


Private Sub CancelButton_Click()
    If MsgBox("Are you sure you want to Cancel ?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
        StatusLabel = "Cancelling..."
        Animation1.Close
        CancelSearch = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Move -Width, -Height
    DoEvents
End Sub

Private Function FitText(ByRef Ctl As Control, ByVal strCtlCaption) As String
    Dim lngCtlLeft As Long
    Dim lngMaxWidth As Long
    Dim lngTextWidth As Long
    Dim lngX As Long

    lngCtlLeft = Ctl.Left
    lngMaxWidth = Ctl.Width
    lngTextWidth = Ctl.Parent.TextWidth(strCtlCaption)

    lngX = (Len(strCtlCaption) \ 2) - 2
    While lngTextWidth > lngMaxWidth And lngX > 3
        strCtlCaption = Left(strCtlCaption, lngX) & "..." & Right(strCtlCaption, lngX)
        lngTextWidth = Ctl.Parent.TextWidth(strCtlCaption)
        lngX = lngX - 1
    Wend

    FitText = strCtlCaption
End Function

Private Function FormatFileSize(ByVal dblFileSize As Double, Optional ByVal strFormatMask As String) As String
    Select Case dblFileSize
        Case 0 To 1023
            FormatFileSize = Format(dblFileSize) & " bytes"
        Case 1024 To 1048575
            If strFormatMask = Empty Then strFormatMask = "###0"
            FormatFileSize = Format(dblFileSize / 1024#, strFormatMask) & " KB"
        Case 1024# ^ 2 To 1073741823
            If strFormatMask = Empty Then strFormatMask = "###0.0"
            FormatFileSize = Format(dblFileSize / (1024# ^ 2), strFormatMask) & " MB"
        Case Is > 1073741823#
            If strFormatMask = Empty Then strFormatMask = "###0.0"
            FormatFileSize = Format(dblFileSize / (1024# ^ 3), strFormatMask) & " GB"
    End Select
End Function

Private Function FormatTime(ByVal sglTime As Single) As String
    Select Case sglTime
        Case 0 To 59
            FormatTime = Format(sglTime, "0") & " sec"
        Case 60 To 3599
            FormatTime = Format(Int(sglTime / 60), "#0") & " min " & Format(sglTime Mod 60, "0") & " sec"
        Case Else
            FormatTime = Format(Int(sglTime / 3600), "#0") & " hr " & Format(sglTime / 60 Mod 60, "0") & " min"
    End Select
End Function

Private Function DiskFreeSpace(strDrive As String) As Double
    Dim SectorsPerCluster As Long
    Dim BytesPerSector As Long
    Dim NumberOfFreeClusters As Long
    Dim TotalNumberOfClusters As Long
    Dim FreeBytes As Long
    Dim spaceInt As Integer

    strDrive = QualifyPath(strDrive)

    GetDiskFreeSpace strDrive, SectorsPerCluster, BytesPerSector, NumberOFreeClusters, TotalNumberOfClusters

    DiskFreeSpace = NumberOFreeClusters * SectorsPerCluster * BytesPerSector
End Function

Private Function QualifyPath(strPath As String) As String
    QualifyPath = IIf(Right(strPath, 1) = "\", strPath, strPath & "\")
End Function


Private Function ReturnFileOrFolder(FullPath As String, ReturnFile As Boolean, Optional IsURL As Boolean = False) As String
    Dim intDelimiterIndex As Integer

    intDelimiterIndex = InStrRev(FullPath, IIf(IsURL, "/", "\"))
    
    If intDelimiterIndex = 0 Then
        ReturnFileOrFolder = FullPath
    Else
        ReturnFileOrFolder = IIf(ReturnFile, Right(FullPath, Len(FullPath) - intDelimiterIndex), Left(FullPath, intDelimiterIndex))
    End If
End Function

