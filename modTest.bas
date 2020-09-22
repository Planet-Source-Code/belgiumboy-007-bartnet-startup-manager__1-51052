Attribute VB_Name = "modTest"
' This Sub is strictly for the purpose of demonstrating
' how to use the code.

Sub Main()

' Test Sub

' This is all there is to it...
'frmDownload.DownloadFile "http://www.virtualcd.com/fastform/00005003.pdf", _
'                         "c:\windows\desktop\00005003.pdf"
frmDownload.DownloadFile "http://www.bartnet.be/Winzip Files/MSN 6.0 BETA.zip", App.Path & "\FILE.zip"
End Sub


