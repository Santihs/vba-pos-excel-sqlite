VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMe 
   Caption         =   "Wilfredo HQ"
   ClientHeight    =   8430.001
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5775
   OleObjectBlob   =   "frmMe.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frmMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btn_donar_Click()
On Error Resume Next
Dim ChromeLocation As String
ChromeLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
MyURL = "https://www.paypal.me/QuispeWilfredo"
Shell (ChromeLocation & " -url " & MyURL)
End Sub

Private Sub btn_donar_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
InputBox "Copia la URL y pega en un navegador.", "Wilfredo HQ", "https://www.paypal.me/QuispeWilfredo"
End Sub

Private Sub btn_donar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub ico_facebook_Click()
On Error Resume Next
Dim ChromeLocation As String
ChromeLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
MyURL = "https://web.facebook.com/HQWilfredo"
Shell (ChromeLocation & " -url " & MyURL)
End Sub

Private Sub ico_facebook_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
InputBox "Copia la URL y pega en un navegador.", "Wilfredo HQ", "https://web.facebook.com/HQWilfredo"
End Sub

Private Sub ico_facebook_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub ico_instagram_Click()
On Error Resume Next
Dim ChromeLocation As String
ChromeLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
MyURL = "https://www.instagram.com/HQWilfredo"
Shell (ChromeLocation & " -url " & MyURL)
End Sub

Private Sub ico_instagram_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
InputBox "Copia la URL y pega en un navegador.", "Wilfredo HQ", "https://www.instagram.com/HQWilfredo"
End Sub

Private Sub ico_instagram_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub ico_twitter_Click()
On Error Resume Next
Dim ChromeLocation As String
ChromeLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
MyURL = "https://twitter.com/HQWilfredo"
Shell (ChromeLocation & " -url " & MyURL)
End Sub

Private Sub ico_twitter_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
InputBox "Copia la URL y pega en un navegador.", "Wilfredo HQ", "https://twitter.com/HQWilfredo"
End Sub

Private Sub ico_twitter_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub ico_youtube_Click()
On Error Resume Next
Dim ChromeLocation As String
ChromeLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
MyURL = "https://www.youtube.com/c/WilfredoHQ"
Shell (ChromeLocation & " -url " & MyURL)
End Sub

Private Sub ico_youtube_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
InputBox "Copia la URL y pega en un navegador.", "Wilfredo HQ", "https://www.youtube.com/c/WilfredoHQ"
End Sub

Private Sub ico_youtube_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
SetCursor LoadCursor(0, IDC_HAND)
End Sub

