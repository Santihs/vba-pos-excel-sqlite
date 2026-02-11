VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplash 
   Caption         =   "Cargando..."
   ClientHeight    =   855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5295
   OleObjectBlob   =   "frmSplash.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
Dim Contador As Integer, Maximo As Integer, Intervalo As Integer
Dim Inicio As Double
Dim X As Integer

FormDesign Me
Me.lblProgressBar.BackStyle = fmBackStyleOpaque
Me.lblProgressBar.BackColor = &HC000&

Maximo = 228
Me.Show

For Contador = 1 To Maximo
  Inicio = Timer
  Do Until Timer - Inicio > Intervalo
    X = DoEvents()
  Loop
  
  Me.lblProgressBar.Width = Contador
  Me.lblPercentage.Caption = Format(Contador / Maximo, "Percent")
Next Contador
End
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then
  Cancel = True
End If
End Sub

