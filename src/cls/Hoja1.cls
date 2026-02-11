VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdBD_Click()
frmBD.Show
End Sub

Private Sub cmdInicio_Click()
Application.Visible = False
frmDash.Show
End Sub
