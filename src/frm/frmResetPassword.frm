VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmResetPassword 
   Caption         =   "Cambiar contraseña"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5415
   OleObjectBlob   =   "frmResetPassword.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frmResetPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUpdate_Click()
Dim oData As Object
Dim sql As String

If ValidateFields(Me) = False Then Exit Sub

If Me.txtNewPassword.Text <> Me.txtNewPasswordRepeat Then
MsgBox "Su contraseña nueva y la repetición no coinciden"
Exit Sub
End If

sql = "SELECT secretKey FROM users WHERE idEmployee=" & Hoja2.Cells(2, 1) & " AND idState<>3"
Set oData = ExecuteQuery(sql)

If Not oData.EOF Then
  If oData.Fields("secretKey") = SHA256(Me.txtOldPassword.Text) Then
    sql = "UPDATE users SET secretKey='" & SHA256(Me.txtNewPassword.Text) & "' WHERE idEmployee=" & Hoja2.Cells(2, 1)
    Set oData = ExecuteQuery(sql)
    Set oData = Nothing
  Else
    MsgBox "Su contraseña anterior es incorrecta"
    Exit Sub
  End If
End If

CleanControls Me
Me.txtOldPassword.SetFocus

MsgBox "Su contraseña ha sido actualizada exitosamente"

Unload Me
frmLogin.Show
End Sub

Private Sub UserForm_Initialize()
FormDesign Me
End Sub
