VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Agregar como cajero"
   ClientHeight    =   2775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2895
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProcess_Click()
Dim oData As Object
Dim sql As String

If ValidateFields(Me) = False Then Exit Sub

If Me.cmdProcess.Caption = "Crear" Then
  sql = "INSERT INTO cashiers (cashier, serialNumber) VALUES ('" & Me.txtCashier.Text & "', '" & Me.lblSerial.Caption & "')"
Else
  sql = "UPDATE cashiers SET cashier='" & Me.txtCashier.Text & "' WHERE serialNumber='" & Me.lblSerial.Caption & "'"
End If

Set oData = ExecuteQuery(sql)
Set oData = Nothing

Hoja2.Cells(5, 2) = Me.txtCashier.Text

MsgBox "El proceso se realizó exitosamente"

If Me.cmdProcess.Caption = "Crear" Then
  Unload Me
  frmLogin.Show
End If
End Sub

Private Sub UserForm_Initialize()
Dim oData As Object
Dim sql As String

FormDesign Me
Me.lblSerial.Caption = GetSerialNumber

sql = "SELECT cashier FROM cashiers WHERE serialNumber='" & GetSerialNumber & "' AND idState<>3"
  Set oData = ExecuteQuery(sql)
  
If Not oData.EOF Then
  Me.txtCashier.Text = oData.Fields("cashier")
  Me.cmdProcess.Caption = "Actualizar"
Else
  Me.cmdProcess.Caption = "Crear"
End If
End Sub
