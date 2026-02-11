VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Iniciar sesión"
   ClientHeight    =   2775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4335
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
  ThisWorkbook.Application.DisplayAlerts = False
  ThisWorkbook.Application.Quit
  Unload Me
End Sub

Private Sub cmdLogin_Click()

Dim oData As Object
Dim sql As String

If ValidateFields(Me) = False Then Exit Sub

sql = "SELECT idUser, employees.idEmployee, dni, name, surname, sales, orders, units, categories, products, customers, employees, positions, settings, database FROM employees INNER JOIN users ON employees.idEmployee = users.idEmployee INNER JOIN positions ON employees.idPosition = positions.idPosition WHERE email='" & Me.txtEmail.Text & "' AND secretKey='" & SHA256(Me.txtPassword.Text) & "' AND employees.idState<>3 AND users.idState<>3  AND positions.idState<>3"
Set oData = ExecuteQuery(sql)

If Not oData.EOF Then

  Hoja2.Cells(2, 1) = oData.Fields("idEmployee")
  Hoja2.Cells(2, 2) = oData.Fields("dni")
  Hoja2.Cells(2, 3) = oData.Fields("name")
  Hoja2.Cells(2, 4) = oData.Fields("surname")
  
  Hoja2.Cells(8, 1) = oData.Fields("sales")
  Hoja2.Cells(8, 2) = oData.Fields("orders")
  Hoja2.Cells(8, 3) = oData.Fields("units")
  Hoja2.Cells(8, 4) = oData.Fields("categories")
  Hoja2.Cells(8, 5) = oData.Fields("products")
  Hoja2.Cells(8, 6) = oData.Fields("customers")
  Hoja2.Cells(8, 7) = oData.Fields("employees")
  Hoja2.Cells(8, 8) = oData.Fields("positions")
  Hoja2.Cells(8, 9) = oData.Fields("settings")
  Hoja2.Cells(8, 10) = oData.Fields("database")
  
  xRibbon.InvalidateControl ("sales")
  xRibbon.InvalidateControl ("orders")
  xRibbon.InvalidateControl ("units")
  xRibbon.InvalidateControl ("categories")
  xRibbon.InvalidateControl ("products")
  xRibbon.InvalidateControl ("customers")
  xRibbon.InvalidateControl ("employees")
  xRibbon.InvalidateControl ("positions")
  xRibbon.InvalidateControl ("settings")
  xRibbon.InvalidateControl ("database")
  
  sql = "UPDATE cashiers SET idUser='" & oData.Fields("idUser") & "' WHERE serialNumber='" & GetSerialNumber & "'"
  Set oData = ExecuteQuery(sql)
  
  sql = "SELECT idCashier, cashier FROM cashiers WHERE serialNumber='" & GetSerialNumber & "' AND idState<>3"
  Set oData = ExecuteQuery(sql)
  
  If Not oData.EOF Then
    Hoja2.Cells(5, 1) = oData.Fields("idCashier")
    Hoja2.Cells(5, 2) = oData.Fields("cashier")
  End If
  
  ThisWorkbook.Save
  Unload Me
Else
  MsgBox "Verifique nuevamente su usuario y contraseña"
End If

Set oData = Nothing
End Sub

Private Sub UserForm_Initialize()
FormDesign Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then
  ThisWorkbook.Application.DisplayAlerts = False
  ThisWorkbook.Application.Quit
  Unload Me
End If
End Sub
