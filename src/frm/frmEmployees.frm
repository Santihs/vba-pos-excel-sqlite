VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEmployees 
   Caption         =   "Empleados"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4335
   OleObjectBlob   =   "frmEmployees.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch_Click()
Dim oData As Object

If Not IsNumeric(Me.txtDni.Text) Then Exit Sub

Set oData = ConsultRecords("employees", "dni", Me.txtDni.Text, True)

If Not oData.EOF Then
  Do While Not oData.EOF
    Me.txtIdEmployee.Text = oData.Fields("idEmployee")
    Me.txtName.Text = oData.Fields("name")
    Me.txtSurname.Text = oData.Fields("surname")
    Me.txtPhone.Text = oData.Fields("phone")
    Me.txtEmail.Text = oData.Fields("email")
    Me.txtAddress.Text = oData.Fields("address")
    Me.txtPosition.Text = oData.Fields("idPosition")
  
    oData.movenext
  Loop
  
  Me.cmdCreate.Enabled = False
  Me.cmdSearch.Enabled = False
  Me.cmdUpdate.Enabled = True
  Me.cmdDelete.Enabled = True
Else
  MsgBox "No se encontraron registros"
End If

Set oData = Nothing
End Sub

Private Sub cmdCreate_Click()
Dim oData As Object
Dim sql As String

If ValidateFields(Me) = False Then Exit Sub
If MsgBox("¿Desea crear el registro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

sql = "INSERT INTO employees (dni, name, surname, phone, email, address, idPosition) VALUES ('" & Me.txtDni.Text & "','" & Me.txtName.Text & "','" & Me.txtSurname.Text & "','" & Me.txtPhone.Text & "','" & Me.txtEmail.Text & "','" & Me.txtAddress.Text & "','" & Me.txtPosition.Text & "')"
Set oData = ExecuteQuery(sql)

sql = "SELECT idEmployee FROM employees WHERE dni='" & Me.txtDni.Text & "' AND idState<>3"
Set oData = ExecuteQuery(sql)

sql = "INSERT INTO users (idEmployee, secretKey) VALUES ('" & oData.Fields("idEmployee") & "', '" & SHA256(Me.txtDni.Text) & "')"
Set oData = ExecuteQuery(sql)
Set oData = Nothing

CleanControls Me
Me.txtDni.SetFocus

MsgBox "El registro se ha creado exitosamente"
End Sub

Private Sub cmdUpdate_Click()
Dim oData As Object
Dim sql As String

If ValidateFields(Me) = False Then Exit Sub
If MsgBox("¿Desea actualizar el registro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

sql = "UPDATE employees SET dni='" & Me.txtDni.Text & "', name='" & Me.txtName.Text & "', surname='" & Me.txtSurname.Text & "', phone='" & Me.txtPhone.Text & "', email='" & Me.txtEmail.Text & "', address='" & Me.txtAddress.Text & "', idPosition='" & Me.txtPosition.Text & "' WHERE idEmployee=" & Me.txtIdEmployee.Text
Set oData = ExecuteQuery(sql)
Set oData = Nothing

Hoja2.Cells(2, 2) = Me.txtDni.Text
Hoja2.Cells(2, 3) = Me.txtName.Text
Hoja2.Cells(2, 4) = Me.txtSurname.Text

Me.cmdCreate.Enabled = True
Me.cmdSearch.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False

CleanControls Me
Me.txtDni.SetFocus

MsgBox "El registro se ha actualizado exitosamente"
End Sub

Private Sub cmdDelete_Click()
Dim oData As Object
Dim sql As String

If MsgBox("¿Desea eliminar el registro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

sql = "UPDATE employees SET idState='3' WHERE idEmployee=" & Me.txtIdEmployee.Text
Set oData = ExecuteQuery(sql)
Set oData = Nothing

Me.cmdCreate.Enabled = True
Me.cmdSearch.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False

CleanControls Me
Me.txtDni.SetFocus

MsgBox "El registro se ha eliminado exitosamente"
End Sub

Private Sub UserForm_Initialize()
Dim oData As Object
Dim sql As String

FormDesign Me

sql = "SELECT idPosition, position FROM positions WHERE idState<>3"
Set oData = ExecuteQuery(sql)

Do While Not oData.EOF
  Me.txtPosition.AddItem
  Me.txtPosition.List(Me.txtPosition.ListCount - 1, 0) = oData.Fields("idPosition")
  Me.txtPosition.List(Me.txtPosition.ListCount - 1, 1) = oData.Fields("position")
    
  oData.movenext
Loop

oData.Close
Set oData = Nothing
End Sub
