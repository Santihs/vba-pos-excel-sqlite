VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCustomers 
   Caption         =   "Clientes"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4335
   OleObjectBlob   =   "frmCustomers.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch_Click()
Dim oData As Object

If Not IsNumeric(Me.txtDni.Text) Then Exit Sub

Set oData = ConsultRecords("customers", "dni", Me.txtDni.Text, True)

If Not oData.EOF Then
  Do While Not oData.EOF
    Me.txtIdCustomer.Text = oData.Fields("idCustomer")
    Me.txtName.Text = oData.Fields("name")
    Me.txtSurname.Text = oData.Fields("surname")
    Me.txtPhone.Text = oData.Fields("phone")
    Me.txtEmail.Text = oData.Fields("email")
    Me.txtAddress.Text = oData.Fields("address")
  
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

sql = "INSERT INTO customers (dni, name, surname, phone, email, address) VALUES ('" & Me.txtDni.Text & "','" & Me.txtName.Text & "','" & Me.txtSurname.Text & "','" & Me.txtPhone.Text & "','" & Me.txtEmail.Text & "','" & Me.txtAddress.Text & "')"
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

sql = "UPDATE customers SET dni='" & Me.txtDni.Text & "', name='" & Me.txtName.Text & "', surname='" & Me.txtSurname.Text & "', phone='" & Me.txtPhone.Text & "', email='" & Me.txtEmail.Text & "', address='" & Me.txtAddress.Text & "' WHERE idCustomer=" & Me.txtIdCustomer.Text
Set oData = ExecuteQuery(sql)
Set oData = Nothing

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

sql = "UPDATE customers SET idState='3' WHERE idCustomer=" & Me.txtIdCustomer.Text
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
FormDesign Me
End Sub

