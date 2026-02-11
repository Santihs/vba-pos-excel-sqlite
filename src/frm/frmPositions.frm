VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPositions 
   Caption         =   "Cargos"
   ClientHeight    =   9015.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4935
   OleObjectBlob   =   "frmPositions.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmPositions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCreate_Click()
Dim oData As Object
Dim sql As String

If ValidateFields(Me) = False Then Exit Sub
If MsgBox("¿Desea crear el registro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

sql = "INSERT INTO positions (position, sales, orders, units, categories, products, customers, employees, positions, settings, database) VALUES ('" & Me.txtPosition.Text & "','" & CByte(Me.checkSales.Value) & "','" & CByte(Me.checkOrders.Value) & "','" & CByte(Me.checkUnits.Value) & "','" & CByte(Me.checkCategories.Value) & "','" & CByte(Me.checkProducts.Value) & "','" & CByte(Me.checkCustomers.Value) & "','" & CByte(Me.checkEmployees.Value) & "','" & CByte(Me.checkPositions.Value) & "','" & CByte(Me.checkSettings.Value) & "','" & CByte(Me.checkDatabase.Value) & "')"
Set oData = ExecuteQuery(sql)
Set oData = Nothing

CleanControls Me
Me.txtPosition.SetFocus

MsgBox "El registro se ha creado exitosamente"
UserForm_Initialize
End Sub

Private Sub cmdDelete_Click()
Dim oData As Object
Dim sql As String

If MsgBox("¿Desea eliminar el registro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

sql = "UPDATE positions SET idState='3' WHERE idPosition=" & Me.txtIdPosition.Text
Set oData = ExecuteQuery(sql)
Set oData = Nothing

Me.cmdCreate.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False

CleanControls Me
Me.txtPosition.SetFocus

MsgBox "El registro se ha eliminado exitosamente"
UserForm_Initialize
End Sub

Private Sub cmdUpdate_Click()
Dim oData As Object
Dim sql As String

If ValidateFields(Me) = False Then Exit Sub
If MsgBox("¿Desea actualizar el registro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

sql = "UPDATE positions SET position='" & Me.txtPosition.Text & "', sales='" & CByte(Me.checkSales.Value) & "', orders='" & CByte(Me.checkOrders.Value) & "', units='" & CByte(Me.checkUnits.Value) & "', categories='" & CByte(Me.checkCategories.Value) & "', products='" & CByte(Me.checkProducts.Value) & "', customers='" & CByte(Me.checkCustomers.Value) & "', employees='" & CByte(Me.checkEmployees.Value) & "', positions='" & CByte(Me.checkPositions.Value) & "', settings='" & CByte(Me.checkSettings.Value) & "', database='" & CByte(Me.checkDatabase.Value) & "' WHERE idPosition=" & Me.txtIdPosition.Text
Set oData = ExecuteQuery(sql)
Set oData = Nothing

Me.cmdCreate.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False

CleanControls Me
Me.txtPosition.SetFocus

MsgBox "El registro se ha actualizado exitosamente"
UserForm_Initialize
End Sub

Private Sub listPosition_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim oData As Object
Dim sql As String

If Me.listPosition.ListIndex <> -1 Then

  Me.txtPosition.Text = Me.listPosition.List(Me.listPosition.ListIndex, 0)
  Me.txtIdPosition.Text = Me.listPosition.List(Me.listPosition.ListIndex, 1)

  sql = "SELECT sales, orders, units, categories, products, customers, employees, positions, settings, database FROM positions WHERE idPosition=" & Me.txtIdPosition.Text & " AND positions.idState<>3"
  Set oData = ExecuteQuery(sql)

  If Not oData.EOF Then
    Me.checkSales.Value = oData.Fields("sales")
    Me.checkOrders.Value = oData.Fields("orders")
    Me.checkUnits.Value = oData.Fields("units")
    Me.checkCategories.Value = oData.Fields("categories")
    Me.checkProducts.Value = oData.Fields("products")
    Me.checkCustomers.Value = oData.Fields("customers")
    Me.checkEmployees.Value = oData.Fields("employees")
    Me.checkPositions.Value = oData.Fields("positions")
    Me.checkSettings.Value = oData.Fields("settings")
    Me.checkDatabase.Value = oData.Fields("database")
  End If
  
  Me.cmdCreate.Enabled = False
  Me.cmdUpdate.Enabled = True
  Me.cmdDelete.Enabled = True
  
  Set oData = Nothing
End If
End Sub

Private Sub UserForm_Initialize()
Dim oData As Object

FormDesign Me

Me.listPosition.Clear

Set oData = ConsultTable("positions")

If Not oData.EOF Then
  Do While Not oData.EOF
    Me.listPosition.AddItem oData.Fields("position")
    Me.listPosition.List(Me.listPosition.ListCount - 1, 1) = oData.Fields("idPosition")
    oData.movenext
  Loop
End If

Set oData = Nothing
End Sub


