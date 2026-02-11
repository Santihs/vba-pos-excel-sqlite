VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProducts 
   Caption         =   "Productos"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4815
   OleObjectBlob   =   "frmProducts.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch_Click()
Dim oData As Object

If Me.txtCode.Text = Empty Then Exit Sub

Set oData = ConsultRecords("products", "barcode", "'" & Me.txtCode.Text & "'", True)

If Not oData.EOF Then
  Do While Not oData.EOF
    Me.txtIdProduct.Text = oData.Fields("idProduct")
    Me.txtCode.Text = oData.Fields("barcode")
    Me.txtDescription.Text = oData.Fields("product")
    Me.txtCost.Text = oData.Fields("cost")
    Me.txtPrice.Text = oData.Fields("price")
    Me.txtCategory.Text = oData.Fields("idCategory")
    Me.txtUnit.Text = oData.Fields("idUnit")
  
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

sql = "INSERT INTO products (barcode, product, cost, price, idCategory, idUnit) " & _
"VALUES ('" & Me.txtCode.Text & "','" & Me.txtDescription.Text & "','" & Me.txtCost.Text & "','" & Me.txtPrice.Text & "','" & Me.txtCategory.Text & "','" & Me.txtUnit.Text & "')"
Set oData = ExecuteQuery(sql)
Set oData = Nothing

CleanControls Me
Me.txtCode.SetFocus

MsgBox "El registro se ha creado exitosamente"
End Sub

Private Sub cmdUpdate_Click()
Dim oData As Object
Dim sql As String

If ValidateFields(Me) = False Then Exit Sub
If MsgBox("¿Desea actualizar el registro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

sql = "UPDATE products SET barcode='" & Me.txtCode.Text & "', product='" & Me.txtDescription.Text & "', cost='" & Me.txtCost.Text & "', price='" & Me.txtPrice.Text & "', idCategory='" & Me.txtCategory.Text & "', idUnit='" & Me.txtUnit.Text & "' WHERE idProduct=" & Me.txtIdProduct.Text

Set oData = ExecuteQuery(sql)
Set oData = Nothing

Me.cmdCreate.Enabled = True
Me.cmdSearch.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False

CleanControls Me
Me.txtCode.SetFocus

MsgBox "El registro se ha actualizado exitosamente"
End Sub

Private Sub cmdDelete_Click()
Dim oData As Object
Dim sql As String

If MsgBox("¿Desea eliminar el registro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

sql = "UPDATE products SET idState='3' WHERE idProduct=" & Me.txtIdProduct.Text
Set oData = ExecuteQuery(sql)
Set oData = Nothing

Me.cmdCreate.Enabled = True
Me.cmdSearch.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False

CleanControls Me
Me.txtCode.SetFocus

MsgBox "El registro se ha eliminado exitosamente"
End Sub

Private Sub UserForm_Initialize()
Dim oData As Object
Dim sql As String

FormDesign Me

sql = "SELECT idCategory, category FROM categories WHERE idState<>3"
Set oData = ExecuteQuery(sql)

Do While Not oData.EOF
  Me.txtCategory.AddItem
  Me.txtCategory.List(Me.txtCategory.ListCount - 1, 0) = oData.Fields("idCategory")
  Me.txtCategory.List(Me.txtCategory.ListCount - 1, 1) = oData.Fields("category")
    
  oData.movenext
Loop

sql = "SELECT idUnit, unit FROM units WHERE idState<>3"
Set oData = ExecuteQuery(sql)

Do While Not oData.EOF
  Me.txtUnit.AddItem
  Me.txtUnit.List(Me.txtUnit.ListCount - 1, 0) = oData.Fields("idUnit")
  Me.txtUnit.List(Me.txtUnit.ListCount - 1, 1) = oData.Fields("unit")
    
  oData.movenext
Loop

Set oData = Nothing

End Sub
