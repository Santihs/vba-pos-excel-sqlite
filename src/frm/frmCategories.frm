VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCategories 
   Caption         =   "Categorias"
   ClientHeight    =   8535.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7335
   OleObjectBlob   =   "frmCategories.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCategories"
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

sql = "INSERT INTO categories (category, description) " & _
"VALUES ('" & Me.txtCategory.Text & "','" & Me.txtDescription.Text & "')"
Set oData = ExecuteQuery(sql)
Set oData = Nothing

CleanControls Me
Me.txtCategory.SetFocus

MsgBox "El registro se ha creado exitosamente"
UserForm_Initialize
End Sub

Private Sub cmdDelete_Click()
Dim oData As Object
Dim sql As String

If MsgBox("¿Desea eliminar el registro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

sql = "UPDATE categories SET idState='3' WHERE idCategory=" & Me.txtIdCategory.Text
Set oData = ExecuteQuery(sql)
Set oData = Nothing

Me.cmdCreate.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False

CleanControls Me
Me.txtCategory.SetFocus

MsgBox "El registro se ha eliminado exitosamente"
UserForm_Initialize
End Sub

Private Sub cmdUpdate_Click()
Dim oData As Object
Dim sql As String

If ValidateFields(Me) = False Then Exit Sub
If MsgBox("¿Desea actualizar el registro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

sql = "UPDATE categories SET category='" & Me.txtCategory.Text & "', description='" & Me.txtDescription.Text & "' WHERE idCategory=" & Me.txtIdCategory.Text
Set oData = ExecuteQuery(sql)
Set oData = Nothing

Me.cmdCreate.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False

CleanControls Me
Me.txtCategory.SetFocus

MsgBox "El registro se ha actualizado exitosamente"
UserForm_Initialize
End Sub

Private Sub listCategory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Me.listCategory.ListIndex <> -1 Then
  Me.txtIdCategory.Text = Me.listCategory.List(Me.listCategory.ListIndex, 2)
  Me.txtCategory.Text = Me.listCategory.List(Me.listCategory.ListIndex, 0)
  Me.txtDescription.Text = Me.listCategory.List(Me.listCategory.ListIndex, 1)
  
  Me.cmdCreate.Enabled = False
  Me.cmdUpdate.Enabled = True
  Me.cmdDelete.Enabled = True
End If
End Sub

Private Sub UserForm_Initialize()
Dim oData As Object

FormDesign Me

Me.listCategory.Clear

Set oData = ConsultTable("categories")

If Not oData.EOF Then
  Do While Not oData.EOF
    Me.listCategory.AddItem oData.Fields("category")
    Me.listCategory.List(Me.listCategory.ListCount - 1, 1) = oData.Fields("description")
    Me.listCategory.List(Me.listCategory.ListCount - 1, 2) = oData.Fields("idCategory")

    oData.movenext
  Loop
End If

Set oData = Nothing
End Sub
