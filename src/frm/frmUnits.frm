VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUnits 
   Caption         =   "Unidades"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4935
   OleObjectBlob   =   "frmUnits.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmUnits"
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

sql = "INSERT INTO units (unit) VALUES ('" & Me.txtUnit.Text & "')"
Set oData = ExecuteQuery(sql)
Set oData = Nothing

CleanControls Me
Me.txtUnit.SetFocus

MsgBox "El registro se ha creado exitosamente"
UserForm_Initialize
End Sub

Private Sub cmdDelete_Click()
Dim oData As Object
Dim sql As String

If MsgBox("¿Desea eliminar el registro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

sql = "UPDATE units SET idState='3' WHERE idUnit=" & Me.txtIdUnit.Text
Set oData = ExecuteQuery(sql)
Set oData = Nothing

Me.cmdCreate.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False

CleanControls Me
Me.txtUnit.SetFocus

MsgBox "El registro se ha eliminado exitosamente"
UserForm_Initialize
End Sub

Private Sub cmdUpdate_Click()
Dim oData As Object
Dim sql As String

If ValidateFields(Me) = False Then Exit Sub
If MsgBox("¿Desea actualizar el registro?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

sql = "UPDATE units SET unit='" & Me.txtUnit.Text & "' WHERE idUnit=" & Me.txtIdUnit.Text
Set oData = ExecuteQuery(sql)
Set oData = Nothing

Me.cmdCreate.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False

CleanControls Me
Me.txtUnit.SetFocus

MsgBox "El registro se ha actualizado exitosamente"
UserForm_Initialize
End Sub

Private Sub listunit_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim oData As Object
Dim sql As String

If Me.listUnit.ListIndex <> -1 Then

  Me.txtUnit.Text = Me.listUnit.List(Me.listUnit.ListIndex, 0)
  Me.txtIdUnit.Text = Me.listUnit.List(Me.listUnit.ListIndex, 1)
  
  Me.cmdCreate.Enabled = False
  Me.cmdUpdate.Enabled = True
  Me.cmdDelete.Enabled = True
  
End If
End Sub

Private Sub UserForm_Initialize()
Dim oData As Object

FormDesign Me

Me.listUnit.Clear

Set oData = ConsultTable("units")

If Not oData.EOF Then
  Do While Not oData.EOF
    Me.listUnit.AddItem oData.Fields("unit")
    Me.listUnit.List(Me.listUnit.ListCount - 1, 1) = oData.Fields("idUnit")
    oData.movenext
  Loop
End If

Set oData = Nothing
End Sub



