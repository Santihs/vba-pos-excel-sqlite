VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrders 
   Caption         =   "Ordenes"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11175
   OleObjectBlob   =   "frmOrders.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "frmOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch_Click()
     frmSearch.Show
End Sub

Private Sub cmdOrder_Click()
Dim oData As Object
Dim sql As String
Dim i As Integer
Dim xIdOrder As String
Dim xUuid As String

If Me.listProduct.ListCount <= 0 Then Exit Sub

xUuid = Uuid

sql = "INSERT INTO orders (uuid, total, date) " & _
"VALUES ('" & xUuid & "','" & Me.txtGrandTotal.Text & "','" & CLng(Date) & "')"
Set oData = ExecuteQuery(sql)

For i = 1 To Me.listProduct.ListCount
  With Me.listProduct
    sql = "SELECT stock FROM products WHERE idProduct=" & .List(i - 1, 5) & " AND idState<>3"
    Set oData = ExecuteQuery(sql)
    
    sql = "UPDATE products SET stock='" & CDbl(oData.Fields("stock")) + CDbl(.List(i - 1, 1)) & "', cost='" & .List(i - 1, 3) & "' WHERE idProduct=" & .List(i - 1, 5)
    Set oData = ExecuteQuery(sql)
  End With
Next i

sql = "SELECT idOrder FROM orders WHERE uuid='" & xUuid & "' AND idState<>3"
Set oData = ExecuteQuery(sql)
xIdOrder = oData.Fields("idOrder")

For i = 1 To Me.listProduct.ListCount
  With Me.listProduct
    sql = "INSERT INTO orderDetails (idOrder, idProduct, quantity, cost) VALUES ('" & xIdOrder & "','" & .List(i - 1, 5) & "','" & .List(i - 1, 1) & "','" & .List(i - 1, 3) & "')"
    Set oData = ExecuteQuery(sql)
  End With
Next i
Set oData = Nothing

CleanControls Me
Me.listProduct.Clear
Me.txtGrandTotal.Text = 0
Me.txtTotal.Text = 0
Me.txtCode.SetFocus

MsgBox "El pedido se ha registrado exitosamente"
End Sub

Private Sub listProduct_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Me.listProduct.ListIndex <> -1 Then
  Me.txtGrandTotal.Text = CDbl(Me.txtGrandTotal.Text) - CDbl(Me.listProduct.List(Me.listProduct.ListIndex, 4))
  Me.listProduct.RemoveItem (Me.listProduct.ListIndex)
  Me.listProduct.ListIndex = -1
  Me.txtCode.SetFocus
End If
End Sub

Private Sub txtCode_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Me.txtDescription.Text <> Empty Then
  Me.txtQuantity.Text = 1
End If
End Sub

Private Sub txtQuantity_Change()
If Not IsNumeric(Me.txtQuantity.Text) Then Exit Sub

If Me.txtQuantity.Text <> Empty And Me.txtCost.Text <> Empty Then
  Me.txtTotal.Text = Me.txtCost.Text * Me.txtQuantity.Text
Else
  Me.txtTotal.Text = Empty
End If
End Sub

Private Sub txtQuantity_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Dim i As Integer

If Me.txtQuantity.Text = Empty Or Me.txtDescription.Text = Empty Then Exit Sub

For i = 1 To Me.listProduct.ListCount
  If Me.listProduct.List(i - 1, 0) = Me.txtCode.Text Then
    Me.txtQuantity.Text = CDbl(Me.listProduct.List(i - 1, 1)) + CDbl(Me.txtQuantity.Text)
    Me.txtGrandTotal.Text = CDbl(Me.txtGrandTotal.Text) - CDbl(Me.listProduct.List(i - 1, 4))
    Me.listProduct.RemoveItem (i - 1)
    Me.listProduct.ListIndex = -1
    Exit For
  End If
Next i

Me.listProduct.AddItem Me.txtCode.Text
Me.listProduct.List(Me.listProduct.ListCount - 1, 1) = Me.txtQuantity.Text
Me.listProduct.List(Me.listProduct.ListCount - 1, 2) = Me.txtDescription.Text
Me.listProduct.List(Me.listProduct.ListCount - 1, 3) = Me.txtCost.Text
Me.listProduct.List(Me.listProduct.ListCount - 1, 4) = Me.txtTotal.Text
Me.listProduct.List(Me.listProduct.ListCount - 1, 5) = Me.txtIdProduct.Text

Me.txtGrandTotal.Text = CDbl(Me.txtGrandTotal.Text) + CDbl(Me.txtTotal.Text)
CleanControls Me, "clean"
Me.txtTotal.Text = 0
End Sub

Private Sub txtCode_Change()
Dim oData As Object
Dim sql As String

sql = "SELECT idProduct, product, cost FROM products WHERE barcode='" & Me.txtCode.Text & "' AND idState<>3"
Set oData = ExecuteQuery(sql)

If Not oData.EOF Then
  Me.txtDescription.Text = oData.Fields("product")
  Me.txtCost.Text = oData.Fields("cost")
  Me.txtIdProduct.Text = oData.Fields("idProduct")
End If

Set oData = Nothing
End Sub

Private Sub txtCode_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyF2 Then
     cmdSearch_Click
ElseIf KeyCode = vbKeyF3 Then
     cmdOrder_Click
End If
End Sub

Private Sub txtQuantity_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  Me.txtCode.SetFocus
  KeyCode = vbKeyCancel
End If
End Sub

Private Sub UserForm_Initialize()
FormDesign Me
End Sub
