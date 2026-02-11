VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSales 
   Caption         =   "Ventas"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11175
   OleObjectBlob   =   "frmSales.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSale_Click()
Dim oData As Object
Dim sql As String
Dim i As Integer
Dim xIdSale As String
Dim xUuid As String
Dim customerDni As String

If Me.listProduct.ListCount <= 0 Then Exit Sub
If Me.txtTurned.Text < 0 Then Exit Sub
If ValidateFields(Me) = False Then Exit Sub

xUuid = Uuid

customerDni = InputBox("DNI del cliente")
customerDni = IIf(customerDni = Empty, 0, customerDni)

sql = "SELECT idCustomer FROM customers WHERE dni='" & customerDni & "' AND idState<>3"
Set oData = ExecuteQuery(sql)

sql = "INSERT INTO sales (idCashier, idCustomer, idEmployee, uuid, total, date) VALUES ('" & Hoja2.Cells(5, 1) & "','" & IIf(Not oData.EOF, oData.Fields("idCustomer"), 1) & "','" & Hoja2.Cells(2, 1) & "','" & xUuid & "','" & Me.txtGrandTotal.Text & "','" & CLng(Date) & "')"

Set oData = ExecuteQuery(sql)

For i = 1 To Me.listProduct.ListCount
  With Me.listProduct
    sql = "SELECT stock FROM products WHERE idProduct=" & .List(i - 1, 5) & " AND idState<>3"
    Set oData = ExecuteQuery(sql)
    
    sql = "UPDATE products SET stock='" & CDbl(oData.Fields("stock")) - CDbl(.List(i - 1, 1)) & "' WHERE idProduct=" & .List(i - 1, 5)
    Set oData = ExecuteQuery(sql)
  End With
Next i

sql = "SELECT idSale FROM sales WHERE uuid='" & xUuid & "' AND idState<>3"
Set oData = ExecuteQuery(sql)
xIdSale = oData.Fields("idSale")

For i = 1 To Me.listProduct.ListCount
  With Me.listProduct
    sql = "INSERT INTO saleDetails (idSale, idProduct, quantity, price) VALUES ('" & xIdSale & "','" & .List(i - 1, 5) & "','" & .List(i - 1, 1) & "','" & .List(i - 1, 3) & "')"
    Set oData = ExecuteQuery(sql)
  End With
Next i
Set oData = Nothing

MsgBox "La venta se ha registrado exitosamente"
GenerateTicket xIdSale, customerDni

CleanControls Me, "clean"
Me.listProduct.Clear
Me.txtTotal.Text = 0
Me.txtGrandTotal.Text = 0
Me.txtTurned.Text = 0
Me.txtCode.SetFocus

End Sub

Private Sub cmdSearch_Click()
  frmSearch.Show
End Sub

Private Sub listProduct_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Me.listProduct.ListIndex <> -1 Then
  Me.txtGrandTotal.Text = CDbl(Me.txtGrandTotal.Text) - CDbl(Me.listProduct.List(Me.listProduct.ListIndex, 4))
  Me.listProduct.RemoveItem (Me.listProduct.ListIndex)
  Me.listProduct.ListIndex = -1
  Me.txtCode.SetFocus
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
Me.listProduct.List(Me.listProduct.ListCount - 1, 3) = Me.txtPrice.Text
Me.listProduct.List(Me.listProduct.ListCount - 1, 4) = Me.txtTotal.Text
Me.listProduct.List(Me.listProduct.ListCount - 1, 5) = Me.txtIdProduct.Text

Me.txtGrandTotal.Text = CDbl(Me.txtGrandTotal.Text) + CDbl(Me.txtTotal.Text)

CleanControls Me, "clean"
Me.txtTotal.Text = 0
End Sub

Private Sub txtQuantity_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  Me.txtCode.SetFocus
  KeyCode = vbKeyCancel
End If
End Sub

Private Sub txtCode_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If Me.txtDescription.Text <> Empty Then
  Me.txtQuantity.Text = 1
End If
End Sub

Private Sub txtCode_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyF2 Then
  cmdSearch_Click
End If
End Sub

Private Sub txtEffective_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = vbKeyF3 Then
  cmdSale_Click
End If
End Sub

Private Sub txtEffective_Change()
If Not IsNumeric(Me.txtEffective.Text) Then Exit Sub

If Me.txtEffective.Text <> Empty Then
  Me.txtTurned.Text = Round(CDbl(CDbl(Me.txtEffective.Text) - CDbl(Me.txtGrandTotal.Text)), 2)
Else
  Me.txtTurned.Text = 0
End If
End Sub

Private Sub txtCode_Change()

Dim oData As Object
Dim sql As String

sql = "SELECT idProduct, product, cost FROM products WHERE barcode='" & Me.txtCode.Text & "' AND idState<>3"
Set oData = ExecuteQuery(sql)

If Not oData.EOF Then
  Me.txtDescription.Text = oData.Fields("product")
  Me.txtPrice.Text = oData.Fields("cost")
  Me.txtIdProduct.Text = oData.Fields("idProduct")
End If

Set oData = Nothing

End Sub

Private Sub txtQuantity_Change()
If Not IsNumeric(Me.txtQuantity.Text) Then Exit Sub

If Me.txtQuantity.Text <> Empty And Me.txtPrice.Text <> Empty Then
  Me.txtTotal.Text = Me.txtPrice.Text * Me.txtQuantity.Text
Else
  Me.txtTotal.Text = Empty
End If
End Sub

Private Sub UserForm_Initialize()

FormDesign Me

Me.txtCashier.Text = Hoja2.Cells(5, 2)
Me.txtEmployee.Text = Hoja2.Cells(2, 3) & " " & Hoja2.Cells(2, 4)
Me.txtDate.Text = Date

End Sub

Private Sub GenerateTicket(idSale, customerDni)
Dim oData As Object
Dim sql As String
Dim i As Long
Dim j As Long

i = GetEmptyCell(Hoja3, 18, 2) - 1

If i >= 18 Then
  For j = 18 To i
       Hoja3.Rows(18).Delete Shift:=xlUp
  Next j
End If

For j = 2 To Me.listProduct.ListCount
  Hoja3.Rows(18).Insert Shift:=xlDown
  Hoja3.Range("C18:E18").Merge
Next j

Hoja3.Cells(9, 4) = idSale
Hoja3.Cells(10, 4) = Now
Hoja3.Cells(11, 4) = Me.txtEmployee.Text

sql = "SELECT dni, name, surname FROM customers WHERE dni='" & customerDni & "' AND idState<>3"
Set oData = ExecuteQuery(sql)

If Not oData.EOF Then
  Hoja3.Cells(13, 4) = oData.Fields("name") & " " & oData.Fields("surname")
  Hoja3.Cells(14, 4) = oData.Fields("dni")
Else
  Hoja3.Cells(13, 4) = "Cliente Genérico"
  Hoja3.Cells(14, 4) = "0"
End If
Set oData = Nothing

For j = 1 To Me.listProduct.ListCount
  Hoja3.Cells(j + 16, 2) = Me.listProduct.List(j - 1, 1)
  Hoja3.Cells(j + 16, 3) = Me.listProduct.List(j - 1, 2)
  Hoja3.Cells(j + 16, 6) = Me.listProduct.List(j - 1, 4)
Next j

Hoja3.Cells(Me.listProduct.ListCount + 19, 6) = Me.txtGrandTotal.Text
Hoja3.Cells(Me.listProduct.ListCount + 20, 6) = "0"
Hoja3.Cells(Me.listProduct.ListCount + 21, 6) = Me.txtGrandTotal.Text
Hoja3.Cells(Me.listProduct.ListCount + 22, 6) = Me.txtEffective.Text
Hoja3.Cells(Me.listProduct.ListCount + 23, 6) = Me.txtTurned.Text

ThisWorkbook.Save

Hoja3.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False

End Sub
