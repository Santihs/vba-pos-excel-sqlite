VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearch 
   Caption         =   "Buscar productos"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10695
   OleObjectBlob   =   "frmSearch.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub listProduct_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim frm As UserForm

If Me.listProduct.ListIndex <> -1 Then
  If frmSales.Visible = True Then
    Set frm = frmSales
  ElseIf frmOrders.Visible = True Then
    Set frm = frmOrders
  End If
  
  frm.txtCode.Text = Me.listProduct.List(Me.listProduct.ListIndex, 0)
  frm.txtQuantity.SetFocus
  Unload Me
End If
End Sub

Private Sub txtSearch_Change()
Dim oData As Object
Dim sql As String

Me.listProduct.Clear
        
sql = "SELECT barcode, product, cost , price, category FROM products " & _
     "INNER JOIN categories ON products.idCategory = categories.idCategory " & _
     "WHERE barcode LIKE '%" & Me.txtSearch.Text & "%' OR product LIKE '%" & Me.txtSearch.Text & "%' OR category LIKE '%" & Me.txtSearch.Text & "%' AND products.idState<>3  AND categories.idState<>3"
Set oData = ExecuteQuery(sql)

If Not oData.EOF Then
  Do While Not oData.EOF
    Me.listProduct.AddItem oData.Fields("barcode")
     Me.listProduct.List(Me.listProduct.ListCount - 1, 1) = oData.Fields("product")
     Me.listProduct.List(Me.listProduct.ListCount - 1, 2) = oData.Fields("category")
     Me.listProduct.List(Me.listProduct.ListCount - 1, 3) = CDbl(oData.Fields("cost"))
     Me.listProduct.List(Me.listProduct.ListCount - 1, 4) = CDbl(oData.Fields("price"))

    oData.movenext
  Loop
End If

Set oData = Nothing
End Sub

Private Sub UserForm_Initialize()
FormDesign Me
End Sub

