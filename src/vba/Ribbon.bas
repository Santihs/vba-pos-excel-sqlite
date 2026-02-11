Attribute VB_Name = "Ribbon"
Public xRibbon As IRibbonUI

'Callback for customUI.onLoad
Sub LoadRibbon(ribbon As IRibbonUI)
Dim oData As Object
Dim sql As String

Set xRibbon = ribbon

sql = "SELECT idUser FROM cashiers WHERE serialNumber='" & GetSerialNumber & "' AND idState<>3"
Set oData = ExecuteQuery(sql)

If Not oData Is Nothing Then
  If Not oData.EOF Then
    If IsNull(oData.Fields("idUser")) Then
      frmLogin.Show
    End If
  End If
End If

Set oData = Nothing

End Sub

'Callback for sales onAction
Sub Sales(control As IRibbonControl)
frmSales.Show
End Sub

'Callback for sales getVisible
Sub ReturnOfSales(control As IRibbonControl, ByRef returnedVal)
returnedVal = Hoja2.Cells(8, 1)
End Sub

'Callback for orders onAction
Sub Orders(control As IRibbonControl)
frmOrders.Show
End Sub

'Callback for orders getVisible
Sub ReturnOfOrders(control As IRibbonControl, ByRef returnedVal)
returnedVal = Hoja2.Cells(8, 2)
End Sub

'Callback for units onAction
Sub Units(control As IRibbonControl)
frmUnits.Show
End Sub

'Callback for units getVisible
Sub ReturnOfUnits(control As IRibbonControl, ByRef returnedVal)
returnedVal = Hoja2.Cells(8, 3)
End Sub

'Callback for categories onAction
Sub Categories(control As IRibbonControl)
frmCategories.Show
End Sub

'Callback for categories getVisible
Sub ReturnOfCategories(control As IRibbonControl, ByRef returnedVal)
returnedVal = Hoja2.Cells(8, 4)
End Sub

'Callback for products onAction
Sub Products(control As IRibbonControl)
frmProducts.Show
End Sub

'Callback for products getVisible
Sub ReturnOfProducts(control As IRibbonControl, ByRef returnedVal)
returnedVal = Hoja2.Cells(8, 5)
End Sub

'Callback for customers onAction
Sub Customers(control As IRibbonControl)
frmCustomers.Show
End Sub

'Callback for customers getVisible
Sub ReturnOfCustomers(control As IRibbonControl, ByRef returnedVal)
returnedVal = Hoja2.Cells(8, 6)
End Sub

'Callback for employees onAction
Sub Employees(control As IRibbonControl)
frmEmployees.Show
End Sub

'Callback for employees getVisible
Sub ReturnOfEmployees(control As IRibbonControl, ByRef returnedVal)
returnedVal = Hoja2.Cells(8, 7)
End Sub

'Callback for positions onAction
Sub Positions(control As IRibbonControl)
frmPositions.Show
End Sub

'Callback for positions getVisible
Sub ReturnOfPositions(control As IRibbonControl, ByRef returnedVal)
returnedVal = Hoja2.Cells(8, 8)
End Sub

'Callback for settings onAction
Sub Settings(control As IRibbonControl)
frmSettings.Show
End Sub

'Callback for settings getVisible
Sub ReturnOfSettings(control As IRibbonControl, ByRef returnedVal)
returnedVal = Hoja2.Cells(8, 9)
End Sub

'Callback for database onAction
Sub Database(control As IRibbonControl)
frmDDBB.Show
End Sub

'Callback for database getVisible
Sub ReturnOfDatabase(control As IRibbonControl, ByRef returnedVal)
returnedVal = Hoja2.Cells(8, 10)
End Sub

'Callback for reset-password onAction
Sub ResetPassword(control As IRibbonControl)
frmResetPassword.Show
End Sub

'Callback for save onAction
Sub Save(control As IRibbonControl)
ThisWorkbook.Save
End Sub

'Callback for reload onAction
Sub Reload(control As IRibbonControl)
ThisWorkbook.Save
frmLogin.Show
End Sub

'Callback for about onAction
Sub About(control As IRibbonControl)
frmMe.Show
End Sub

'Callback for logout onAction
Sub Logout(control As IRibbonControl)
Dim oData As Object
Dim sql As String

sql = "UPDATE cashiers SET idUser=Null WHERE serialNumber='" & GetSerialNumber & "'"
Set oData = ExecuteQuery(sql)
Set oData = Nothing

ThisWorkbook.Save
ThisWorkbook.Application.DisplayAlerts = False
ThisWorkbook.Application.Quit
End Sub
