VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDDBB 
   Caption         =   "Configurar base de datos"
   ClientHeight    =   1695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4920
   OleObjectBlob   =   "frmDDBB.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmDDBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGetDB_Click()
Dim fileName As String
Dim xSeparated() As String

fileName = Application.GetOpenFilename("Access Files (*.db), *.db")

If Not Len(Dir(fileName)) = 0 Then
  xSeparated = Split(fileName, "\")
  
  Me.txtDB.Text = xSeparated(UBound(xSeparated))
  Me.txtDB.Tag = fileName
End If
End Sub

Private Sub cmdSave_Click()
Dim oData As Object
Dim sql As String

If Me.txtDB.Text <> Empty Then
  Hoja2.Cells(5, 4) = Me.txtDB.Tag
  
  sql = "SELECT cashier FROM cashiers WHERE serialNumber='" & GetSerialNumber & "' AND idState<>3"
  Set oData = ExecuteQuery(sql)
    
  Unload Me
    
  If Not oData.EOF Then
    frmLogin.Show
  Else
    frmSettings.Show
  End If
Else
  MsgBox "No ha seleccionado ninguna base de datos"
End If
End Sub

Private Sub UserForm_Initialize()
FormDesign Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 And Len(Dir(Hoja2.Cells(5, 4))) = 0 Then
  ThisWorkbook.Application.DisplayAlerts = False
  ThisWorkbook.Application.Quit
  Unload Me
End If
End Sub
