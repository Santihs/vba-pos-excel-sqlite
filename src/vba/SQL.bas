Attribute VB_Name = "SQL"
Option Explicit
Option Private Module

Public Function ConsultTable(pInTheTable As String) As Object
    
  Dim sql As String
  
  sql = "SELECT * FROM " & pInTheTable & " WHERE idState<>3"

  Set ConsultTable = ExecuteQuery(sql)
  
End Function

Public Function ConsultRecords(pInTheTable As String, pInTheField As String, pWantedValue As String, pIsExact As Boolean) As Object

  Dim sql As String
  Dim operador As String
  
  If pIsExact = True Then
    operador = " = "
  Else
    operador = " like "
    pWantedValue = "'%" & pWantedValue & "%'"
  End If
  
  sql = "SELECT * FROM " & pInTheTable & " WHERE " & pInTheField & operador & pWantedValue & " AND idState<>3"
  
  Set ConsultRecords = ExecuteQuery(sql)
            
End Function

Public Function ExecuteQuery(sql As String) As Object

  On Error GoTo handleError

  Dim cn As Object
  Dim ddbb As String

'  If Not Len(Dir(ThisWorkbook.Path & "\DDBB.accdb")) = 0 Then 'Access
  If Not Len(Dir(Hoja2.Cells(5, 4))) = 0 Then 'SQLite
    Set cn = CreateObject("ADODB.Connection")
  
'    ddbb = "Provider=Microsoft.ACE.OLEDB.12.0; Data source=" & ThisWorkbook.Path & "\DDBB.accdb" 'Access
    ddbb = "Driver=SQLite3 ODBC Driver; Database=" & Hoja2.Cells(5, 4) 'SQLite
            
    cn.Open ddbb
  
    Set ExecuteQuery = cn.Execute(sql)
    Set cn = Nothing
  Else
    frmDDBB.Show
  End If
  
handleError:
  If Err.Number <> 0 Then
    MsgBox Err.Description
    End
  End If
            
End Function

Public Function DeleteRecords(pInTheTable As String, pInTheField As String, pValueToDeleted As String) As Object

  Dim sql As String

  sql = "DELETE FROM " & pInTheTable & " WHERE " & pInTheField & " = " & pValueToDeleted

  Set DeleteRecords = ExecuteQuery(sql)

End Function
