Attribute VB_Name = "Procedures"
Option Explicit
Option Private Module

Public Sub CleanControls(pFrm As UserForm, Optional pTag As String)

  Dim ctrl As control

  For Each ctrl In pFrm.Controls
  
    On Error Resume Next
    
    If pTag <> Empty Then
      If Not CBool(InStr(ctrl.Tag, pTag)) Then GoTo continueForLoop
    End If
    
    If TypeOf ctrl Is MSForms.TextBox Then
      ctrl.Text = Empty
    End If
    
    If TypeOf ctrl Is MSForms.ComboBox Then
      ctrl.Style = 0
      ctrl.Text = Empty
      ctrl.Style = 2
    End If

    If TypeOf ctrl Is MSForms.CheckBox Then
      ctrl.Value = False
    End If
    
    If TypeOf ctrl Is MSForms.OptionButton Then
      ctrl.Value = False
    End If

    If TypeOf ctrl Is MSForms.ListBox Then
      ctrl.Clear
    End If
    
continueForLoop:
  Next ctrl

End Sub

Public Sub FillComboBox(pInTheSheet As Worksheet, pFromTheRow As Integer, pInTheColumn As Integer, pComboBox As control)

  Dim lastRowContained As Long
  Dim i As Long

  lastRowContained = GetEmptyCell(pInTheSheet, CLng(pFromTheRow), CLng(pInTheColumn)) - 1

  For i = pFromTheRow To lastRowContained
      pComboBox.AddItem pInTheSheet.Cells(i, pInTheColumn)
  Next i

End Sub

Public Sub FormDesign(pFrm As UserForm)

  Dim ctrl As control

  pFrm.BackColor = vbWhite
  
  For Each ctrl In pFrm.Controls
      
    On Error Resume Next

    ctrl.ForeColor = 4210752
    ctrl.BackColor = vbWhite
    ctrl.BackStyle = fmBackStyleTransparent
    ctrl.SpecialEffect = fmSpecialEffectEtched
    ctrl.TabStop = False
    'ctrl.Font.Name = "Calibri"
    'ctrl.Font.Size = 10
    
    If TypeOf ctrl Is MSForms.CommandButton Then
      ctrl.TakeFocusOnClick = False
    End If
    
    If TypeOf ctrl Is MSForms.Frame Then
      ctrl.SpecialEffect = fmSpecialEffectFlat
      ctrl.ForeColor = &HC0C000
      ctrl.BorderStyle = fmBorderStyleSingle
      ctrl.BorderColor = 12632256
    End If
    
    If TypeOf ctrl Is MSForms.TextBox Then
      ctrl.TabStop = True
      ctrl.SelectionMargin = False
      ctrl.TextAlign = 2
    End If
    
    If TypeOf ctrl Is MSForms.ComboBox Then
      ctrl.TabStop = True
      ctrl.SelectionMargin = False
      ctrl.TextAlign = 2
    End If
    
    If TypeOf ctrl Is MSForms.Label Then
      ctrl.SpecialEffect = fmSpecialEffectFlat
    End If

    If TypeOf ctrl Is MSForms.CheckBox Then
      ctrl.SpecialEffect = fmSpecialEffectFlat
    End If
    
    If TypeOf ctrl Is MSForms.OptionButton Then
      ctrl.SpecialEffect = fmSpecialEffectFlat
    End If

    On Error GoTo 0

  Next ctrl

End Sub
