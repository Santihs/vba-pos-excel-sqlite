VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    ' Suppress "Update Links" warning for Premium Add-in
    Application.AskToUpdateLinks = False

    ActiveWindow.WindowState = xlMaximized
    frmSplash.Show
    
    ' NEW: Load premium add-in (Week 0)
    LoadPremiumAddIn
End Sub

' ============================================================================
' PREMIUM INTEGRATION (Week 0)
' ============================================================================
Private Sub LoadPremiumAddIn()
    ' Automatically load premium add-in if available
    ' This function checks for premium XLA and loads it if found

    On Error Resume Next ' Continue if premium not found

    Dim premiumPath As String
    Dim alreadyLoaded As Boolean
    Dim startTime As Double

    startTime = Timer

    ' Define path to premium add-in
    ' NOTE: Adjust this path if your premium location is different
    ' Determine dynamic path to premium add-in
    premiumPath = GetPremiumPath_Private()

    Debug.Print String(60, "=")
    Debug.Print "LoadPremiumAddIn: Starting at " & Now
    Debug.Print "  Resolved Path: " & premiumPath

    ' Check if premium add-in file exists
    If premiumPath = "" Or Dir(premiumPath) = "" Then
        Debug.Print "  Result: Premium add-in NOT FOUND at resolved path"
        Debug.Print "  Action: Continuing without premium features"
        Debug.Print String(60, "=")
        Exit Sub
    End If

    Debug.Print "  Result: Premium add-in file found"

    ' Check if workbook is already open
    On Error Resume Next
    Dim wb As Workbook
    Set wb = Workbooks("PremiumAddon.xla")
    On Error GoTo 0
    
    If Not wb Is Nothing Then
        alreadyLoaded = True
        Debug.Print "  Status: Already open"
    End If

    ' If not loaded, open it
    If Not alreadyLoaded Then
        Debug.Print "  Status: Not loaded yet"
        Set wb = Workbooks.Open(premiumPath)
        Debug.Print "  Action: Opened workbook directly"
        
        ' HIDE THE ADD-IN WINDOW
        On Error Resume Next
        wb.IsAddin = True ' Ensures it behaves like an add-in
        wb.Windows(1).Visible = False
        On Error GoTo 0
        
        ' FORCE INITIALIZATION to ensure state is set
        On Error Resume Next
        Application.Run "'PremiumAddon.xla'!PremiumCore.InitializePremium"
        On Error GoTo 0
        
        ' Reactivate this workbook to ensure add-in stays in background
        ThisWorkbook.Activate
    End If

    Debug.Print "  Elapsed: " & Format((Timer - startTime) * 1000, "0.00") & " ms"
    Debug.Print "LoadPremiumAddIn: Complete"
    Debug.Print String(60, "=")

End Sub

' ============================================================================
' PREMIUM UTILITY FUNCTIONS (Week 0)
' ============================================================================
Public Function IsPremiumAvailable() As Boolean
    ' Check if premium add-in is loaded and accessible

    On Error Resume Next

    ' Try to call a premium function
    ' Syntax: 'WorkbookName'!ModuleName.MacroName
    Application.Run "'PremiumAddon.xla'!PremiumCore.IsPremiumLoaded"

    ' If no error, premium is available
    IsPremiumAvailable = (Err.Number = 0)

    If IsPremiumAvailable Then
        Debug.Print "IsPremiumAvailable: True"
    Else
        Debug.Print "IsPremiumAvailable: False (Error: " & Err.Number & ")"
    End If

End Function

Public Function GetPremiumVersionIfAvailable() As String
    ' Get premium version if loaded

    On Error Resume Next

    If IsPremiumAvailable() Then
        GetPremiumVersionIfAvailable = Application.Run("'PremiumAddon.xla'!PremiumCore.GetPremiumVersion")
    Else
        GetPremiumVersionIfAvailable = "Not Loaded"
    End If

End Function

' ============================================================================
' PATH RESOLUTION HELPER (Week 1)
' ============================================================================
Private Function GetPremiumPath_Private() As String
    ' Resolve path to PremiumAddon.xla dynamically
    ' Priority:
    ' 1. Environment Variable (TURBOPOS_PREMIUM_PATH)
    ' 2. Relative Path (..\..\vba-pos-premium\src\)
    
    On Error Resume Next
    
    Dim envPath As String
    Dim relativePath As String
    Dim rootPath As String
    
    ' 1. Check Environment Variable
    envPath = Environ("TURBOPOS_PREMIUM_PATH")
    If envPath <> "" Then
        If Dir(envPath) <> "" Then
            GetPremiumPath_Private = envPath
            Exit Function
        End If
    End If
    
    ' 2. Check Relative Path based on ThisWorkbook
    rootPath = ThisWorkbook.Path
    
    ' HANDLE ONEDRIVE URLS
    ' If path is HTTP, convert to local path first
    If Left(rootPath, 4) = "http" Then
        rootPath = ConvertOneDrivePath(rootPath)
    End If
    
    ' Now proceed with local path manipulation
    ' Structure: [root]\vba-pos-excel-sqlite\src\pos.xlsm
    ' Target:    [root]\vba-pos-premium\src\PremiumAddon.xla
    
    Dim parts() As String
    parts = Split(rootPath, Application.PathSeparator)
    
    ' Remove 'src'
    If UBound(parts) > 0 Then ReDim Preserve parts(UBound(parts) - 1)
    ' Remove 'vba-pos-excel-sqlite'
    If UBound(parts) > 0 Then ReDim Preserve parts(UBound(parts) - 1)
    
    rootPath = Join(parts, Application.PathSeparator)
    relativePath = rootPath & "\vba-pos-premium\src\PremiumAddon.xla"
    
    Debug.Print "  Path Calc: " & relativePath
    
    If Dir(relativePath) <> "" Then
        GetPremiumPath_Private = relativePath
        Exit Function
    End If
    
    GetPremiumPath_Private = ""
    
End Function

Private Function ConvertOneDrivePath(ByVal url As String) As String
    ' Helper to convert OneDrive URL to Local Path
    ' URL: https://d.docs.live.net/ID/Documents/path...
    ' Local: %OneDrive%\Documents\path...
    
    On Error Resume Next
    
    ' Normalize separators
    url = Replace(url, "/", "\")
    
    ' Try to find "\Documents\" which is common
    Dim pos As Long
    pos = InStr(1, url, "\Documents\", vbTextCompare)
    
    If pos > 0 Then
        Dim distinctPath As String
        distinctPath = Mid(url, pos) ' \Documents\barrier\...
        
        ' Combine with OneDrive Root
        Dim oneDriveRoot As String
        oneDriveRoot = Environ("OneDrive")
        
        If oneDriveRoot <> "" Then
            ConvertOneDrivePath = oneDriveRoot & distinctPath
            Exit Function
        End If
    End If
    
    ' Fallback: return as is (might fail Dir but keeps logic safe)
    ConvertOneDrivePath = url
End Function
