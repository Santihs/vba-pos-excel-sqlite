Sub DebugPremiumConnection()
    ' Run this to test the connection to the Premium Add-in
    
    Dim result As String
    Dim isLoaded As Boolean
    
    Debug.Print "--- DEBUG START ---"
    
    ' 1. Check if Workbook is open (New Strategy)
    On Error Resume Next
    Dim wb As Workbook
    Set wb = Workbooks("PremiumAddon.xla")
    If wb Is Nothing Then
        Debug.Print "CRITICAL: PremiumAddon.xla is NOT OPEN in Workbooks collection."
        Debug.Print "Attempting to open it now..."
        
        ' Dynamic Path Logic (Week 1)
        Dim rootPath As String, relativePath As String, parts() As String
        rootPath = ThisWorkbook.Path
        
        ' FIX ONEDRIVE URLS
        If Left(rootPath, 4) = "http" Then
            ' Normalize: https://.../Documents/ -> %OneDrive%/Documents/
            Dim pos As Long
            rootPath = Replace(rootPath, "/", "\")
            pos = InStr(1, rootPath, "\Documents\", vbTextCompare)
            If pos > 0 Then
                rootPath = Environ("OneDrive") & Mid(rootPath, pos)
                Debug.Print "  OneDrive Localized Path: " & rootPath
            End If
        End If
        
        parts = Split(rootPath, Application.PathSeparator)
        If UBound(parts) > 0 Then ReDim Preserve parts(UBound(parts) - 1) ' Remove src
        If UBound(parts) > 0 Then ReDim Preserve parts(UBound(parts) - 1) ' Remove repo
        rootPath = Join(parts, Application.PathSeparator)
        relativePath = rootPath & "\vba-pos-premium\src\PremiumAddon.xla"
        
        Debug.Print "  Attemping Path: " & relativePath
        
        Set wb = Workbooks.Open(relativePath)
        If wb Is Nothing Then
             Debug.Print "FAILED to open file at path."
        Else
             Debug.Print "SUCCESS: Opened file manually."
        End If
    Else
        Debug.Print "SUCCESS: PremiumAddon.xla is OPEN."
    End If
    On Error GoTo 0
    
    ' 2. Try to call IsPremiumLoaded
    On Error Resume Next
    isLoaded = Application.Run("'PremiumAddon.xla'!PremiumCore.IsPremiumLoaded")
    If Err.Number <> 0 Then
        Debug.Print "Error calling IsPremiumLoaded: " & Err.Description & " (" & Err.Number & ")"
    Else
        Debug.Print "IsPremiumLoaded returned: " & isLoaded
    End If
    On Error GoTo 0
    
    ' 3. Try to force button creation
    On Error Resume Next
    Debug.Print "Attempting to create button manually..."
    Application.Run "'PremiumAddon.xla'!PremiumCore.CreatePremiumButton"
    If Err.Number <> 0 Then
        Debug.Print "Error calling CreatePremiumButton: " & Err.Description
    Else
        Debug.Print "CreatePremiumButton call successful (check sheet)"
    End If
    
    Debug.Print "--- DEBUG END ---"
End Sub
