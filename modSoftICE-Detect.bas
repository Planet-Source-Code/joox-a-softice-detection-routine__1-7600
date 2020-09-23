Attribute VB_Name = "modDetect"
' SoftICE-Detect last updated 3/31/2000
'
' E-mail:  joox@tech-productions.de
' Website: http://www.Tech-Productions.de
'
' What does this code?
' It will check if SoftICE from NuMega is loaded into memory!
'
' But what is SoftICE?
' SoftICE is a low-level Debugger for Windows(TM).
'
' Who needs this code?
' Everybody who want to distribute his VB-app as shareware!
'
' Why?
' SoftICE is also used by crackers who don't want to pay for your software!
' So they use the Debugger to register your software without paying!
'
' How could i protect my software?
' Just check if SoftICE is loaded with my function. If true close the app.
'
'
' Please send me a E-mail when you use this routine in your app
'
'
Public Declare Function CreateFileNS Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function WriteFileNO Lib "kernel32" Alias "WriteFile" (ByVal hfile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80

Public Function SoftICELoaded() As Boolean
Dim hfile As Long, retval As Long
    hfile = CreateFileNS("\\.\SICE", GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hfile <> -1 Then
        ' SoftICE is detected.
        retval = CloseHandle(hfile) ' Close the file handle
        SoftICELoaded = True
    Else
    ' SoftICE is not found.
    SoftICELoaded = False
    End If
End Function

Sub Main()
    If SoftICELoaded Then ' check if softice is loaded
        MsgBox "SoftICE is detected! Closing now!", vbMsgBoxSetForeground + vbInformation, "SoftICE-Detector by Joox"
        End ' if true finish the app
    End If
    MsgBox "SoftICE was not found in memory!", vbMsgBoxSetForeground + vbInformation, "SoftICE-Detector by Joox"
End Sub
