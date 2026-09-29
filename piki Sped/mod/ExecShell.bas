Attribute VB_Name = "ExecShell"
Option Explicit

Private Declare Function CreatePipe Lib "kernel32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As Any, _
    ByVal nSize As Long) As Long

Private Declare Function SetStdHandle Lib "kernel32" ( _
    ByVal nStdHandle As Long, _
    ByVal hHandle As Long) As Long

Private Declare Function GetStdHandle Lib "kernel32" ( _
    ByVal nStdHandle As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" ( _
    ByVal lpApplicationName As Long, _
    ByVal lpCommandLine As String, _
    ByVal lpProcessAttributes As Any, _
    ByVal lpThreadAttributes As Any, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Any, _
    ByVal lpCurrentDirectory As Any, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Const STARTF_USESTDHANDLES = &H100&
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_HIDE = 0

Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&

Public Function ExecCmd(cmdline As String) As String
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim hReadPipe As Long
    Dim hWritePipe As Long
    Dim buffer As String * 1024
    Dim bytesRead As Long
    Dim result As String
    
    ' Create the pipe
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1
    sa.lpSecurityDescriptor = 0
    CreatePipe hReadPipe, hWritePipe, sa, 0
    SetStdHandle STD_OUTPUT_HANDLE, hWritePipe
    SetStdHandle STD_ERROR_HANDLE, hWritePipe
    
    ' Initialize the STARTUPINFO structure
    start.cb = Len(start)
    start.dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
    start.wShowWindow = SW_HIDE
    start.hStdOutput = hWritePipe
    start.hStdError = hWritePipe
    
    ' Start the shelled application
    CreateProcessA 0&, cmdline, 0&, 0&, 1&, 0&, 0&, 0&, start, proc
    
    
    ' Read the output from the pipe
    Do
        ReadFile hReadPipe, buffer, Len(buffer), bytesRead, 0
        If bytesRead > 0 Then
            result = result & Left$(buffer, bytesRead)
        End If
    Loop While bytesRead > 0
    
    ' Close handles
    CloseHandle proc.hProcess
    CloseHandle proc.hThread
    CloseHandle hReadPipe
    CloseHandle hWritePipe
    
    ExecCmd = result
End Function


