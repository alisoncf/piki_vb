Attribute VB_Name = "modulo"
Const SQL_NULL_HANDLE = 0
Const SQL_HANDLE_ENV = 1
Const SQL_FETCH_NEXT = 1
Const SQL_FETCH_FIRST = 2
Const SQL_SUCCESS = 0
Const SQL_ATTR_ODBC_VERSION = 200
Const SQL_OV_ODBC2 = 2
Const SQL_IS_INTEGER = -6

Dim nRetCode As Long

Private Declare Function SQLDrivers Lib "odbc32.dll" (ByVal _
    EnvironmentHandle As Long, ByVal Direction As Integer, _
    ByVal DriverDescription As String, ByVal BufferLength1 As Integer, _
    DescriptionLengthPtr As Integer, ByVal DriverAttributes As String, _
    ByVal BufferLength2 As Integer, AttributesLengthPtr As Integer) As Integer

' Note that pointers to numbers are passed as numbers by reference!
Private Declare Function SQLDataSources Lib "odbc32.dll" (ByVal _
    EnvironmentHandle As Long, ByVal Direction As Integer, _
    ByVal ServerName As String, ByVal BufferLength1 As Integer, _
    NameLength1Ptr As Integer, ByVal Description As String, _
    ByVal BufferLength2 As Integer, NameLength2Ptr As Integer) As Integer


Private Declare Function SQLFreeHandle Lib "odbc32.dll" (ByVal _
    HandleType As Integer, ByVal Handle As Long) As Integer

Private Declare Function SQLAllocHandle Lib "odbc32.dll" (ByVal _
    HandleType As Integer, ByVal InputHandle As Long, _
    OutputHandlePtr As Long) As Integer

Private Declare Function SQLSetEnvAttr Lib "odbc32.dll" (ByVal _
    EnvironmentHandle As Long, ByVal EnvAttribute As Long, _
    ByVal ValuePtr As Long, ByVal StringLength As Long) As Integer


Private Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal _
    ConnectionHandle As Long) As Integer

'//This function will print list of installed driver name and attributes
Sub main()
    frmMain.Show 1
End Sub
Public Function ListODBCDrivers() As Collection
    Dim c As New Collection
  ' Prints a list of ODBC drivers on system

  Dim lHEnv       As Long
  Dim sDriverDesc    As String * 1024
  Dim sDriverAttr    As String * 1024
  Dim sDriverAttributes As String
  Dim nDriverDescLength As Integer
  Dim nAttrLength As Integer
  Dim x As Integer
  Dim sAll As String

  ' Allocate an environment handle.
  nRetCode = SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, lHEnv)

  ' Set ODBC behavior
  nRetCode = SQLSetEnvAttr(lHEnv, SQL_ATTR_ODBC_VERSION, _
      SQL_OV_ODBC2, SQL_IS_INTEGER)

  ' Get first driver
  nRetCode = SQLDrivers(lHEnv, _
      SQL_FETCH_FIRST, _
      sDriverDesc, _
      Len(sDriverDesc), _
      nDriverDescLength, _
      sDriverAttr, _
      Len(sDriverAttr), _
      nAttrLength)

  sAll = ""

  Dim strSQL As String
    Dim iTab As Integer
  Do While nRetCode = SQL_SUCCESS

    '//Replace NULL separators with colons in attribute string
    sDriverAttributes = Left$(sDriverAttr, nAttrLength - 1)
    sDriverAttributes = Replace(sDriverAttributes, Chr(0), " : ")

    '//Remove NULL in ODBC Driver name
    sDriverDesc = Left$(sDriverDesc, nDriverDescLength)
    'sDriverDesc = Replace(sDriverDesc, Chr(32), "%")
    iTab = InStr(1, sDriverDesc, Chr(32))
    
    sDriverDesc = RTrim(sDriverDesc)
    c.Add (sDriverDesc)
    
'    For i = 0 To Len(sDriverDesc) - 1
      '  MsgBox Asc(Mid(sDriverDesc, i, 1))
 '   Next
    
    sAll = sAll & sDriverDesc & " ==> " & sDriverAttributes & vbCrLf

    strSQL = strSQL & sDriverDesc                 '//Print driver name
    'strSQL = strSQL & sDriverAttributes             '//Print driver attribute
    'strSQL = strSQL & String(100, "-")

    
  
  
    nRetCode = SQLDrivers(lHEnv, _
        SQL_FETCH_NEXT, _
        sDriverDesc, _
        Len(sDriverDesc), _
        nDriverDescLength, _
        sDriverAttr, _
        Len(sDriverAttr), _
        nAttrLength)
  Loop

  'Debug.Print "ODBC Drivers"
  'Debug.Print sAll

  '  MsgBox Trim(strSQL)
  nRetCode = SQLFreeHandle(SQL_HANDLE_ENV, lHEnv)
  
  Set ListODBCDrivers = c
  
End Function



Function GravaArquivoExterno(caminho As String, conteudo) As Boolean
    
    On Error Resume Next
    Close #1
    
    Open caminho For Output As #1
    Print #1, conteudo
    
    Close #1
    GravaArquivoExterno = True
End Function

