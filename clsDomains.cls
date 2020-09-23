VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDomains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// This class was put together by Powersoft Programming.
'// You can find us at: http://psprogramming.virtualave.net/
'//
'// This class will provide the functionality
'// to enumerate all domains in your network.
'//
'// Feel free to use this class in your own projects,
'// just make sure this description is also included.


Private cDomains As New Collection

Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lppEnumHwnd As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal pEnumHwnd As Long, lpcCount As Long, lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal p_lngEnumHwnd As Long) As Long
Private Declare Function StrLenA Lib "kernel32" Alias "lstrlenA" (ByVal Ptr As Long) As Long
Private Declare Function StrCopyA Lib "kernel32" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As Long) As Long

Private Const RESOURCE_GLOBALNET As Long = &H2&
Private Const RESOURCETYPE_ANY As Long = &H0&
Private Const RESOURCEUSAGE_ALL As Long = &H0&
Private Const RESOURCE_ENUM_ALL As Long = &HFFFF
Private Const NO_ERROR As Long = 0&
Private Const MAX_RESOURCES As Long = 256

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    pLocalName As Long
    pRemoteName As Long
    pComment As Long
    pProvider As Long
End Type

Public Function GetCount() As Long
    '// Return the number of domains in
    '// the collection.
    GetCount = cDomains.Count
End Function


Public Function GetItem(Index As Long) As String
    '// Check for legal value
    If Index < 1 Or Index > cDomains.Count Then
        '// Raise an error
        Err.Raise 7771, , "The index is not within valid range!"
        Exit Function
    End If
    
    '// Return an item in the domains collection
    GetItem = cDomains.Item(Index)
End Function


Public Sub Refresh()
    Dim tNetAPI(0 To MAX_RESOURCES) As NETRESOURCE
    Dim lBufferSize As Long
    Dim lEnumHwnd As Long
    Dim lReturn As Long
    Dim lCount As Long
    Dim lLoop As Long
    
    '// Clear the domains collection
    Do While cDomains.Count > 0
        cDomains.Remove 1
    Loop
    
    '// Refresh the domains collection
    lReturn = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, RESOURCEUSAGE_ALL, tNetAPI(0), lEnumHwnd)
    
    If lReturn = NO_ERROR Then
        lCount = RESOURCE_ENUM_ALL
        
        lBufferSize = UBound(tNetAPI) * Len(tNetAPI(0))
        lReturn = WNetEnumResource(lEnumHwnd, lCount, tNetAPI(0), lBufferSize)
        
        If lCount > 0 Then
            For lLoop = 0 To lCount - 1
                '// Add domain to collection
                cDomains.Add PointerToAsciiStr(tNetAPI(lLoop).pRemoteName)
            Next lLoop
        End If
    End If
    
    If lEnumHwnd <> 0 Then
        Call WNetCloseEnum(lEnumHwnd)
    End If
End Sub

Private Function PointerToAsciiStr(ByVal lPtrToString As Long) As String
    On Local Error Resume Next
    Dim lLength As Long
    Dim sStringValue As String
    Dim lNullPos As Long
    Dim lReturn As Long
    
    lLength = StrLenA(lPtrToString)
    
    If lPtrToString > 0 And lLength > 0 Then
        '// Copy Pointer to String
        sStringValue = Space$(lLength + 1)
        lReturn = StrCopyA(sStringValue, lPtrToString)
        
        '// Find null terminator
        lNullPos = InStr(1, sStringValue, Chr$(0))
        
        If lNullPos > 0 Then
            '// Lose the null terminator
            PointerToAsciiStr = Left$(sStringValue, lNullPos - 1)
        Else
            '// Just pass the string
            PointerToAsciiStr = sStringValue
        End If
    Else
        PointerToAsciiStr = ""
    End If
End Function



Private Sub Class_Initialize()
    '// Refresh the domains
    Call Refresh
End Sub


Private Sub Class_Terminate()
    '// Clear the collection
    Do While cDomains.Count > 0
        cDomains.Remove 1
    Loop
End Sub


