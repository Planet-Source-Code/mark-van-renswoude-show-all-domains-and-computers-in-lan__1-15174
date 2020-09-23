VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComputers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// This class was put together by Powersoft Programming.
'// You can find us at: http://psprogramming.virtualave.net/
'//
'// This class will provide the functionality
'// to enumerate all computers in a domain of your network.
'//
'// Feel free to use this class in your own projects,
'// just make sure this description is also included.


Private cComputers As New Collection
Private sDomain As String

Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, ByVal lpBuffer As Long, ByRef lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function StrLenA Lib "kernel32" Alias "lstrlenA" (ByVal Ptr As Long) As Long
Private Declare Function StrCopyA Lib "kernel32" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Const RESOURCE_GLOBALNET = &H2
Private Const RESOURCETYPE_DISK = &H1
Private Const ERROR_MORE_DATA = 234
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As Long
    lpRemoteName As Long
    lpComment As Long
    lpProvider As Long
End Type

Private Type NETRES2
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type
Public Property Get Domain() As String
    Domain = sDomain
End Property

Public Property Let Domain(Value As String)
    sDomain = Value
End Property


Public Function GetCount() As Long
    '// Return the number of computers in
    '// the collection.
    GetCount = cComputers.Count
End Function

Public Function GetItem(Index As Long) As String
    '// Check for legal value
    If Index < 1 Or Index > cComputers.Count Then
        '// Raise an error
        Err.Raise 7771, , "The index is not within valid range!"
        Exit Function
    End If
    
    '// Return an item in the domains collection
    GetItem = cComputers.Item(Index)
End Function
Public Sub Refresh()
    Dim lBufferPtrTemp As Long
    Dim sComputer As String
    Dim tNetRes As NETRES2
    Dim tNR As NETRESOURCE
    Dim lBufferPtr As Long
    Dim lEnumHwnd As Long
    Dim lReturn As Long
    Dim lBuffer As Long
    Dim lCount As Long
    Dim lK As Long
    
    '// Check if domain is already set
    If Len(sDomain) = 0 Then
        '// Raise an error
        Err.Raise 7772, , "The domain has not been set!"
        Exit Sub
    End If

    With tNetRes
        .lpRemoteName = sDomain
        .dwDisplayType = 1
    End With
    
    lReturn = WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_DISK, 0, tNetRes, lEnumHwnd)
    
    If lReturn <> 0 Then
        '// Raise error
        Err.Raise 7773, , "Could not enumerate computers in domain!"
        Exit Sub
    End If
    
    lBuffer = 1024 * 10
    lBufferPtr = GlobalAlloc(GPTR, lBuffer)
    Do
        '// Number of entries to return from enumeration:
        '// -1 causes all objects to be returned
        lCount = -1
        lReturn = WNetEnumResource(lEnumHwnd, lCount, lBufferPtr, lBuffer)
        
        If lReturn = ERROR_MORE_DATA Then
            '// Enumeration indicates that the lBufferPtr
            '// is not big enough to hold all of the
            '// information in the NETRESOURCE structure.
            '// lBuffer has been updated to hold the required
            '// amount of space.
            
            '//Free up memory
            GlobalFree lBufferPtr
            
            '// Allocate a new space for the lBuffer
            '// requested by the enumeration
            lBufferPtr = GlobalAlloc(GPTR, lBuffer)
        Else
            If lReturn = 0 Then
                lBufferPtrTemp = lBufferPtr
                For lK = 1 To lCount
                    CopyMemory tNR, ByVal lBufferPtrTemp, LenB(tNR)
                    sComputer = PointerToAsciiStr(tNR.lpRemoteName)
                    
                    If sComputer <> "" Then
                        sComputer = Mid(sComputer, InStr(sComputer, "\\") + 2)
                        
                        '// Add computer to the collection
                        cComputers.Add sComputer
                    End If
                    
                    '// Step forward in the buffer by
                    '// the length of the copied structure
                    lBufferPtrTemp = lBufferPtrTemp + LenB(tNR)
                Next
            End If
        End If
    Loop Until lCount = 0
    
    If lEnumHwnd <> 0 Then
        lReturn = WNetCloseEnum(lEnumHwnd)
    End If
    
    '// Free up memory
    GlobalFree lBufferPtr
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

