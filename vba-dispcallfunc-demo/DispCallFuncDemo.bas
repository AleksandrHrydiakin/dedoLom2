Attribute VB_Name = "DispCallFuncDemo"
Option Explicit

' ---------------------------------------------------------------------------
' Benign DispCallFunc demo for x64 Excel / VBA on Windows.
'
' Shows how DispCallFunc (exported by OleAut32) acts as a calling-convention
' bridge: it lets VBA invoke a raw function pointer with arbitrary argument
' types, marshalling them into the Microsoft x64 ABI on our behalf.
'
' Both demos call ordinary, exported Win32 APIs whose addresses we get the
' normal way (LoadLibraryW + GetProcAddress). This is purely to illustrate
' the mechanism -- nothing here resolves syscall numbers, allocates RWX
' memory, or executes a runtime-built stub.
' ---------------------------------------------------------------------------

' DispCallFunc CALLCONV value. On x64 Windows there is only one calling
' convention; CC_STDCALL is what OleAut32 expects.
Private Const CC_STDCALL As Long = 4

' VARTYPE constants we use below.
Private Const VT_I4      As Integer = 3    ' 32-bit signed int
Private Const VT_UI4     As Integer = 19   ' 32-bit unsigned int
Private Const VT_I8      As Integer = 20   ' 64-bit signed int  (used for x64 pointers/handles)

' DispCallFunc:
'   HRESULT DispCallFunc(
'       void*         pvInstance,   ' NULL -> oVft is treated as a raw function pointer
'       ULONG_PTR     oVft,         ' function pointer when pvInstance == NULL
'       CALLCONV      cc,
'       VARTYPE       vtReturn,     ' return type
'       UINT          cActuals,     ' number of arguments
'       VARTYPE*      prgvt,        ' array of VARTYPE, one per argument
'       VARIANTARG**  prgpvarg,     ' array of pointers to VARIANTs holding the values
'       VARIANT*      pvargResult); ' out: result
Private Declare PtrSafe Function DispCallFunc Lib "OleAut32.dll" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal CallConv As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByRef prgvt As Integer, _
    ByRef prgpvarg As LongPtr, _
    ByRef pvargResult As Variant) As Long

Private Declare PtrSafe Function LoadLibraryW Lib "kernel32" ( _
    ByVal lpLibFileName As LongPtr) As LongPtr

Private Declare PtrSafe Function GetProcAddress Lib "kernel32" ( _
    ByVal hModule As LongPtr, _
    ByVal lpProcName As String) As LongPtr


' ---------------------------------------------------------------------------
' Demo 1: GetTickCount  ->  DWORD GetTickCount(void)
' Zero arguments, 32-bit unsigned return. The simplest possible bridge call.
' ---------------------------------------------------------------------------
Public Sub Demo_GetTickCount()
    Dim hMod As LongPtr, pFn As LongPtr
    Dim hr As Long
    Dim result As Variant

    hMod = LoadLibraryW(StrPtr("kernel32.dll"))
    pFn = GetProcAddress(hMod, "GetTickCount")
    If pFn = 0 Then
        Debug.Print "GetProcAddress(GetTickCount) failed"
        Exit Sub
    End If

    ' cActuals = 0, but DispCallFunc still wants addressable arrays for prgvt/prgpvarg.
    Dim dummyVt As Integer
    Dim dummyArg As LongPtr

    hr = DispCallFunc( _
        0, _                ' pvInstance = NULL  -> oVft IS the function pointer
        pFn, _              ' GetTickCount address
        CC_STDCALL, _
        VT_UI4, _           ' return type
        0, _                ' no args
        dummyVt, _
        dummyArg, _
        result)

    If hr = 0 Then
        Debug.Print "GetTickCount via DispCallFunc -> " & CStr(result) & " ms since boot"
    Else
        Debug.Print "DispCallFunc failed, HRESULT = 0x" & Hex$(hr)
    End If
End Sub


' ---------------------------------------------------------------------------
' Demo 2: MessageBoxW  ->  int MessageBoxW(HWND, LPCWSTR text, LPCWSTR caption, UINT type)
' Four arguments; demonstrates passing 64-bit handles and pointers plus a 32-bit UINT,
' and getting a 32-bit int back. This is the case where DispCallFunc actually has
' to lay out a real x64 frame for us.
' ---------------------------------------------------------------------------
Public Sub Demo_MessageBoxW()
    Dim hMod As LongPtr, pFn As LongPtr
    Dim hr As Long
    Dim result As Variant

    hMod = LoadLibraryW(StrPtr("user32.dll"))
    pFn = GetProcAddress(hMod, "MessageBoxW")
    If pFn = 0 Then
        Debug.Print "GetProcAddress(MessageBoxW) failed"
        Exit Sub
    End If

    Dim sText As String, sCaption As String
    sText = "Hello from DispCallFunc!"
    sCaption = "DispCallFunc Demo"

    ' Each argument lives in a VARIANT. The internal storage type must match
    ' what we declare in the VARTYPE array.
    Dim vHwnd As Variant, vText As Variant, vCaption As Variant, vType As Variant
    vHwnd = CLngLng(0)                    ' HWND_DESKTOP equivalent: NULL
    vText = CLngLng(StrPtr(sText))        ' UTF-16 buffer pointer
    vCaption = CLngLng(StrPtr(sCaption))
    vType = CLng(0)                       ' MB_OK

    Dim vt(0 To 3) As Integer
    vt(0) = VT_I8                         ' HWND on x64 is 64-bit
    vt(1) = VT_I8                         ' LPCWSTR pointer
    vt(2) = VT_I8                         ' LPCWSTR pointer
    vt(3) = VT_I4                         ' UINT (32-bit)

    Dim args(0 To 3) As LongPtr
    args(0) = VarPtr(vHwnd)
    args(1) = VarPtr(vText)
    args(2) = VarPtr(vCaption)
    args(3) = VarPtr(vType)

    hr = DispCallFunc( _
        0, _
        pFn, _
        CC_STDCALL, _
        VT_I4, _                          ' MessageBoxW returns int
        4, _
        vt(0), _
        args(0), _
        result)

    If hr = 0 Then
        Debug.Print "MessageBoxW via DispCallFunc -> " & CStr(result)
    Else
        Debug.Print "DispCallFunc failed, HRESULT = 0x" & Hex$(hr)
    End If
End Sub


' Convenience entry point: run both demos.
Public Sub RunAllDemos()
    Demo_GetTickCount
    Demo_MessageBoxW
End Sub
