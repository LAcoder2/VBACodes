This example shows how to auto-initialize the data needed for a function to work. To do this, you need to create an additional function - an initializer, to which we assign the name of the original function (we change the name of the main function). When first launched, the initializer initializes (tautology, sorry) all the necessary data, then replaces its pointer with the pointer of the main function. And returns the result of the main function. Subsequent calls will be redirected to the main function. This trick is not optimal for use in VBA (although it works in VBA 32bit), since VBA often experiences "loss of state" - reset of all data due to various errors, even implicit ones. With such a reset, the changed function pointers are not reset - they are reset only when the code is changed and recompiled.
```vba
Private Sub Example()
    Debug.Print BaseXEncode(565654651, 36)
    Debug.Print BaseXEncode(565654654, 36)    
End Sub
```
```vba
Option Explicit
' Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Dst As Any, Src As Any, ByVal ln As Long)
Private Declare Sub GetMem4 Lib "msvbvm60" (Src As Any, Dst As Any) 'As Long
Private Declare Sub GetMem8 Lib "msvbvm60" (Src As Any, Dst As Any) 'As Long
Private Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Const PAGE_EXECUTE_READWRITE = &H40

Private Const MaxBase As Long = 36
Private ChrTbl$(MaxBase - 1)

Function BaseXEncode(ByVal lNum As Long, ByVal Base As Long) As String
    Static Init As Boolean
    Dim i&, sChrRng$
    Debug.Print "init proc"
    If Init Then
    Else
        sChrRng = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        For i = 0 To Base - 1
            ChrTbl(i) = Mid$(sChrRng, i + 1, 1)
        Next
        Init = True
    End If
    RplaceFunPtr AddressOf Module1.BaseXEncode, AddressOf Module1.BaseXEncode_
    BaseXEncode = BaseXEncode_(lNum, Base)
End Function
Private Function BaseXEncode_(ByVal lNum As Long, ByVal Base As Long) As String
    Select Case Base
    Case 2 To MaxBase
        Do
            BaseXEncode_ = ChrTbl(lNum Mod Base) & BaseXEncode_
            lNum = lNum \ Base
        Loop While lNum
    End Select
End Function
'Based on patch by The trick: https://www.cyberforum.ru/visual-basic/thread1150127-page3.html#post8172932
Sub RplaceFunPtr(ByVal AddrDst As Long, ByVal AddrSrc As Long)
    Dim InIDE As Boolean
    Debug.Assert MakeTrue(InIDE)
    If InIDE Then
        GetMem4 ByVal AddrDst + &H16, AddrDst
        GetMem4 ByVal AddrSrc + &H16, AddrSrc
    Else
        VirtualProtect AddrDst, 8, PAGE_EXECUTE_READWRITE, 0
    End If
    GetMem8 ByVal AddrSrc, ByVal AddrDst
End Sub
Public Function MakeTrue(ByRef blVar As Boolean) As Boolean
    blVar = True: MakeTrue = True
End Function
```
