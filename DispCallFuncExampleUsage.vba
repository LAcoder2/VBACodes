Option Explicit

'Enum LongPtr
'    [_]
'End Enum
Enum HRESULT
    [_]
End Enum
Enum BOOL
    apiFALSE
    apiTRUE
End Enum
Enum CALLCONV
    CC_CDECL = 1
    CC_STDCALL = 4
End Enum
Private Const NullPtr As LongPtr = 0
Private Declare Function DispCallFunc Lib "oleaut32.dll" ( _
                            ByVal pvInstance As LongPtr, _
                            ByVal oVft As Long, _
                   Optional ByVal cc As CALLCONV = CC_STDCALL, _
                   Optional ByVal vtReturn As VbVarType, _
                   Optional ByVal cntArgs As Long, _
                         Optional prgvt As Integer, _
                         Optional prgpvarg As LongPtr, _
                         Optional pvargResult As Variant) As HRESULT
Public Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Dst As Any, Src As Any, ByVal Ln As LongPtr)

Function DispCallHelper(ByVal pObj As LongPtr, ByVal pFun As LongPtr, ByVal callType As CALLCONV, _
                                                      ByVal vtReturn As VbVarType, ParamArray arInp())
    Dim Types(63) As Integer
    Dim Ptrs(63)  As LongPtr
    Dim lRes as HRESULT&, i&
    
    For i = 0 To UBound(arInp)
        CopyMemory Types(i), arInp(i), 2
        Ptrs(i) = VarPtr(arInp(i))
    Next
    
    lRes = DispCallFunc(pObj, pFun, callType, vtReturn, i, Types(0), Ptrs(0), DispCallHelper)
End Function

Private Function TestFunction(ByVal lBV&, lBR&) As Double
    TestFunction = lBV * lBR
End Function
Private Function TestFunction2(vArg)
    vArg = 555
    TestFunction2 = 123.456
End Function

Private Sub TestDispCallHelper()
    Dim vRet, obj As Object, vArg, lRes&, lBV&, lBR&, lCnt&
    
    'Вызов функции по указателю
    'пример передачи типизированных аргументов - ByVal нужно поместить в скобки
    lBV = 2: lBR = 7
    vRet = DispCallHelper(0, AddressOf TestFunction, CC_STDCALL, vbDouble, (lBV), lBR)
    
    'Пример передачи вариантных аргументов ByRef (Variant обычно передается ByRef) и возврата Variant
    vArg = 123
'    vRet = DispCallHelper(0, AddressOf TestFunction2, CC_STDCALL, vbVariant, VarPtr(vArg)) 'v1
    
    DispCallHelper 0, AddressOf TestFunction2, CC_STDCALL, 0, VarPtr(vRet), VarPtr(vArg)   'v2 вызов обычной функции как процедуры
    
'    lRes = DispCallFunc(ByVal 0, AddressOf TestFunction2, , vbVariant, 1, _
                                                 vbLong, VarPtr(CVar(VarPtr(vArg))), vRet) 'v3
    'Вызов свойства com-объекта
    'При вызове com-интерфейса функция должна возвращать lRes, а собственный ответ функция возвращает
    'в последний аргумент (который нужно дополнительно добавлять)
    Set obj = CreateObject("Scripting.Dictionary")
    obj.Add "key1", "item1"
    vArg = "key1"
    lRes = DispCallHelper(ObjPtr(obj), &H24, CC_STDCALL, vbLong, VarPtr(vArg), VarPtr(vRet)) 'item get
    'Вместо VarPtr(vArg) можно использовать VarPtr(CVar("key1"))
'    lRes = DispCallHelper(dict, &H2C, CC_STDCALL, vbLong, lCnt)                    'v1 (dict.Count)
'    lRes = DispCallHelper(Obj, &H2C, CC_STDCALL, 0, lCnt)                          'v1 вызов как процедуры (без возврата hresult)
    DispCallFunc ObjPtr(obj), &H2C, , , 1, vbLong, VarPtr(CVar(VarPtr(lCnt)))       'v2
End Sub
