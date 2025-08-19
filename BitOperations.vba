Option Explicit

Private Type tCur
    val As Currency
End Type
Private Type tLng
    val As Long
End Type

'https://www.cyberforum.ru/visual-basic/thread1479493.html
Private isDeg2Init As Boolean, Deg2(31) As Long
Private isDeg3Init As Boolean, Deg3(31) As Currency

'Dim LNG_POW2(32)
Private Const Log2 As Double = 0.693147180559945 'Log(2)

'End Type CCAB
Private Sub initDeg2()
    Dim i&
    If isDeg2Init Then Exit Sub
    Deg2(0) = 1
    For i = 1 To 30
        Deg2(i) = 2 ^ i
    Next
    Deg2(i) = -2 ^ i
    isDeg2Init = True
End Sub
Private Sub initDeg3()
    Dim i&
    If isDeg3Init Then Exit Sub
    Deg3(0) = 1 * 0.0001@
    For i = 1 To 31
        Deg3(i) = 2@ ^ i * 0.0001@
    Next
    isDeg3Init = True
End Sub

'Показать бит
Function BitGet(ByVal m&, n As Byte) As Boolean
    If isDeg2Init Then Else initDeg2
    BitGet = m And Deg2(n)
End Function
'Выставить бит
Function BitPut(ByVal m&, ByVal n As Byte) As Long
    If isDeg2Init Then Else initDeg2
    BitPut = m Or Deg2(n)
End Function
'Сбросить (обнулить) бит
Function BitOut(ByVal m&, n As Byte) As Long
    If isDeg2Init Then Else initDeg2
    BitOut = m And Not Deg2(n)
End Function
'Перекинуть (поменять) бит
Function BitSwp(ByVal m&, n As Byte) As Long
    If isDeg2Init Then Else initDeg2
    BitPut = m Xor Deg2(n)
End Function

'Битовые сдвиги
Function LeftBitShift(ByVal lNum As Long, ByVal bitCnt As Long) As Long
    Dim tCur As tCur, tLng As tLng
    If isDeg3Init Then Else initDeg3
    tCur.val = lNum * Deg3(bitCnt)
    LSet tLng = tCur
    LeftBitShift = tLng.val
End Function
Function RightBitShift(ByVal l As Long, ByVal bitCnt As Long) As Long
    If isDeg2Init Then Else initDeg2
    RightBitShift = l \ Deg2(bitCnt) '2 ^ bitCnt
End Function

'Getting the position of the most significant bit
Function GetHighestBitPosition(ByVal number As Long) As Long
    If number Then Else Exit Function
    GetHighestBitPosition = Int(Log(number) / Log2)
End Function
'Getting the position of the least significant bit
Function GetLowestBitPosition(ByVal number As Long) As Long
    If number = 0 Then Exit Function
    GetLowestBitPosition = Log(number And -number) / Log2
End Function

Private Sub testLeftShift()
    Dim l&
'    l = 14
'    initDeg2
'    Debug.Print GetBitMaskS(VarPtr(l), 4)
    l = RightBitShift(1073741500, 29)
'    Debug.Print GetBitMaskS(VarPtr(l), 4)
    l = LeftBitShift(85471223, 31)
'    Debug.Print GetBitMaskS(VarPtr(l), 4)
End Sub
Private Sub TestBitOp()
    Dim l&, i&
    initDeg2
    For i = 0 To 31
        l = BitPut(l, i)
        Debug.Print BitPut(l, 1), BitOperation.BitPut(0, i)
    Next
End Sub

'Function LeftBitShift2(ByVal l As Long, ByVal bitCnt As Long) As Long
'    LeftBitShift2 = l * Deg2(bitCnt) '2 ^ bitCnt
'End Function
'Private Function LShift32(ByVal lX As Long, ByVal lN As Long) As Long
'    If lN = 0 Then
'        LShift32 = lX
'    Else
'        LShift32 = (lX And (LNG_POW2(31 - lN) - 1)) * LNG_POW2(lN) Or -((lX And LNG_POW2(31 - lN)) <> 0) * &H80000000
'    End If
'End Function
'Function Left8BitShift(ByVal lVal As Long) As Long
'    Left8BitShift = lVal * h100
'End Function
