Attribute VB_Name = "modVBAEval"
Option Explicit

Public s$

Private Sub Examples_ExecuteLine()
    Dim sCode$
    
    sCode = _
    "?""Hellow "":" & _
    "?""World"""
    ExecuteLine sCode

    sCode = _
    "For i& = 0 To 11:" & _
        "Select Case True:" & _
        "Case i And 1:" & _
            "? i:" & _
        "End Select:" & _
    "Next"
    ExecuteLine sCode
    
    sCode = _
    "s$ = vbNullString:" & _
    "For i& = 0 To 11:" & _
        "Select Case True:" & _
        "Case i And 1:" & _
            "s$ = s & i & vbNewLine: " & _
        "End Select:" & _
    "Next:" & _
    "MsgBox s"
    ExecuteLine sCode
    
    MsgBox s
End Sub

Sub ExecuteLine(sCode)
    Application.Run "'StumbSub 0&: " & sCode & "'"
End Sub
Sub StumbSub(Optional ByVal stumb As Long)
End Sub

Private Sub Examples_ExecuteExpression()
    Debug.Print ExecuteExpression("Rnd")
    Debug.Print ExecuteExpression("Activecell.Parent.Parent.FullName")
    Debug.Print ExecuteExpression( _
        "CStr(CDec(""426632324442343,243242122"")*CDec(""5465421321654,645334323""))")
End Sub
Function ExecuteExpression(Param)
    Static ret, execFlg As Boolean
    If execFlg Then
        If Not IsObject(Param) Then ret = Param Else Set ret = Param
    Else
        execFlg = True
        Application.Run "'ExecuteExpression " & Param & "'"
        execFlg = False
        If Not IsObject(Param) Then
            ExecuteExpression = ret
        Else: Set ExecuteExpression = ret
        End If
    End If
End Function
