Attribute VB_Name = "modExecuteLine"
Option Explicit

Public s$

Private Sub Test_ExecuteLine()
    Dim sCode$
    
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
    
    MsgBox s
End Sub

Sub ExecuteLine(sCode)
    Application.Run "'StumbSub 0&: " & sCode & "'"
End Sub
Sub StumbSub(Optional ByVal stumb As Long)
End Sub
Sub Strumb2()

End Sub
