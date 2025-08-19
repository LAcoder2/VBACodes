Attribute VB_Name = "modExecuteLine"
'Option Explicit

Private Sub Test_ExecuteLine()
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
End Sub

Sub ExecuteLine(sCode As String)
    Application.Run "'StumbSub 0: " & sCode & "'"
End Sub
Sub StumbSub(ByVal stumb As Long)
End Sub

