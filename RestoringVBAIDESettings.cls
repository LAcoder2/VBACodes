Option Explicit
'This code saves when closing the book and restores the VBAIDE editor state settings (active module, highlighted area and scrolling)
'The code needs to be placed in the workbook object module

Private Sub Workbook_Open()
    RestVBIDEOpts
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    SaveVBIDEOpts
End Sub

Sub RestVBIDEOpts()
    Const q$ = """"
    Dim Opts$(), vsCnt&, RestFlg&, wnd As Window, mdlTp&
    Opts = Split(Me.CustomDocumentProperties("CurrVBIDEPos"), ",")
    With ThisWorkbook.VBProject.VBComponents(Opts(0))
'        mdlTp = .Type
'        If mdlTp = 1 Then .Activate 'vbext_ct_StdModule' 1й вариант активации модуля, работает правильно только со стандартными модулями
        With .CodeModule.CodePane
            .TopLine = Opts(1)
            .SetSelection Opts(2), Opts(3), Opts(4), Opts(5) 
            With .VBE.MainWindow '                        2й вариант активации модуля (с мельканием окна VBA IDE)
                If Not .Visible Then
                    .Visible = True ':DoEvents: DoEvents
                Else: Opts(6) = "-1"
                End If
                Application.OnTime Now, "'ЭтаКнига.ActivateCP " & q & Opts(0) & q & ", " & Opts(6) & "&'"
            End With
        End With
    End With
End Sub
Sub ActivateCP(ByVal CPNm As String, ByVal vs As Long)
    With ThisWorkbook.VBProject.VBComponents(CPNm).CodeModule.CodePane
        .Show
        If vs = 0 Then
            .VBE.MainWindow.Visible = False
        End If
    End With
End Sub

Sub SaveVBIDEOpts()
    Const optName$ = "CurrVBIDEPos", dlm$ = ","
    Dim Opt As DocumentProperty, sVal$
    Dim tl&, sl&, sc&, el&, ec&, vs&
    
    With ThisWorkbook.VBProject.VBE.ActiveCodePane.CodeModule
        .CodePane.GetSelection sl, sc, el, ec
'        Stop
        tl = .CodePane.TopLine
        vs = .VBE.MainWindow.Visible
        sVal = .Name & dlm & tl & dlm & sl & dlm & sc & dlm & el & dlm & ec & dlm & vs
    End With
    For Each Opt In Me.CustomDocumentProperties
        If Opt.Name = optName Then Exit For
    Next
    If Not Opt Is Nothing Then
        Opt.Value = sVal
    Else
        Me.CustomDocumentProperties.Add optName, False, msoPropertyTypeString, sVal
    End If
End Sub
