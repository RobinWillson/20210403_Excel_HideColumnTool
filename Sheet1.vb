Private Sub WorkSheet_BeforeDoubleClick( _
    ByVal Target As Range, Cancel As Boolean)
    'Target.Select
    Unprotect "1234"

    If Target.Column = 3 Then
        If (Target.Row >= 5 And Target.Row <= 15) Or _
           (Target.Row >= 21 And Target.Row <= 35) Then
            If Target.Value <> "V" Then
                Target.Value = "V"
                GoTo mCancel
            End If
            If Target.Value = "V" Then Target.Value = ""
        End If
    End If
mCancel:
    Cancel = True
    
    Protect "1234"
End Sub