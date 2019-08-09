Public Class MyPanel
    Inherits Panel

    Protected Overrides Sub DefWndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg <> 131 Then
            MyBase.DefWndProc(m)
        End If
    End Sub

End Class
