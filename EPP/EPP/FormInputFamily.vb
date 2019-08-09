Public Class FormInputFamily
    Private bs As New BindingSource
    Dim myrow As DataRowView
    Public Sub New(ByRef bs As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()
        myrow = bs.Current
        Me.bs = bs
        ' Add any initialization after the InitializeComponent() call.
        TextBox1.DataBindings.Add("Text", bs, "familyid")
        TextBox2.DataBindings.Add("Text", bs, "familyname")

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ErrorProvider1.SetError(TextBox1, "")
        ErrorProvider1.SetError(TextBox2, "")
        Me.Validate()
        If Not checktextbox Then
            Me.DialogResult = DialogResult.None
            bs.CancelEdit()
            Exit Sub
        End If
        bs.EndEdit()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        bs.CancelEdit()
    End Sub

    Private Function checktextbox() As Boolean
        If CInt(TextBox1.Text) <= -1 Then
            ErrorProvider1.SetError(TextBox1, "Not a valid value.")
            Return False
        End If
        If TextBox2.Text = "" Then
            ErrorProvider1.SetError(TextBox2, "Value cannot be blank.")
            Return False
        End If
        Return True
    End Function

End Class