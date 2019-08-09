Public Class FormInputKAM
    Private myrow As DataRowView
    Private bs As BindingSource
    Public Sub New(ByRef bs As BindingSource)
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        TextBox1.DataBindings.Add("Text", bs, "employeenumber")
        TextBox2.DataBindings.Add("Text", bs, "employeename")
        myrow = bs.Current
        Me.bs = bs

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Validate()
        bs.EndEdit()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        bs.CancelEdit()
    End Sub
End Class