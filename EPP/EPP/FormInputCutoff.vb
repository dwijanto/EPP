Public Class FormInputCutoff
    Private myrow As DataRowView
    Private bs As New BindingSource
    Public Sub New(ByRef bs As BindingSource)
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        Me.bs = bs
        'DateTimePicker1.DataBindings.Clear()
        'DateTimePicker2.DataBindings.Clear()
        DateTimePicker1.DataBindings.Add("value", Me.bs, "cutoff")
        DateTimePicker2.DataBindings.Add("value", Me.bs, "emailnotificationdate")
        DateTimePicker3.DataBindings.Add("value", Me.bs, "chequesubmitdate")
        DateTimePicker4.DataBindings.Add("value", Me.bs, "collectiondate")

        myrow = bs.Current
        Me.bs = bs
        'DateTimePicker1.Value = myrow.Item(0)
        'DateTimePicker2.Value = myrow.Item(1)


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Validate()
        'myrow.Item(0) = DateTimePicker1.Value
        'myrow.Item(1) = DateTimePicker2.Value
        bs.EndEdit()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        bs.CancelEdit()
    End Sub

    Private Sub FormInputCutoff_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        DateTimePicker1.DataBindings.Clear()
        DateTimePicker2.DataBindings.Clear()
        DateTimePicker3.DataBindings.Clear()
        DateTimePicker4.DataBindings.Clear()
    End Sub
End Class