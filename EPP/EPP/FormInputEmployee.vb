Public Class FormInputEmployee
    Private myrow As DataRowView
    Private bs As BindingSource
    Public Sub New(ByRef bs As BindingSource, ByVal addnew As Boolean)
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        TextBox1.DataBindings.Add("Text", bs, "employeenumber")
        TextBox2.DataBindings.Add("Text", bs, "sn")
        TextBox3.DataBindings.Add("Text", bs, "givenname")
        TextBox4.DataBindings.Add("Text", bs, "department")
        TextBox5.DataBindings.Add("Text", bs, "telephonenumber")
        TextBox6.DataBindings.Add("Text", bs, "mail")
        TextBox7.DataBindings.Add("Text", bs, "samaaccountname")
        TextBox8.DataBindings.Add("Text", bs, "manager")
        DateTimePicker1.DataBindings.Add("Text", bs, "blockenddate")
        myrow = bs.Current
        Me.bs = bs

        TextBox1.Enabled = addnew


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Validate()
        If Not DateTimePicker1.Checked Then
            myrow.Row.Item("blockenddate") = DBNull.Value
        End If
        bs.EndEdit()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        bs.CancelEdit()
    End Sub

    Private Sub FormInputEmployee_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class