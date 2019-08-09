Imports EPP.PublicClass
Public Class FormBillingTo
    Public Property EmployeeNumber As String
    Public Property Surname As String
    Public Property GivenName As String
    Dim DT As New DataTable
    Public Property bs As New BindingSource
    Dim userinfo As UserInfo
    Public Sub New(ByVal userinfo As UserInfo)
        InitializeComponent()
        Me.userinfo = userinfo
        DT.Columns.Add("employeenumber", GetType(String))
        DT.Columns.Add("surname", GetType(String))
        DT.Columns.Add("givenname", GetType(String))
        bs.DataSource = DT
        Dim dr As DataRowView = bs.AddNew
        dr.Row.Item("employeenumber") = userinfo.employeenumber
        dr.Row.Item("surname") = userinfo.Surname
        dr.Row.Item("givenname") = userinfo.GivenName
        DT.Rows.Add(dr.Row)

        TextBox1.DataBindings.Add("Text", DT, "employeenumber")
        TextBox2.DataBindings.Add("Text", DT, "surname")
        TextBox3.DataBindings.Add("Text", DT, "givenname")

        Me.EmployeeNumber = userinfo.employeenumber
        Me.Surname = userinfo.Surname
        Me.GivenName = userinfo.GivenName
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If validateTextBox() Then
            Dim mymessage As String = String.Empty
            If DbAdapter1.inputupdatepromoter(TextBox1.Text, TextBox2.Text, TextBox3.Text, userinfo.employeenumber, mymessage) Then
            Else
                MessageBox.Show(mymessage)
            End If
        End If

    End Sub

    Private Function validateTextBox() As Boolean
        Dim myret As Boolean = False
        If UserInfo.employeenumber <> TextBox1.Text Then
            Return True
        End If
        Return myret
    End Function

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim myobj = CType(sender, TextBox)
        If myobj.Name = "TextBox1" Then
            If myobj.Text = "" Then
                'TextBox1.Text = userinfo.employeenumber
                'TextBox2.Text = userinfo.Surname
                'TextBox3.Text = userinfo.GivenName
            ElseIf TextBox1.Text <> userinfo.employeenumber Then

                TextBox2.Text = ""
                TextBox3.Text = ""
            End If
        End If
    End Sub

    Private Sub TextBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Validated
        ErrorProvider1.SetError(sender, "")
    End Sub


    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        Dim ds As New DataSet
        Dim sqlstr As String = String.Empty
        If sender.text = "" Then
            ErrorProvider1.SetError(sender, "Value cannot be null.")
            e.Cancel = vbTrue
        ElseIf sender.text = userinfo.employeenumber Then
            TextBox2.Text = userinfo.Surname
            TextBox3.Text = userinfo.GivenName
        Else
            sqlstr = "select * from shop.employee where employeenumber = '" & TextBox1.Text & "'"
            Dim mymessage As String = String.Empty
            If DbAdapter1.GetDataSet(sqlstr, ds, mymessage) Then
                If ds.Tables(0).Rows.Count <> 0 Then
                    If Not IsDBNull(ds.Tables(0).Rows(0).Item("blockenddate")) Then
                        If ds.Tables(0).Rows(0).Item("blockenddate") > Today.Date Then
                            MessageBox.Show(String.Format("Sory, You cannot place an order for this person. Please come back again after {0:dd-MMM-yyyy}", ds.Tables(0).Rows(0).Item("blockenddate")))
                            e.Cancel = vbTrue
                            Exit Sub
                        End If

                    End If
                    Dim dr As DataRowView = bs.Current
                    dr.Row.Item("employeenumber") = ds.Tables(0).Rows(0).Item("employeenumber")
                    dr.Row.Item("surname") = ds.Tables(0).Rows(0).Item("sn")
                    dr.Row.Item("givenname") = ds.Tables(0).Rows(0).Item("givenname")
                    TextBox2.Text = ds.Tables(0).Rows(0).Item("sn")
                    TextBox3.Text = ds.Tables(0).Rows(0).Item("givenname")

                    'e.Cancel = vbTrue
                Else
                    MessageBox.Show("Please register to HR/Administrator for this Employee Number.")
                    e.Cancel = vbTrue
                End If
            Else
                MessageBox.Show(mymessage)
                e.Cancel = vbTrue
            End If
        End If
    End Sub


End Class