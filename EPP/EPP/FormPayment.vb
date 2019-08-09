Imports EPP.PublicClass

Public Class FormPayment
    Dim bs As BindingSource
    Dim dtlBS As BindingSource
    Dim DS As New DataSet
    Dim dr As DataRow
    Public Sub New(ByRef bs As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()
        Me.bs = bs
        dr = CType(bs.Current, DataRowView).Row
        ' Add any initialization after the InitializeComponent() call.

    End Sub




    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Validate()
        'dtlBS.Current.row.item("chequenumber") = TextBox1.Text
        dtlBS.EndEdit()

        Dim ds2 = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            'Dim docnumber As String = String.Empty
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If DbAdapter1.paymenttx(Me, mye) Then
                Dim myquery = From row As DataRow In DS.Tables(0).Rows
                              Where row.RowState = DataRowState.Added

                For Each rec In myquery.ToArray
                    rec.Delete()
                Next
                DS.Merge(ds2)
                DS.AcceptChanges()
                dr.Item("chequenumber") = ds2.Tables(0).Rows(0).Item("chequenumber")
                dr.Item("bankcode") = ds2.Tables(0).Rows(0).Item("bankcode")
                'MessageBox.Show("Saved!")
            Else
                MessageBox.Show(mye.message)
                Me.DialogResult = DialogResult.Cancel
                Exit Sub
            End If
        Else
            MessageBox.Show("Nothing to save.")
        End If

        MessageBox.Show("Thank you, Please send your cheque to Administration.")
    End Sub

    Private Sub FormPayment_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Private Sub loaddata()
        Dim sqlstr = String.Empty
        sqlstr = "select * from shop.payment where paymentid = " & dr.Item("pohdid")
        Dim mymessage As String = String.Empty
        If DbAdapter1.GetDataSet(sqlstr, DS, mymessage) Then
            dtlBS = New BindingSource
            dtlBS.DataSource = DS.Tables(0)


            'if not avail row then add row
            DS.Tables(0).TableName = "Payment"
            Dim drv As DataRowView
            Dim paydr As DataRow

            If DS.Tables(0).Rows.Count = 0 Then
                drv = dtlBS.AddNew()
                paydr = drv.Row
                DS.Tables(0).Rows.Add(paydr)
            Else
                drv = dtlBS.Current
                paydr = drv.Row
            End If

            TextBox1.DataBindings.Clear()
            TextBox2.DataBindings.Clear()
            'TextBox3.DataBindings.Clear()

            TextBox1.DataBindings.Add("Text", dtlBS, "chequenumber")
            TextBox2.DataBindings.Add("Text", dtlBS, "bankcode")
            'TextBox3.DataBindings.Add("Text", dtlBS, "totalamount")

            TextBox1.Text = "" & paydr.Item("chequenumber")
            TextBox2.Text = "" & paydr.Item("bankcode")
            TextBox3.Text = dr.Item("totalamount")


            paydr.Item("paymentid") = dr.Item("pohdid")
            paydr.Item("chequenumber") = dr.Item("chequenumber")
            paydr.Item("bankcode") = dr.Item("bankcode")
            paydr.Item("amount") = dr.Item("totalamount")

            Label5.Text = dr.Item("billrefno")
        Else
            MessageBox.Show(mymessage)
        End If
    End Sub

    Private Sub TextBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Validated, TextBox2.Validated
        ErrorProvider1.SetError(sender, "")
    End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating, TextBox2.Validating
        'If sender.text = "" Then
        '    ErrorProvider1.SetError(sender, "Value cannot be blank.")
        '    e.Cancel = True
        'End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'set causesvalidation = false in control if  you want to by pass text validation
        dtlBS.CancelEdit()
        Me.Close()
    End Sub
End Class