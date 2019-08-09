Imports System.Threading
Imports EPP.PublicClass
Public Class FormChequePayment
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim ds As DataSet
    Dim bs As BindingSource
    Dim mycriteria As String = String.Empty

    Sub DoWork()
        ProgressReport(2, "Populating Data. Please wait..")
        Dim sqlstr = "select ph.pohdid,billrefno,orderdate,billingto,e.sn || ' ' || e.givenname as billingtoname, e.sn,e.givenname,shop.getbillingamount(amount,ph.pohdid) as totalamount,chequenumber,bankcode,statusname from shop.pohd ph" &
                     " left join shop.payment py on py.paymentid = ph.pohdid" &
                     " left join shop.employee e on e.employeenumber = ph.billingto " &
                     " left join shop.status s on s.statusid = ph.status" & mycriteria

        Dim message As String = String.Empty
        ds = New DataSet
        If DbAdapter1.GetDataSet(sqlstr, ds, message) Then
            ProgressReport(8, "Init Data")
            ProgressReport(2, "Populating Data. Done.")
        Else
            ProgressReport(2, message)
        End If

    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 2
                    ToolStripStatusLabel1.Text = message
                Case 3
                    ToolStripStatusLabel2.Text = message
                Case 4
                    ToolStripStatusLabel1.Text = message
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ToolStripProgressBar1.Minimum = 1
                    ToolStripProgressBar1.Value = myvalue(0)
                    ToolStripProgressBar1.Maximum = myvalue(1)
                Case 8
                    bs = New BindingSource
                    bs.DataSource = ds.Tables(0)
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = bs
                    Dim mysubtotal As Decimal = 0
                    For Each drv As DataRowView In bs.List
                        mysubtotal += drv.Row.Item("totalamount")
                    Next
                    ToolStripStatusLabel2.Text = "Total Amount : " & String.Format("{0:#,##0.00}", mysubtotal)
            End Select
        End If
    End Sub

    Private Sub FormSetCurrentCutoff_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ToolStripComboBox1.SelectedIndex = 0
        If IsNothing(DbAdapter1) Then
            DbAdapter1 = New DbAdapter
            dbtools1.Userid = DbAdapter1.userid
            dbtools1.Password = DbAdapter1.password
            HelperClass1 = New HelperClass
        End If
        Try
            loglogin(DbAdapter1.userid)
        Catch ex As Exception
        End Try

    End Sub
    Private Sub loglogin(ByVal userid As String)
        Dim applicationname As String = "e-Staff Cheque Validation"
        Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        Dim computername As String = My.Computer.Name
        Dim time_stamp As DateTime = Now
        DbAdapter1.loglogin(applicationname, userid, username, computername, time_stamp)
    End Sub
    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        If Not myThread.IsAlive Then
            Try

                Select Case ToolStripComboBox1.SelectedIndex
                    Case 0
                        mycriteria = " where lower(chequenumber) = "
                    Case 1
                        mycriteria = " where lower(billrefno) = "
                    Case 2
                        mycriteria = " where lower(billingto) = "
                    Case 3
                        mycriteria = " where lower(e.sn) = "
                    Case 4
                        mycriteria = " where lower(e.givenname) = "
                End Select
                mycriteria = mycriteria & "'" & ToolStripTextBox1.Text.ToLower.Replace("'", "''") & "'"

                myThread = New System.Threading.Thread(myThreadDelegate)
                myThread.SetApartmentState(ApartmentState.MTA)
                myThread.Start()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim mymessage As String = String.Empty
        If e.ColumnIndex = 8 Then
            Dim mydr As DataRow = CType(bs.Current, DataRowView).Row
            If mydr.Item("statusname") = "Completed" Then
                MessageBox.Show("Order is completed. You cannot modify the payment information.")
                Exit Sub
            ElseIf mydr.Item("statusname") = "Rejected Order" Then
                MessageBox.Show("Order is Rejected. You cannot pay for rejected order.")
                Exit Sub
            ElseIf mydr.Item("statusname") = "Paid" Then
                MessageBox.Show("Order has been paid!.")
                Exit Sub
            ElseIf mydr.Item("statusname") = "New Order" Then
                MessageBox.Show("Payment can be done after you received order confirmation.")
                Exit Sub
            ElseIf mydr.Item("statusname") = "Accepted Order" Then
            Else
                MessageBox.Show("Only Accepted Order allowed.")
                Exit Sub
            End If
            Dim mycheck As New FormPayment(bs)
            If mycheck.ShowDialog() = Windows.Forms.DialogResult.OK Then
                'Update Status Paid
                'If Not setpaid(mymessage) Then
                ' MessageBox.Show(mymessage)
                'End If
                ds.AcceptChanges()
            End If
        ElseIf e.ColumnIndex = 9 Then
            'Update Status Paid
            Dim mydr As DataRow = CType(bs.Current, DataRowView).Row
            If IsDBNull(mydr.Item("chequenumber")) Or IsDBNull(mydr.Item("bankcode")) Then
                MessageBox.Show("Cheque information is blank! Cannot change status.")
                Exit Sub
            End If
            If mydr.Item("statusname") <> "Accepted Order" Then
                MessageBox.Show("Only Accepted Order allowed!")
                Exit Sub
            End If
            If MessageBox.Show("Do you want to change Status Order?", "Change Status", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                If Not setpaid(mymessage) Then
                    MessageBox.Show(mymessage)
                End If
                ds.AcceptChanges()
            End If
        ElseIf e.ColumnIndex = 10 Then
            Dim mydr As DataRow = CType(bs.Current, DataRowView).Row
            If mydr.Item("statusname") <> "New Order" Then
                MessageBox.Show("Only New Order allowed!")
                Exit Sub
            End If
            If MessageBox.Show("Do you want to cancel the Order?", "Cancel Order", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                If Not setCancel(mymessage) Then
                    MessageBox.Show(mymessage)
                End If
                ds.AcceptChanges()
            End If
        End If
    End Sub

    Function setpaid(ByRef mymessage As String) As Boolean
        Dim myret As Boolean = False
        Dim drv As DataRowView = bs.Current
        Dim dr As DataRow = drv.Row
        If dr.Item("statusname") = "Accepted Order" Then
            dr.Item("statusname") = "Paid"
            'ds.Tables(0).AcceptChanges()
            'Update Database
            Dim sqlstr = "update shop.pohd set status = 4 where pohdid = " & dr.Item("pohdid")
            Dim ra As Integer
            If DbAdapter1.ExecuteNonQuery(sqlstr, ra, mymessage) Then
                myret = True
            End If

        Else
            mymessage = "Only Accepted Order allowed."
        End If
        Return myret
    End Function

    Function setCancel(ByRef mymessage As String) As Boolean
        Dim myret As Boolean = False
        Dim drv As DataRowView = bs.Current
        Dim dr As DataRow = drv.Row
        If dr.Item("statusname") = "New Order" Then
            dr.Item("statusname") = "Canceled"
            'ds.Tables(0).AcceptChanges()
            'Update Database
            Dim sqlstr = "update shop.pohd set status = 6 where pohdid = " & dr.Item("pohdid")
            Dim ra As Integer
            If DbAdapter1.ExecuteNonQuery(sqlstr, ra, mymessage) Then
                myret = True
            End If

        Else
            mymessage = "Only New Order allowed."
        End If
        Return myret
    End Function

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Dim myform As New FormTxReport
        myform.ShowDialog()
        myform.Dispose()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Dim myform As New FormChequeDepositReport
        myform.ShowDialog()
        myform.Dispose()

    End Sub

    Private Sub ToolStripTextBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.Click

    End Sub

    Private Sub ToolStripTextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ToolStripTextBox1.KeyUp

        If e.KeyCode = Keys.Enter Then
            'Dim oritext = ComboBox1.Text
            'For i = 0 To ComboBox1.Items.Count - 1
            '    If ComboBox1.Items(i).row.item("refno").ToString.ToUpper = oritext.ToUpper Then
            '        ComboBox1.SelectedValue = ComboBox1.Items(i).row.item("itemid")

            '    End If
            'Next

            ToolStripButton1.PerformClick()
        End If

    End Sub
End Class