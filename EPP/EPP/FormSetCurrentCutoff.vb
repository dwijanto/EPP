Imports System.Threading
Imports EPP.PublicClass
Public Class FormSetCurrentCutoff

    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim ds As DataSet
    Dim bs As BindingSource

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sqlstr = "update shop.paramhd ph set nvalue = " & ComboBox1.SelectedValue & " where ph.paramname = 'Current Cutoff';" &
                     "update shop.paramhd ph set bvalue = false where ph.paramname = 'SendEmailToLogistics'"
        Dim mymessage As String = String.Empty
        If DbAdapter1.ExecuteScalar(sqlstr, message:=mymessage) Then
            TextBox1.Text = ComboBox1.Text
            ProgressReport(2, "Done.")
        Else
            MessageBox.Show(mymessage)
        End If
    End Sub

    Sub DoWork()
        ProgressReport(2, "Populating Data. Please wait..")
        Dim sqlstr = "Select c.*,cutoff::character varying as mycutoff  from  shop.cutoff c where date_part('Year',c.cutoff) >= " & Date.Today.Year & " order by cutoff desc;" &
                     " select ph.*,c.cutoff::character varying as mycutoff from shop.paramhd ph " &
                     " left join shop.cutoff c on c.cutoffid = ph.nvalue where ph.paramname = 'Current Cutoff';"

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
                    ComboBox1.DataSource = bs
                    ComboBox1.DisplayMember = "mycutoff"
                    ComboBox1.ValueMember = "cutoffid"

                    ComboBox1.SelectedValue = ds.Tables(1).Rows(0).Item("nvalue")
                    TextBox1.Text = ds.Tables(1).Rows(0).Item("mycutoff")
            End Select
        End If
    End Sub

    Private Sub FormSetCurrentCutoff_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not myThread.IsAlive Then
            Try
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
End Class