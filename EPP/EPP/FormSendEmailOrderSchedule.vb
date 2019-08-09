Imports System.Threading
Imports EPP.PublicClass
Public Class FormSendEmailOrderScheduled
    Dim ParamDS As DataSet
    Dim CutoffDS As DataSet
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim currentcutoff As Integer
    Dim SendToAllEmployee As String
    Dim SendOrderSchedule As String
    
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
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

    Sub DoWork()
        ProgressReport(6, "Marquee")
        Dim mymessage As String = String.Empty
        Dim sqlstr = "select * from shop.paramhd ph  where ph.parent = 1 order by ivalue;"

        ParamDS = New DataSet
        If DbAdapter1.GetDataSet(sqlstr, ParamDS, mymessage) Then
            ParamDS.Tables(0).TableName = "Param"
            Dim idx(0) As DataColumn
            idx(0) = ParamDS.Tables(0).Columns("paramhdid")
            ParamDS.Tables(0).PrimaryKey = idx
            'currentcutoff = ParamDS.Tables(0).Rows(0).Item("tvalue")
            currentcutoff = ParamDS.Tables(0).Rows(0).Item("nvalue")
            SendToAllEmployee = ParamDS.Tables(0).Rows(5).Item("cvalue")
            SendOrderSchedule = ParamDS.Tables(0).Rows(11).Item("cvalue")

        Else
            MessageBox.Show(mymessage)
            Exit Sub
        End If
        CutoffDS = New DataSet
        'sqlstr = String.Format("select * from shop.cutoff where cutoff = '{0:yyyy-MM-dd HH:mm:ss}'", currentcutoff)
        sqlstr = String.Format("select * from shop.cutoff where cutoffid = '{0}'", currentcutoff)
        Try
            If DbAdapter1.GetDataSet(sqlstr, CutoffDS, mymessage) Then
                Dim cutoffdate As DateTime = CutoffDS.Tables(0).Rows(0).Item("cutoff")
                Dim collectiondate As Date = CutoffDS.Tables(0).Rows(0).Item("collectiondate")
                Dim notificationemail As Date = CutoffDS.Tables(0).Rows(0).Item("emailnotificationdate")
                Dim chequesubmit As Date = CutoffDS.Tables(0).Rows(0).Item("chequesubmitdate")

                'Dim mymail = New SendEmailOrderSchedule(cutoffdate, collectiondate, notificationemail, chequesubmit, "no-reply@groupeseb.com", SendToAllEmployee)
                Dim mymail = New SendEmailOrderSchedule(cutoffdate, collectiondate, notificationemail, chequesubmit, SendOrderSchedule, SendToAllEmployee)
                mymail.isBodyHtml = True
                If mymail.send(mymessage) Then
                    ProgressReport(5, "Marquee")
                    MessageBox.Show("Sent.")
                Else
                    ProgressReport(5, "Marquee")
                    MessageBox.Show(mymessage)
                End If
            Else
                ProgressReport(5, "Marquee")
                MessageBox.Show(mymessage)
            End If
        Catch ex As Exception
            ProgressReport(5, "Marquee")
            MessageBox.Show(ex.Message)
        End Try

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
            End Select
        End If
    End Sub
End Class