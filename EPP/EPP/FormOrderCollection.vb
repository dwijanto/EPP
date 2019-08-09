Imports System.Threading
Imports EPP.PublicClass
Public Class FormOrderCollection

    Dim DS As DataSet
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim currentcutoff As Integer
    Dim SendToAllEmployee As String
    Dim sendManually As Boolean = False
    Dim bs As BindingSource

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not myThread.IsAlive Then
            sendManually = True
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
        ProgressReport(2, "")
        Dim mymessage As String = String.Empty

        Dim sqlstr = "Select s.* from shop.paramhd  p left join shop.cutoff s on s.cutoffid = p.nvalue where p.paramname = 'Current Cutoff';" &
                 " Select ph.*,e.sn || ' ' || e.givenname as employeename,b.employeenumber as billingto, b.sn || ' ' || b.givenname as billingtoname,e.mail from shop.pohd ph " &
                 " left join shop.status s on s.statusid = ph.status" &
                 " left join shop.employee e on e.employeenumber =  ph.employeenumber" &
                 " left join shop.employee b on b.employeenumber =  ph.billingto" &
                 " where qtyconfirmation = true and status = 4;" &
                 " select ph.stamp as mystamp, pd.*,i.refno,d.descriptionname from shop.podtl pd " &
                 " left join shop.pohd ph on ph.pohdid = pd.pohdid" &
                 " left join shop.item i on i.itemid = pd.itemid" &
                 " left join shop.description d on d.descriptionid = i.descriptionid" &
                 " where qtyconfirmation = true and status = 4; " &
                 " select * from shop.paramhd ph where ph.paramname = 'Administrator email';" &
                 " select * from shop.cutoff c where c.cutoff = shop.getcurrentcutoff();"


        sqlstr = "Select s.* from shop.paramhd  p left join shop.cutoff s on s.cutoffid = p.nvalue where p.paramname = 'Current Cutoff';" &
                 " Select ph.*,e.sn || ' ' || e.givenname as employeename,b.employeenumber as billingto, b.sn || ' ' || b.givenname as billingtoname,e.mail from shop.pohd ph " &
                 " left join shop.status s on s.statusid = ph.status" &
                 " left join shop.employee e on e.employeenumber =  ph.employeenumber" &
                 " left join shop.employee b on b.employeenumber =  ph.billingto" &
                 " left join shop.paramhd p on p.paramname = 'Current Cutoff'" &
                 " where qtyconfirmation = true and status = 4 and ph.cutoffid = p.nvalue;" &
                 " select ph.stamp as mystamp,ph.cutoffid, pd.*,i.refno,d.descriptionname from shop.podtl pd " &
                 " left join shop.pohd ph on ph.pohdid = pd.pohdid" &
                 " left join shop.item i on i.itemid = pd.itemid" &
                 " left join shop.description d on d.descriptionid = i.descriptionid" &
                 " left join shop.paramhd p on p.paramname = 'Current Cutoff'" &
                 " where qtyconfirmation = true and status = 4 and ph.cutoffid = p.nvalue; " &
                 " select * from shop.paramhd ph where ph.paramname = 'Administrator email';" &
                 " select * from shop.cutoff c where c.cutoff = shop.getcurrentcutoff();"

        DS = New DataSet
        If Not DbAdapter1.GetDataSet(sqlstr, DS, mymessage) Then
            ProgressReport(2, mymessage)
            Logger.log(mymessage)
            Exit Sub
        End If

        Dim relation As DataRelation
        relation = New DataRelation("relHDDTL", DS.Tables(1).Columns("pohdid"), DS.Tables(2).Columns("pohdid"))
        DS.Relations.Add(relation)

        bs = New BindingSource
        bs.DataSource = DS.Tables(1)
        If Not sendManually Then
            bs.Filter = "pickupitem = False"
        End If
        If bs.List.Count > 0 Then
            Dim emailsent As Integer = 0
            For Each drv As DataRowView In bs.List
                Dim dr As DataRow = drv.Row

                Dim myOrder As Object
                myOrder = New PickupItem(dr, DS.Tables(4).Rows(0))

                'Debug.Print(myOrder.bodymessage())
                Dim myemail As New Email
                myemail.sendto = dr.Item("mail")
                myemail.sender = DS.Tables(3).Rows(0).Item("cvalue")
                myemail.body = myOrder.bodymessage
                myemail.subject = myOrder.subject
                myemail.isBodyHtml = True
                'send email and flag qtyconfirmation to true

                If myemail.send(mymessage) Then
                    'flag sent to hd 1 by 1, 
                    emailsent += 1
                    sqlstr = "update shop.pohd ph set pickupitem = true where ph.pohdid = " & dr.Item("pohdid")
                    If Not DbAdapter1.ExecuteNonQuery(sqlstr, message:=mymessage) Then
                        ProgressReport(2, mymessage)
                        Logger.log(mymessage)
                    End If
                End If
            Next
            ProgressReport(2, "Done.")
            Logger.log("Done. Email sent(" & emailsent & ")")
        Else
            ProgressReport(5, "Marquee")
            ProgressReport(2, "Nothing to send.")
            Logger.log("Nothing to send.")
        End If
        
        ProgressReport(5, "Marquee")
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