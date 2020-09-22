Imports System.Threading
Imports EPP.PublicClass
Imports EPP.SharedClass
Imports Microsoft.Office.Interop

Public Class FormImportAvailableQty
    Dim myThreadDelegate As New ThreadStart(AddressOf dowork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim myselectedPath As String


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ProgressBar1.Value = 0
        ProgressBar2.Value = 0
        Label1.Text = ""
        Label2.Text = ""

        If Not myThread.IsAlive Then
            'With FolderBrowserDialog1
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"

            If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                myselectedPath = OpenFileDialog1.FileName

                Try
                    myThread = New System.Threading.Thread(myThreadDelegate)
                    myThread.SetApartmentState(ApartmentState.MTA)
                    myThread.Start()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
            'End With
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Private Sub dowork()
        '*******************
        ' Get Related Data
        ' get cutoff order from system parameters
        ' Order Header
        ' Order Details
        Dim sqlstr As String = String.Empty
        Dim orderdtlBS As New BindingSource
        Dim orderHDBS As New BindingSource
        Dim sw As New Stopwatch
        sw.Start()

        'qtyconfirmation is sendemailqtyconfirmation
        sqlstr = "Select * from shop.pohd ph " &
                 " left join shop.status s on s.statusid = ph.status" &
                 " where statusname = 'New Order' and not qtyconfirmation and stamp <= shop.getcurrentcutoff();" &
                 " select ph.stamp as mystamp, pd.*,i.refno from shop.podtl pd " &
                 " left join shop.pohd ph on ph.pohdid = pd.pohdid" &
                 " left join shop.item i on i.itemid = pd.itemid" &
                 " left join shop.status s on s.statusid = ph.status" &
                 " where statusname = 'New Order' and not qtyconfirmation and ph.stamp <= shop.getcurrentcutoff();" &
                 " select * from shop.getcurrentcutoff(); " &
                 " select * from shop.cutoff c where c.cutoff = shop.getcurrentcutoff();"
        Dim DS As New DataSet
        Dim mymessage As String = String.Empty

        If DbAdapter1.GetDataSet(sqlstr, DS, mymessage) Then
            If DS.Tables(3).Rows.Count = 0 Then
                ProgressReport(11, "Current Cutoff Date is not available. Please Check the Cutoff Table and System Parameter.")
                Exit Sub
            End If
            'Do not remove the remarks. this one is correct. with this remark, it will allow to send the email again, even no New Order status in the order.
            'as long as there a Rejected and Accepted order with qtyconfirmation set to false
            'If DS.Tables(1).Rows.Count = 0 Then
            '    ProgressReport(11, "Done!! No new order to process.")
            '    Exit Sub
            'End If
            Dim idx(0) As DataColumn
            idx(0) = DS.Tables(0).Columns(0)
            DS.Tables(0).PrimaryKey = idx

            orderdtlBS.DataSource = DS.Tables(1)


        Else
            MessageBox.Show(mymessage)
            Exit Sub
        End If

        'Assign Qty
        ProgressReport(6, "Set Marque to Progressbar 1")
        ProgressReport(2, "in progres..")
        'Thread.Sleep(5000)
        'Load DataSet Order with status = New Order and stamp <= cutoff order
        'Set Dataset order with "Rejected status"
        'Read Excel

        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Dim myrow As Integer = 0

        '***********Logic **********************
        'For each Row in Excel
        '     availqty = excel availableqty
        '     get orders of item (Load Dataset Order Details of Dataset Order sort by stamp order)
        '     for each order of items
        '          order.confirmedqty = 0
        '          if availqty > 0 then
        '               
        '               'Find Order Header of this order detail. Set POHD Accepted    
        '               if availqty >= order.qty then
        '                   order.confirmedqty = order.qty
        '                   availqty = availqty - order.qty
        '               else
        '                   order.confirmedqty = availqty
        '                   availqty = 0
        '               endif
        '               
        '          endif
        '
        '     next
        'next

        Try
            'Create Object Excel 
            ProgressReport(11, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            oXl.Visible = False
            oXl.DisplayAlerts = False
            ProgressReport(11, "Opening Template...")
            oWb = oXl.Workbooks.Open(myselectedPath)
            ProgressReport(11, "")
            oXl.Visible = False
            oSheet = oWb.Worksheets(1)
            Dim orange = oSheet.Range("A1")
            Dim irows = GetLastRow(oXl, oSheet, orange)

            'Set OrderHD to Rejected first
            For Each dr As DataRow In DS.Tables(0).Rows
                dr.Item("status") = 3
            Next

            For myrow = 2 To irows
                Dim myrecord(2) As String
                For mycol = 1 To 3
                    If Not IsNothing(oSheet.Cells(myrow, mycol).value) Then
                        Dim myvalue = oSheet.Cells(myrow, mycol).value.ToString
                        myrecord(mycol - 1) = myvalue
                    End If
                Next
                Dim availqty = myrecord(2)
                Dim productid = myrecord(0)

                'order dtl filter
                orderdtlBS.Filter = "refno='" & myrecord(0) & "'"
                orderdtlBS.Sort = "mystamp asc"


                If availqty > 0 Then
                    For Each drv As DataRowView In orderdtlBS.List
                        'Find Order Header of this order detail. Set POHD Accepted    
                        Dim dr As DataRow = drv.Row
                        Dim mykey(0) As Object
                        mykey(0) = dr.Item("pohdid")
                        Dim myresult = DS.Tables(0).Rows.Find(mykey)
                        If availqty > 0 Then
                            If Not IsNothing(myresult) Then
                                myresult.Item("status") = 2 'Accepted'
                            End If
                        End If
                        If availqty >= dr.Item("qty") Then
                            dr.Item("confirmedqty") = dr.Item("qty")
                            availqty = availqty - dr.Item("qty")                           
                        ElseIf availqty > 0 Then
                            dr.Item("confirmedqty") = availqty
                            availqty = 0                            
                        End If
                    Next
                End If                
            Next

            Dim ds3 = DS.GetChanges
            If Not IsNothing(ds3) Then
                Dim ra As Integer
                'Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                Dim mye As New ContentBaseEventArgs(ds3, True, mymessage, ra, True)
                If DbAdapter1.UpdatePO(Me, mye) Then
                    Dim myquery = From row As DataRow In DS.Tables(0).Rows
                                  Where row.RowState = DataRowState.Added

                    For Each rec In myquery.ToArray
                        rec.Delete()
                    Next
                    DS.Merge(ds3)
                    DS.AcceptChanges()
                    'BS.Position = myposition
                    'MessageBox.Show("Saved!")

                    'LoadData()
                Else
                    MessageBox.Show(mye.message)
                End If
                ' MessageBox.Show("hello")
            End If

        Catch ex As Exception
        Finally
            Try
                oXl.Quit()
                releaseComObject(oSheet)
                releaseComObject(oWb)
                releaseComObject(oXl)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception

            End Try

            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try
        End Try

        'Update Dataset Order and Order Details
        'Thread.Sleep(5000)
        ProgressReport(4, "Set Continuous to Progressbar 1")
        ProgressReport(2, "Done")


        '************************
        'Send Email rejected or accepted order
        ProgressReport(7, "Set Marque to Progressbar 2")
        ProgressReport(3, "in progres..")

        'For each row in Orders where qtyconfirmation = false and status rejected or accepted
        '   'send email
        '   'if success then
        '       assign confirmationqty = true
        '    else
        '       run send email qtyconfirmation manually for later until no qtyconfirmation = true
        '    endif
        'next

        sqlstr = "Select ph.*,e.sn || ' ' || e.givenname as employeename,b.employeenumber as billingto, b.sn || ' ' || b.givenname as billingtoname,e.mail from shop.pohd ph " &
                 " left join shop.status s on s.statusid = ph.status" &
                 " left join shop.employee e on e.employeenumber =  ph.employeenumber" &
                 " left join shop.employee b on b.employeenumber =  ph.billingto" &
                 " where qtyconfirmation = false and (status = 2 or status = 3) and stamp <= shop.getcurrentcutoff();" &
                 " select ph.stamp as mystamp, pd.*,i.refno,d.descriptionname from shop.podtl pd " &
                 " left join shop.pohd ph on ph.pohdid = pd.pohdid" &
                 " left join shop.item i on i.itemid = pd.itemid" &
                 " left join shop.description d on d.descriptionid = i.descriptionid" &
                 " left join shop.status s on s.statusid = ph.status" &
                 " where qtyconfirmation = false and (status = 2 or status = 3) and ph.stamp <= shop.getcurrentcutoff();" &
                 " select * from shop.paramhd ph where ph.paramname = 'Administrator email'"


        Dim DS2 = New DataSet
        mymessage = String.Empty

        If DbAdapter1.GetDataSet(sqlstr, DS2, mymessage) Then
            'If DS.Tables(3).Rows.Count = 0 Then
            '    ProgressReport(11, "Next Cutoff Date is not available. Please Check the Cutoff Table.")
            '    Exit Sub
            'End If
            Dim idx(0) As DataColumn
            idx(0) = DS2.Tables(0).Columns(0)
            DS2.Tables(0).PrimaryKey = idx
            orderdtlBS.DataSource = DS2.Tables(1)

            Dim relation As New DataRelation("relHDDTL", DS2.Tables(0).Columns("pohdid"), DS2.Tables(1).Columns("pohdid"))
            DS2.Relations.Add(relation)

        Else
            MessageBox.Show(mymessage)
            Exit Sub
        End If



        For Each dr As DataRow In DS2.Tables(0).Rows
            Dim myOrder As Object
            Dim ishtml As Boolean = True
            If dr.Item("status") = 2 Then
                myOrder = New AcceptedOrder(dr, DS.Tables(3).Rows(0))
                ishtml = True
            Else
                myOrder = New RejectedOrder(dr)
            End If
            'Debug.Print(myOrder.bodymessage())
            Dim myemail As New Email
            myemail.sendto = dr.Item("mail")
            myemail.sender = DS2.Tables(2).Rows(0).Item("cvalue")
            myemail.body = myOrder.bodymessage
            myemail.subject = myOrder.subject
            myemail.isBodyHtml = True
            'send email and flag qtyconfirmation to true
            If myemail.send(mymessage) Then
                'flag sent to hd 1 by 1, 
                sqlstr = "update shop.pohd ph set qtyconfirmation = true where ph.pohdid = " & dr.Item("pohdid")
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, message:=mymessage) Then
                    MessageBox.Show(mymessage)
                End If
            End If
        Next

        'Thread.Sleep(5000)

        ProgressReport(5, "Set Continuous to Progressbar 2")
        ProgressReport(3, "Done")
        ProgressReport(13, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 2
                    Label1.Text = message
                Case 3
                    Label2.Text = message
                Case 4
                    ProgressBar1.Style = ProgressBarStyle.Continuous
                Case 5
                    ProgressBar2.Style = ProgressBarStyle.Continuous
                Case 6
                    ProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    ProgressBar2.Style = ProgressBarStyle.Marquee                
                Case 11
                    ToolStripStatusLabel1.Text = message
                Case 12
                    Dim myvalue = message.ToString.Split(",")
                    ProgressBar1.Minimum = 1
                    ProgressBar1.Value = myvalue(0)
                    ProgressBar1.Maximum = myvalue(1)
                Case 13
                    ToolStripStatusLabel1.Text = message
            End Select
        End If
    End Sub
End Class