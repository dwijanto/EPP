Imports System.Threading
Imports EPP.PublicClass
Imports System.IO
Imports Microsoft.Office.Interop
Imports DJLib.Dbtools
Imports System.Text

Public Class FormSendQtyToGSHK
    Dim currentcutoff As DateTime
    Dim GSHKEmail As String
    Dim AdminEmail As String
    Public Property sendemailtologistic As Boolean
    Public Property mymessage As String = String.Empty
    Dim ds As New DataSet
    Dim ParamDS As New DataSet
    Public Property StatusSent As Boolean


    Private Sub FormSendQtyToGSHK_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Visible = False
        'mymessage = "Start Cutoff"
        Logger.log("Start Cutoff")
        dowork()
        Logger.log("Finish Cutoff")
        Me.Close()
    End Sub

    Sub dowork()
        'get currentcutoff
        'Dim sqlstr = "select tvalue from shop.paramhd where paramname = 'Current Cutoff';"
        'Dim mymessage As String = String.Empty
        If IsNothing(DbAdapter1) Then
            DbAdapter1 = New DbAdapter
            dbtools1.Userid = DbAdapter1.userid
            dbtools1.Password = DbAdapter1.password
        End If
        'If Not DbAdapter1.ExecuteScalar(sqlstr, currentcutoff, mymessage) Then
        '    'send email to admin
        '    Exit Sub
        'End If
        'Dim mymessage As String = String.Empty
        Dim sqlstr = "select * from shop.paramhd ph  where ph.parent = 1 order by ivalue;" &
                     "select shop.getcurrentcutoff() as cutoff;"
        If DbAdapter1.GetDataSet(sqlstr, ParamDS, mymessage) Then
            ParamDS.Tables(0).TableName = "Param"
            Dim idx(0) As DataColumn
            idx(0) = ParamDS.Tables(0).Columns("paramhdid")
            ParamDS.Tables(0).PrimaryKey = idx

            currentcutoff = ParamDS.Tables(1).Rows(0).Item("cutoff")
            GSHKEmail = "" & ParamDS.Tables(0).Rows(2).Item("cvalue")
            AdminEmail = "" & ParamDS.Tables(0).Rows(3).Item("cvalue")
            sendemailtologistic = ParamDS.Tables(0).Rows(4).Item("bvalue")
        Else
            Logger.log(mymessage)
        End If
        'compare with date today
        'if date today <= current cutoff then proceed
        '    docutoff
        'else
        '  just quit
        'endif

        'This process will run if today's date is bigger than currentcutoff
        If currentcutoff <= Date.Now Then
            If Not sendemailtologistic Then
                docutoff()
            Else
                mymessage = "Email was not sent. If you want to send this Email, Please Update ""Email Sent Status To Logistics"" in System Parameters, set the value to False."
                Logger.log(mymessage)
            End If

        Else
            mymessage = "Not yet cutoff."
            Logger.log(mymessage)
            '
            'MessageBox.Show("not cutoff yet")
        End If


    End Sub

    Private Sub docutoff()
        'get from po with status = 'New Order' and orderdate <= currentcutoff

        'Add checking for cutoffid, only transaction 
        Dim sqlstr = "select refno,sum(qty) as qty ,'' as ""available qty"" from shop.podtl pd" &
                     " left join shop.item i on i.itemid = pd.itemid" &
                     " left join shop.pohd ph on ph.pohdid = pd.pohdid" &
                     " left join shop.paramhd p on p.paramname = 'Current Cutoff'" &
                     " where ph.status = 1 and ph.stamp <= '" & strdatetime(currentcutoff) & "' and ph.cutoffid = p.nvalue" &
                     " group by refno" &
                     " order by refno"
        'Dim mymessage As String = String.Empty
        If DbAdapter1.GetDataSet(sqlstr, ds, mymessage) Then
            ds.Tables(0).TableName = "PO"
            If ds.Tables(0).Rows.Count > 0 Then
                'create attachment
                'MessageBox.Show("Create Attachment")
                Dim filename As String = String.Empty
                For Each myfile In Directory.EnumerateFiles(Application.StartupPath & "\Attachment")
                    Try
                        File.Delete(myfile)
                    Catch ex As Exception

                    End Try

                Next
                If CreateAttachment(filename, sqlstr) Then
                    Dim attachmentlist As New List(Of String)
                    attachmentlist.Add(filename)
                    'Dim mymail As New Email With {.sender = Adminemail,
                    '                          .sendto = gshkemail,
                    '                          .subject = "Please check available quantity for these items",
                    '                          .cc = adminemail,
                    '                          .body = "Send email using smtp",
                    '                           .attachmentlist = attachmentlist}
                    Dim body As New StringBuilder
                    body.Append("Dear GSHK Logistics Team," & vbCrLf & vbCrLf)
                    body.Append("Please find attached file." & vbCrLf & vbCrLf)
                    body.Append("Best Regards," & vbCrLf & vbCrLf)
                    body.Append("e-Staff Purchase Administrator")
                    Dim mymail As New Email With {.sender = AdminEmail,
                                              .sendto = GSHKEmail,
                                              .subject = "Please check available quantity for these items",
                                              .cc = AdminEmail,
                                              .body = body.ToString,
                                               .attachmentlist = attachmentlist}

                    If mymail.send(mymessage) Then
                        Logger.log("Send Mail success")
                        StatusSent = True
                        'Flag sent this cutoff
                        'sendemailtologistic = True
                        sqlstr = "update shop.paramhd set bvalue = true where paramname = 'SendEmailToLogistics'"
                        If Not DbAdapter1.ExecuteScalar(sqlstr, message:=mymessage) Then
                            Logger.log(mymessage)
                        End If
                        File.Delete(filename)
                        'MessageBox.Show("Sent.")
                    Else
                        mymessage = "Send Mail Failed"
                        Logger.log(mymessage)
                    End If
                Else
                    mymessage = "Create Attachment Failed"
                    Logger.log(mymessage)
                End If

                'send email
                'done
            Else
                mymessage = "No record to process"
                Logger.log(mymessage)
                'No record to process
                Exit Sub
            End If
        Else
            Logger.log(mymessage)
            'Error happened
            Exit Sub
        End If
        'create attachment
        'send email
        'done
    End Sub

    Private Function strdatetime(ByVal currentcutoff As Date) As String
        Dim myret = ""
        Try
            myret = String.Format("{0:yyyy-MM-dd HH:mm:ss}", currentcutoff)
        Catch ex As Exception
        End Try
        Return myret
    End Function

    Function CreateAttachment(ByRef filename As String, ByVal sqlstr As String) As Boolean
        Dim myQueryWorksheetList As New List(Of QueryWorksheet)
        Dim MyDirectory = Application.StartupPath & "\Attachment"

        Dim ReportName = "Orders"
        Dim myqueryworksheet = New QueryWorksheet With {.DataSheet = 1,
                                                        .SheetName = ReportName,
                                                        .Sqlstr = sqlstr}
        myQueryWorksheetList.Add(myqueryworksheet)
        Dim myreport As New ExportToExcelFile(Me, myQueryWorksheetList, MyDirectory, ReportName)

        Return myreport.CreateExcel(filename)
    End Function



End Class