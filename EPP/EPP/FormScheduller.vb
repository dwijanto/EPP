Imports System.Threading.Thread
Imports EPP.PublicClass
Public Class FormScheduler
    Dim ds As DataSet
    Private Sub FormScheduller_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'get currentcutoff
        getcurrentcutoff()

        Logger.log("Start Scheduller")
        If Not IsNothing(ds) Then
            If ds.Tables(0).Rows(0).Item("cutoff") <= Date.Now Then
                cutoffscheduller()
            End If
            If ds.Tables(0).Rows(0).Item("chequesubmitdate") <= Date.Today.Date Then
                reminderscheduller()
            End If
            Dim mydate As Date = ds.Tables(0).Rows(0).Item("collectiondate")
            If mydate.AddDays(-1) <= Date.Today.Date Then
                orderCollectionscheduller()
            End If
        End If
        Me.Close()
        Logger.log("End Scheduller")
    End Sub

    Private Sub cutoffscheduller()
        Logger.log("Start Cutoff")
        Dim myform = New FormSendQtyToGSHK
        'myform.Show()
        myform.dowork()
        Logger.log("End Cutoff")
    End Sub

    Private Sub reminderscheduller()
        Logger.log("Start Reminder")
        Dim myform = New FormChequeReminder
        myform.DoWork()
        myform.Dispose()
        Logger.log("End Reminder")
    End Sub

    Private Sub orderCollectionscheduller()
        Logger.log("Start OrderCollection")
        Dim myform = New FormOrderCollection
        myform.DoWork()
        myform.Dispose()
        Logger.log("End OrderCollection")
    End Sub

    Private Sub getcurrentcutoff()
        If IsNothing(DbAdapter1) Then
            DbAdapter1 = New DbAdapter
            dbtools1.Userid = DbAdapter1.userid
            dbtools1.Password = DbAdapter1.password
        End If
        Dim sqlstr = "select * from shop.cutoff c " &
                     " left join shop.paramhd p on p.paramname = 'Current Cutoff' " &
                     " where c.cutoffid = p.nvalue "
        ds = New DataSet
        Dim mymessage As String = String.Empty
        If DbAdapter1.GetDataSet(sqlstr, ds, mymessage) Then
        Else
            Logger.log(mymessage)
            Me.Close()
        End If
    End Sub

End Class