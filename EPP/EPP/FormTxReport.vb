Imports System.Threading
Imports Microsoft.Office.Interop
Imports EPP.SharedClass
Imports System.Text
Imports EPP.PublicClass
Public Class FormTxReport
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoQuery)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)

    Dim combobs As BindingSource
    Dim DS As DataSet

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        Dim mymessage As String = String.Empty

        Dim mycutoff As Integer
        Dim sqlstr As String
        Dim mycriteria As String = String.Empty

        If ComboBox1.Text <> "All" Then
            mycutoff = ComboBox1.SelectedValue
            mycriteria = " where c.cutoffid = " & mycutoff
        End If
        sqlstr = "select ph.*,pd.*,s.statusname,e.sn || ' ' || e.givenname as employeename, b.sn || ' ' || b.givenname as billingtoname ,i.refno,d.descriptionname,pd.confirmedqty * pd.staffprice as totalamount,c.cutoff,p.chequenumber,p.bankcode from shop.pohd ph " &
                     " left join shop.podtl pd on pd.pohdid = ph.pohdid" &
                     " left join shop.cutoff c on c.cutoffid =  ph.cutoffid" &
                     " left join shop.status s on s.statusid = ph.status" &
                     " left join shop.employee e on e.employeenumber = ph.employeenumber" &
                     " left join shop.employee b on b.employeenumber = ph.billingto" &
                     " left join shop.item i on i.itemid = pd.itemid" &
                     " left join shop.payment p on p.paymentid = ph.pohdid" &
                     " left join shop.description d on d.descriptionid = i.descriptionid" & mycriteria

        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog

        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
            Dim reportname = "TX" '& GetCompanyName()
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback)
            myreport.Datasheet = 2
            myreport.Run(Me, e)
        End If

    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)
        sender.Name = "DB"
    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        owb.Names.Add("DBRange", RefersToR1C1:="=OFFSET('DB'!R1C1,0,0,COUNTA('DB'!C1),COUNTA('DB'!R1))")
        owb.Worksheets(1).select()
        Dim osheet As Excel.Worksheet = owb.Worksheets(1)


        owb.PivotCaches.Create(SourceType:=Excel.XlPivotTableSourceType.xlDatabase, SourceData:= _
         osheet.Name & "!DBRange", Version:=Excel.XlPivotTableVersionList.xlPivotTableVersion12).CreatePivotTable( _
        TableDestination:=osheet.Name & "!R6C1", TableName:="PivotTable1", DefaultVersion _
        :=Excel.XlPivotTableVersionList.xlPivotTableVersion12)

        With osheet.PivotTables("PivotTable1")
            .InGridDropZones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With
        osheet.PivotTables("PivotTable1").PivotFields("statusname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("employeename").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("employeename").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").PivotFields("billingtoname").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("billingtoname").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").PivotFields("billrefno").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("billrefno").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").PivotFields("refno").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("refno").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").PivotFields("descriptionname").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("descriptionname").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'osheet.PivotTables("PivotTable1").PivotFields("qty").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("qty").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'osheet.PivotTables("PivotTable1").PivotFields("confirmedqty").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("confirmedqty").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'osheet.PivotTables("PivotTable1").PivotFields("staffprice").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("staffprice").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty"), "Sum of Qty", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("confirmedqty"), "Sum of ConfirmedQty", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("staffprice"), "Sum of Staffprice", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalamount"), "Sum of totalamount", Excel.XlConsolidationFunction.xlSum)
    End Sub


    Private Sub FormTXReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ProgressReport(1, "Loading.... Please wait!")
        myThread = New System.Threading.Thread(myThreadDelegate)
        myThread.SetApartmentState(ApartmentState.MTA)
        myThread.Start()
    End Sub

    Sub DoQuery()
        'Get All Cutoff Data from Cutoff Table
        Dim sqlstr = "select 'All'::text as cutoff ,0::integer as cutoffid union all (select cutoff::text,cutoffid from shop.cutoff order by cutoff);select shop.getcurrentcutoff()::text as cutoff;"
        DS = New DataSet
        Dim mymessage As String = String.Empty
        If Not DbAdapter1.GetDataSet(sqlstr, DS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            combobs = New BindingSource
            combobs.DataSource = DS.Tables(0)
            If DS.Tables(0).Rows.Count > 0 Then

                ProgressReport(4, "Fill Combo Datasource")
                ProgressReport(1, "Loading.... Done!")
            End If
        End If
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByRef message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
                Case 2
                    Me.ToolStripStatusLabel2.Text = message
                Case 4
                    Me.ComboBox1.ValueMember = "cutoffid"
                    Me.ComboBox1.DisplayMember = "cutoff"
                    Me.ComboBox1.DataSource = combobs
                    Me.ComboBox1.Text = DS.Tables(1).Rows(0).Item("cutoff")
                Case (5)
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