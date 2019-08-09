Imports EPP.HelperClass
Imports EPP.PublicClass
Imports System.Text
Imports System.Threading
Public Class FormSellingGroup
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents SellingGroupBS As BindingSource
    Dim SGPBS As BindingSource
    Dim SGEBS As BindingSource
    Dim PBS As BindingSource
    Dim EBS As BindingSource

    Dim DS As DataSet
    Dim sb As New StringBuilder

    Private Sub FormSupplierCategory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()

        sb.Append("select sellinggroupname,id from shop.sellinggroup order by sellinggroupname;")
        sb.Append("select * from shop.sellinggrouppackage s" &
                  " left join shop.packagename p on p.id = s.packageid;")
        sb.Append("select s.id,s.sellinggroupid,s.canbuy,e.*,sn || ' ' || givenname as employeename from shop.sellinggroupemployee s" &
                  " left join shop.employee e on e.employeenumber = s.employeenumber;")
        sb.Append("select p.packagename,p.id from shop.packagename p order by p.packagename;")
        sb.Append("select * from shop.employee order by employeenumber;")

        If DbAdapter1.GetDataSet(sb.ToString, DS, mymessage) Then
            Try
                DS.Tables(0).TableName = "Selling Group"
                DS.Tables(1).TableName = "Selling Group Package"
                DS.Tables(2).TableName = "Selling Group Employee"
                DS.Tables(3).TableName = "Package Name"
                DS.Tables(4).TableName = "Employee"
            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(4, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Try
                Select Case id
                    Case 1
                        ToolStripStatusLabel1.Text = message
                    Case 2
                        ToolStripStatusLabel1.Text = message
                    Case 4
                        Try
                            SellingGroupBS = New BindingSource
                            SGPBS = New BindingSource
                            SGEBS = New BindingSource
                            PBS = New BindingSource
                            EBS = New BindingSource

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("id")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).Columns("id").AutoIncrement = True
                            DS.Tables(0).Columns("id").AutoIncrementSeed = 0
                            DS.Tables(0).Columns("id").AutoIncrementStep = -1

                            Dim pk2(0) As DataColumn
                            pk2(0) = DS.Tables(1).Columns("id")
                            DS.Tables(1).PrimaryKey = pk2
                            DS.Tables(1).Columns("id").AutoIncrement = True
                            DS.Tables(1).Columns("id").AutoIncrementSeed = 0
                            DS.Tables(1).Columns("id").AutoIncrementStep = -1

                            Dim pk3(0) As DataColumn
                            pk3(0) = DS.Tables(2).Columns("id")
                            DS.Tables(2).PrimaryKey = pk3
                            DS.Tables(2).Columns("id").AutoIncrement = True
                            DS.Tables(2).Columns("id").AutoIncrementSeed = 0
                            DS.Tables(2).Columns("id").AutoIncrementStep = -1


                            Dim Rel As DataRelation
                            Dim SGCol As DataColumn
                            Dim SGPCol As DataColumn
                            Dim SGECol As DataColumn

                            SGCol = DS.Tables(0).Columns("id")
                            SGPCol = DS.Tables(1).Columns("sellinggroupid")
                            Rel = New DataRelation("SGPRel", SGCol, SGPCol)
                            DS.Relations.Add(Rel)

                            SGECol = DS.Tables(2).Columns("sellinggroupid")
                            Rel = New DataRelation("SGERel", SGCol, SGECol)
                            DS.Relations.Add(Rel)


                            SellingGroupBS.DataSource = DS.Tables(0)
                            SGPBS.DataSource = SellingGroupBS
                            SGPBS.DataMember = "SGPRel"

                            SGEBS.DataSource = SellingGroupBS
                            SGEBS.DataMember = "SGERel"

                            PBS.DataSource = DS.Tables(3)
                            EBS.DataSource = DS.Tables(4)

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = SellingGroupBS
                            'DataGridView1.RowTemplate.Height = 22
                            DataGridView2.AutoGenerateColumns = False
                            DataGridView2.DataSource = SGPBS
                            DataGridView3.AutoGenerateColumns = False
                            DataGridView3.DataSource = SGEBS
                            TextBox1.DataBindings.Clear()


                            'DirectCast(DataGridView2.Columns("PackageNameCB"), DataGridViewComboBoxColumn).DataSource = Nothing
                            DirectCast(DataGridView2.Columns("PackageNameCB"), DataGridViewComboBoxColumn).DataSource = PBS
                            DirectCast(DataGridView2.Columns("PackageNameCB"), DataGridViewComboBoxColumn).DisplayMember = "packagename"
                            DirectCast(DataGridView2.Columns("PackageNameCB"), DataGridViewComboBoxColumn).ValueMember = "id"
                            DirectCast(DataGridView2.Columns("PackageNameCB"), DataGridViewComboBoxColumn).DataPropertyName = "packageid"


                            TextBox1.DataBindings.Add(New Binding("Text", SellingGroupBS, "sellinggroupname", True, DataSourceUpdateMode.OnPropertyChanged, ""))


                        Catch ex As Exception
                            message = ex.Message
                        End Try

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub
    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        loaddata()
    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub SCBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles SellingGroupBS.ListChanged
        TextBox1.Enabled = Not IsNothing(SellingGroupBS.Current)

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        TabControl1.SelectedTab = TabPage1
        SellingGroupBS.AddNew()

    End Sub
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        SellingGroupBS.EndEdit()
        If Me.validate Then
            Try
                'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                Dim ds2 As DataSet
                ds2 = DS.GetChanges

                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)

                    If Not DbAdapter1.SellingGroupTx(Me, mye) Then
                        MessageBox.Show(mye.message)
                        Exit Sub
                    End If
                    DS.Merge(ds2)
                    DS.Tables(0).AcceptChanges()
                    'DS.Tables(1).AcceptChanges()
                    'DS.Tables(2).AcceptChanges()
                    DataGridView1.Invalidate()
                    MessageBox.Show("Saved.")
                End If
            Catch ex As Exception
                MessageBox.Show(" Error:: " & ex.Message)
            End Try
        End If
        DataGridView1.Invalidate()
    End Sub



    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        MyBase.Validate()

        For Each drv As DataRowView In SellingGroupBS.List
            If drv.Row.RowState = DataRowState.Modified Or drv.Row.RowState = DataRowState.Added Then
                If Not validaterow(drv) Then
                    myret = False
                End If
            End If
        Next

        'For Each drv As DataRowView In SGPBS 'SellingGroupPackageBS
        '    If drv.Row.RowState = DataRowState.Modified Or drv.Row.RowState = DataRowState.Added Then
        '        If Not validaterow(drv, 2) Then
        '            myret = False
        '        End If
        '    End If
        'Next
        Return myret
    End Function

    Private Function validaterow(ByVal drv As DataRowView) As Boolean
        Dim myret As Boolean = True
        Dim sb As New StringBuilder
        'If type = 1 Then
        If IsDBNull(drv.Row.Item("sellinggroupname")) Then
            myret = False
            sb.Append("Selling Group Name cannot be blank")
        End If
        'ElseIf type = 2 Then
        'If IsDBNull(drv.Row.Item("maxqty")) Then
        '    myret = False
        '    sb.Append("Eligible qty cannot be blank.")
        'ElseIf drv.Row.Item("maxqty") = 0 Then
        '    myret = False
        '    If sb.Length > 0 Then
        '        sb.Append(",")
        '    End If
        '    sb.Append("Eligible qty must be greater than 0.")
        'End If
        'End If
        drv.Row.RowError = sb.ToString
        Return myret
    End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(SellingGroupBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    SellingGroupBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SellingGroupBS.CancelEdit()
    End Sub

    'Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
    '    If Not IsNothing(SellingGroupBS) Then
    '        Dim drv As DataRowView = SellingGroupBS.Current
    '        Dim myform As New FormGroupVendor
    '        myform.groupid = drv.Row.Item("groupid")
    '        myform.groupname = drv.Row.Item("groupname")
    '        myform.Show()
    '    End If

    'End Sub

    'Private Sub ToolStripButton2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
    '    If Not IsNothing(SellingGroupBS) Then
    '        Dim drv As DataRowView = SellingGroupBS.Current
    '        Dim myform As New FormGroupUser
    '        myform.groupid = drv.Row.Item("groupid")
    '        myform.groupname = drv.Row.Item("groupname")
    '        myform.Show()
    '    End If
    'End Sub

    Private Sub AddPackageToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddPackageToolStripMenuItem.Click
        If Not IsNothing(SellingGroupBS.Current) Then
            SellingGroupBS.EndEdit()
            Dim drv As DataRowView = SellingGroupBS.Current
            Dim mydrv As DataRowView = SGPBS.AddNew
            'mydrv.Row.Item("sellinggroupid") = drv.Item("id")
            mydrv.Row.Item("maxqty") = 1
            mydrv.EndEdit()
            'myform.groupid = drv.Row.Item("groupid")
            'myform.groupname = drv.Row.Item("groupname")
            'myform.Show()

        End If
    End Sub

    Private Sub AddEmployeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddEmployeeToolStripMenuItem.Click
        If Not IsNothing(SellingGroupBS.Current) Then

            Dim drv As DataRowView = SGEBS.AddNew()
            Dim dr = drv.Row
            dr.Item("canbuy") = True
            drv.EndEdit()
            'dr.Item("id") = 0
            'DS.Tables(0).Rows.Add(dr)
            'Dim myform = New FormInputJurnal(BS, PosNameBS, PosNameCollection)
            'Dim myform = New FormInputCutoff(BS)
            Dim myform = New FormInputSellingGroupEmployee(DS, SGEBS)

            If Not myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
                DS.Tables(2).Rows.Remove(dr)
            End If

            Me.DataGridView1.Invalidate()


        End If


    End Sub

    Private Sub DeleteRecordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteRecordToolStripMenuItem.Click
        If Not IsNothing(SellingGroupBS.Current) Then
            Dim myform = New FormInputSellingGroupEmployee(DS, SGEBS)
            If Not myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
                SGEBS.CancelEdit()
            End If

        Else
            MessageBox.Show("No record to update.")
        End If

        Me.DataGridView1.Invalidate()
    End Sub

    Private Sub DeletePackageToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeletePackageToolStripMenuItem.Click
        If Not IsNothing(SGPBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView2.SelectedRows
                    SGPBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub DeleteRecordToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteRecordToolStripMenuItem1.Click
        If Not IsNothing(SGEBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView3.SelectedRows
                    SGEBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub
End Class