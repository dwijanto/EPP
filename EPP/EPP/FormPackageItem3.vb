Imports System.Threading
Imports EPP.HelperClass
Imports EPP.PublicClass
Imports System.Text
Public Class FormPackageItem3
    Public Property PackageId As Long = 0
    Private DS As DataSet
    Dim sqlstr As String = String.Empty
    Dim HDBS As BindingSource
    Dim DTLBS As BindingSource
    Dim PackageItemBS As BindingSource
    Dim ItemBS As BindingSource
    Dim ItemBS2 As BindingSource
    Dim SGBS As BindingSource
    Dim cm As CurrencyManager
    Dim SB As StringBuilder
    Dim myThread As New Thread(AddressOf DoLoad)

    Dim HDDR As DataRow

    Public Sub DoLoad()
        'populate Data 

        ProgressReport(6, "Marquee")
        ProgressReport(2, "Loading... Please wait!")

        DS = New DataSet
        Dim mymessage As String = ""
        SB = New StringBuilder
        SB.Append(String.Format("select s.packagename,s.mainitem,s.active,s.promotionalitem,s.id,i.refno,d.descriptionname,s.maxqty from shop.packagename s" &
                                " left join shop.item i on i.itemid = s.mainitem" &
                                " left join shop.description d on d.descriptionid = i.descriptionid order by s.id;"))   '&  " where id = {0};", PackageId))
        SB.Append(String.Format("select i.refno,d.descriptionname,ip.itempackageid,ip.packageid,i.itemid as id from shop.itempackage ip" &
                 " left join shop.item i on i.itemid = ip.itemid" &
                 " left join shop.description d on d.descriptionid = i.descriptionid;")) 'where ip.packageid =  {0};", PackageId))
        SB.Append("select packagename,i.refno,mainitem,active,promotionalitem,id from shop.packagename p left join shop.item i on i.itemid = p.mainitem;")
        SB.Append("select i.refno,i.itemid,d.descriptionname from shop.item i " &
                  " left join shop.description d on d.descriptionid = i.descriptionid order by i.refno ;")
        SB.Append("select i.refno,i.itemid,d.descriptionname from shop.item i " &
                  " left join shop.description d on d.descriptionid = i.descriptionid order by i.refno;")

        SB.Append("select sp.*,sg.sellinggroupname from shop.sellinggrouppackage sp " &
                  " left join shop.sellinggroup sg on sg.id = sp.sellinggroupid;")
        sqlstr = SB.ToString
        If DbAdapter1.GetDataSet(sqlstr, DS, mymessage) Then
            DS.Tables(0).TableName = "HDDS"
            Dim idx(0) As DataColumn
            idx(0) = DS.Tables(0).Columns("id")
            DS.Tables(0).PrimaryKey = idx
            DS.Tables(0).Columns("id").AutoIncrement = True
            DS.Tables(0).Columns("id").AutoIncrementSeed = -1
            DS.Tables(0).Columns("id").AutoIncrementStep = -1

            DS.Tables(1).TableName = "DTLDS"
            Dim idx1(0) As DataColumn
            idx1(0) = DS.Tables(1).Columns("itempackageid")
            DS.Tables(1).PrimaryKey = idx1
            DS.Tables(1).Columns("itempackageid").AutoIncrement = True
            DS.Tables(1).Columns("itempackageid").AutoIncrementSeed = -1
            DS.Tables(1).Columns("itempackageid").AutoIncrementStep = -1

            Dim relation As New DataRelation("hdrel", DS.Tables(0).Columns("id"), DS.Tables(1).Columns("packageid"))
            DS.Relations.Add(relation)

            relation = New DataRelation("sgrel", DS.Tables(0).Columns("id"), DS.Tables(5).Columns("packageid"))
            DS.Relations.Add(relation)
            DS.Tables(5).TableName = "SGDS"
            Dim idx5(0) As DataColumn
            idx5(0) = DS.Tables(5).Columns("id")
            DS.Tables(5).PrimaryKey = idx5
            DS.Tables(5).Columns("id").AutoIncrement = True
            DS.Tables(5).Columns("id").AutoIncrementSeed = -1
            DS.Tables(5).Columns("id").AutoIncrementStep = -1

            HDBS = New BindingSource
            DTLBS = New BindingSource
            PackageItemBS = New BindingSource
            ItemBS = New BindingSource
            ItemBS2 = New BindingSource
            SGBS = New BindingSource

            HDBS.DataSource = DS.Tables(0)
            DTLBS.DataSource = HDBS
            DTLBS.DataMember = "hdrel"

            SGBS.DataSource = HDBS
            SGBS.DataMember = "sgrel"

            'PackageItemBS.DataSource = DS.Tables(2)
            ItemBS.DataSource = DS.Tables(3)
            ItemBS2.DataSource = DS.Tables(4)
            cm = CType(Me.BindingContext(HDBS), CurrencyManager)

            'binding data
            ProgressReport(1, "BindingData")
            ProgressReport(2, "Load. Done!")
        Else
            ProgressReport(2, mymessage)
        End If
        ProgressReport(5, "Continuous")
    End Sub

    Private Sub FormPackageItem_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'If Not IsNothing(getChanges) Then
        HDBS.EndEdit()
        DTLBS.EndEdit()
        Dim ds2 = DS.GetChanges
        If Not IsNothing(ds2) Then
            Select Case MessageBox.Show("Save unsaved records?", "Unsaved records", MessageBoxButtons.YesNoCancel)
                Case Windows.Forms.DialogResult.Yes
                    ToolStripButton4.PerformClick()
                Case Windows.Forms.DialogResult.Cancel
                    e.Cancel = True
            End Select
        End If
    End Sub
    Private Function getChanges() As DataSet
        If IsNothing(HDBS) Then
            Return Nothing
        End If
        'Me.Validate()
        'HDBS.EndEdit()
        'DS.EnforceConstraints = False
        Return DS.GetChanges()

    End Function
    Private Sub FormPackageItem_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadData()
    End Sub

    Private Sub LoadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoLoad)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is complete.")
        End If
    End Sub


    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1

                    'Clear Binding
                    TextBox1.DataBindings.Clear()
                    TextBox2.DataBindings.Clear() 'try to open
                    ComboBox1.DataBindings.Clear()
                    CheckBox1.DataBindings.Clear()
                    CheckBox2.DataBindings.Clear()
                    TextBox3.DataBindings.Clear()
                    DataGridView1.DataSource = Nothing

                    BindingNavigator1.BindingSource = HDBS

                    'Add Binding
                    TextBox1.DataBindings.Add("Text", HDBS, "packagename", True, DataSourceUpdateMode.OnPropertyChanged, "")
                    With CheckBox1
                        .DataBindings.Add("Checked", HDBS, "active", True)
                        '.DataBindings.DefaultDataSourceUpdateMode = DataSourceUpdateMode.OnPropertyChanged
                    End With

                    With CheckBox2
                        .DataBindings.Add("Checked", HDBS, "promotionalitem", True)
                        '.DataBindings.DefaultDataSourceUpdateMode = DataSourceUpdateMode.OnPropertyChanged
                    End With

                    'CheckBox1
                    TextBox2.DataBindings.Add("Text", HDBS, "descriptionname", True, DataSourceUpdateMode.OnPropertyChanged)
                    TextBox3.DataBindings.Add("Text", HDBS, "maxqty", True, DataSourceUpdateMode.OnPropertyChanged, "")
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = DTLBS

                    DataGridView2.AutoGenerateColumns = False
                    DataGridView2.DataSource = SGBS


                    With ComboBox1
                        .DataSource = Nothing
                        .DisplayMember = "refno"
                        .ValueMember = "itemid"
                        .DataSource = ItemBS
                        .DataBindings.Add("SelectedValue", HDBS, "mainitem")
                        '.SelectedIndex = -1
                        '.DataBindings.DefaultDataSourceUpdateMode = DataSourceUpdateMode.OnPropertyChanged 'Sorry it doesn't work
                    End With

                    'With ComboBox1
                    '    .DataSource = ItemBS
                    '    .DisplayMember = "refno"
                    '    .ValueMember = "itemid"
                    '    .DataBindings.Add("SelectedValue", HDBS, "mainitem")
                    'End With

                    'If IsNothing(HDBS.Current) Then
                    '    'HDBS.AddNew()
                    '    'HDBS.Current.row.Item("active") = True
                    '    'HDBS.EndEdit()
                    '    'Debug.Print("hello")
                    'Else
                    '    Dim mycheck = True
                    '    If Not IsDBNull(HDBS.Current.row.item("active")) Then
                    '        mycheck = HDBS.Current.row.item("active")
                    '    End If
                    '    RadioButton1.Checked = mycheck
                    '    RadioButton2.Checked = Not RadioButton1.Checked
                    'End If

                    If IsNothing(HDBS.Current) Then
                        Dim mydr As DataRowView = HDBS.AddNew()
                        mydr.Row.Item("active") = True
                        mydr.EndEdit()
                        ComboBox1.SelectedIndex = -1
                        TextBox2.Text = ""
                    End If
                Case 2
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

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        LoadData()
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        'PackageItemBS need to Refresh or do it int getpackagename
        'Dim myform As New FormGetPackageName(PackageItemBS)
        Dim myform As New FormGetPackageName(HDBS)
        If myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
            'HDBS.Filter = ""
            'Me.PackageId = myform.packageid
            'LoadData()
        End If
    End Sub



    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim mydr As DataRowView = HDBS.AddNew()
        'HDBS.MoveLast()
        'HDBS.Current.row.Item("active") = True
        mydr.Row.Item("active") = True
        mydr.Row.Item("promotionalitem") = False
        TextBox2.Text = ""
        'DS.Tables(0).Rows.Add(mydr.Row)

        mydr.EndEdit()

        'HDBS.MoveLast()
        'HDBS.EndEdit()
    End Sub

    Private Sub ComboBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox1.KeyUp
        'If e.KeyCode = Keys.Enter Then
        '    Dim oritext = ComboBox1.Text
        '    For i = 0 To ComboBox1.Items.Count - 1
        '        If ComboBox1.Items(i).row.item("refno").ToString.ToUpper = oritext.ToUpper Then
        '            ComboBox1.SelectedValue = ComboBox1.Items(i).row.item("itemid")
        '        End If
        '    Next

        'End If
    End Sub
    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        'Force combobox to commit the value


        Dim mycombo = DirectCast(sender, ComboBox)
        '1. Force the Combobox to commit the value 
        For Each binding As Binding In mycombo.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next

        Dim myrow As DataRowView = mycombo.SelectedItem
        Dim drv As DataRowView = HDBS.Current
        drv.Row.Item("mainitem") = myrow.Row.Item("itemid")
        drv.Row.Item("refno") = myrow.Row.Item("refno")
        'drv.Row.Item("descriptionname") = myrow.Row.Item("descriptionname")

        TextBox2.Text = myrow.Item("descriptionname")

        ''TextBox2.Text = ""
        'If Not IsNothing(ComboBox1.SelectedItem) Then
        '    Dim myrow As DataRow = ComboBox1.SelectedItem.row
        '    'If Not HDBS.Current.row.rowstate = DataRowState.Unchanged Then
        '    '    HDBS.Current.row.item("descriptionname") = myrow.Item("descriptionname")
        '    '    TextBox2.Text = myrow.Item("descriptionname")
        '    'End If
        '    'If Not IsNothing(HDBS.Current) Then
        '    'HDBS.Current.row.item("descriptionname") = myrow.Item("descriptionname")
        '    'HDBS.Current.row.item("refno") = myrow.Item("refno")
        '    'End If
        '    TextBox2.Text = myrow.Item("descriptionname")
        '    If Not IsNothing(HDBS.Current) Then
        '        If Not HDBS.Current.row.rowstate = DataRowState.Unchanged Then
        '            'HDBS.Current.row.item("mainitem") = myrow.Item("itemid")
        '            HDBS.Current.row.item("refno") = myrow.Item("refno")
        '            HDBS.Current.row.item("descriptionname") = myrow.Item("descriptionname")
        '        End If

        '        'HDBS.EndEdit()
        '    End If


        'End If
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        'Dim mycombo = DirectCast(sender, ComboBox)
        ''TextBox2.Text = ""
        'If Not IsNothing(ComboBox1.SelectedItem) Then
        '    Dim myrow As DataRow = ComboBox1.SelectedItem.row
        '    'If Not HDBS.Current.row.rowstate = DataRowState.Unchanged Then
        '    '    HDBS.Current.row.item("descriptionname") = myrow.Item("descriptionname")
        '    '    TextBox2.Text = myrow.Item("descriptionname")
        '    'End If
        '    'If Not IsNothing(HDBS.Current) Then
        '    'HDBS.Current.row.item("descriptionname") = myrow.Item("descriptionname")
        '    'HDBS.Current.row.item("refno") = myrow.Item("refno")
        '    'End If
        '    TextBox2.Text = myrow.Item("descriptionname")
        '    If Not IsNothing(HDBS.Current) Then
        '        If Not HDBS.Current.row.rowstate = DataRowState.Unchanged Then
        '            'HDBS.Current.row.item("mainitem") = myrow.Item("itemid")
        '            HDBS.Current.row.item("refno") = myrow.Item("refno")
        '            HDBS.Current.row.item("descriptionname") = myrow.Item("descriptionname")
        '        End If

        '        'HDBS.EndEdit()
        '    End If


        'End If
    End Sub

    Private Sub AddModelNumberToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddModelNumberToolStripMenuItem.Click
        'HDBS.EndEdit()
        'Dim drv As DataRowView = DTLBS.AddNew()
        DTLBS.AddNew()

        'DS.Tables(2).Rows.Add(drv.Row)
        Dim myform As New FormInputItemMustBeSold(DS.Tables(4), DTLBS)
        'If myform.ShowDialog = Windows.Forms.DialogResult.Cancel Then
        If Not myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
            If Not IsNothing(DTLBS.Current) Then
                'DTLBS.RemoveCurrent()
                'DS.Tables(2).Rows.Remove(drv.Row)
            End If
        Else
            'DTLBS.EndEdit()
        End If

        DataGridView1.Invalidate()
    End Sub

    Private Sub UpdateRecordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateRecordToolStripMenuItem.Click
        Dim myform As New FormInputItemMustBeSold(DS.Tables(3), DTLBS)
        'If myform.ShowDialog = Windows.Forms.DialogResult.Cancel Then
        If Not myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
            If Not IsNothing(DTLBS.Current) Then
                'DTLBS.RemoveCurrent()
                'DS.Tables(2).Rows.Remove(drv.Row)
                DTLBS.CancelEdit()
            End If
        Else
            DTLBS.EndEdit()
        End If
    End Sub

    Private Sub DeleteRecordToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteRecordToolStripMenuItem.Click
        If Not IsNothing(DTLBS.Current) Then
            If MessageBox.Show("Delete this record(s)", "Delete Record", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                DTLBS.RemoveCurrent()
            End If
        Else
            MessageBox.Show("No record to delete.")
        End If
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(HDBS.Current) Then
            If MessageBox.Show("Delete this record(s)", "Delete Record", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                HDBS.RemoveCurrent()
                If IsNothing(HDBS.Current) Then
                    'HDBS.AddNew()
                    ' HDBS.Current.row.Item("active") = True
                End If
            End If
        Else
            MessageBox.Show("No record to delete.")
        End If
    End Sub

    Public Overloads Function validate() As Boolean
        MyBase.Validate()
        Dim myret As Boolean = True

        'Check, if not avail mainitem in detail then create one
        For Each drv As DataRowView In HDBS.List
            'For Promotionalitem,if not set maxqty then set 1 as default
            If drv.Item("promotionalitem") Then
                If IsDBNull(drv.Item("maxqty")) Then
                    drv.Item("maxqty") = 1
                End If
            End If
            Dim childrow = drv.Row.GetChildRows("hdrel")
            Dim hasmain As Boolean = False
            For Each dr In childrow
                If drv.Row.Item("mainitem") = dr.Item("id") Then
                    hasmain = True
                    Exit For
                End If
            Next
            If Not hasmain Then
                'add record
                If Not IsDBNull(drv.Row.Item("mainitem")) Then
                    Dim dr As DataRowView = DTLBS.AddNew
                    dr.Row.Item("packageid") = drv.Row.Item("id")
                    dr.Row.Item("id") = drv.Row.Item("mainitem")
                    dr.Row.Item("refno") = drv.Row.Item("refno")
                    dr.Row.Item("descriptionname") = drv.Row.Item("descriptionname")
                    dr.EndEdit()
                End If
            End If
        Next
        Return myret
    End Function

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Dim myposition = cm.Position
        Dim addrecord As Boolean = False
        If Me.validate() Then
            HDBS.EndEdit()
            DTLBS.EndEdit()

            Dim ds2 = DS.GetChanges
            If Not IsNothing(ds2) Then
                Dim mymessage As String = String.Empty
                Dim ra As Integer

                Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                If DbAdapter1.PackageItemTx(Me, mye) Then
                    DS.Merge(ds2)
                    'DS.AcceptChanges()
                    DS.Tables(0).AcceptChanges()
                    DS.Tables(1).AcceptChanges()
                    MessageBox.Show("Saved!")
                    If IsNothing(HDBS.Current) Then
                        LoadData()
                        HDBS.AddNew()
                        HDBS.Current.item("active") = True
                    End If
                Else
                    MessageBox.Show(mye.message)
                End If
            Else
                MessageBox.Show("Nothing to save.")
            End If
            Me.DataGridView1.Invalidate()
        Else

        End If
        Me.DataGridView1.Invalidate()
    End Sub

    
    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        Label5.Visible = CheckBox2.Checked
        TextBox3.Visible = CheckBox2.Checked
    End Sub
End Class