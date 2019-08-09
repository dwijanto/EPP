Imports System.Threading
Imports EPP.HelperClass
Imports EPP.PublicClass
Imports System.Text
Public Class FormKAM

    Private Property sqlstr As String
    Private DS As DataSet
    Private BS As BindingSource
    Private CM As CurrencyManager


    Private Sub FormCutoff_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not IsNothing(getChanges) Then
            Select Case MessageBox.Show("Save unsaved records?", "Unsaved records", MessageBoxButtons.YesNoCancel)
                Case Windows.Forms.DialogResult.Yes
                    ToolStripButton4.PerformClick()
                Case Windows.Forms.DialogResult.Cancel
                    e.Cancel = True
            End Select
        End If
    End Sub

    Private Function getChanges() As DataSet
        If IsNothing(BS) Then
            Return DS
        End If
        Me.Validate()
        BS.EndEdit()
        Return DS.GetChanges()
    End Function

    Private Sub FormCutoff_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadData()
        With DataGridView1.ColumnHeadersDefaultCellStyle
            .BackColor = SystemColors.GradientActiveCaption 'Color.DarkBlue
            .ForeColor = Color.Black 'Color.LemonChiffon
            .Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Bold)
            'New Font(Microsoft Sans Serif, 8.25pt, style=Bold)
        End With
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(255, 255, 192)
        'DataGridView1.BackgroundColor = SystemColors.InactiveCaption
        DataGridView1.RowsDefaultCellStyle.BackColor = Color.FromArgb(192, 255, 255)
        DataGridView1.EnableHeadersVisualStyles = False
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
        DataGridView1.ColumnHeadersHeight = 25
    End Sub

    Public Sub LoadData()
        'populate Data For Pembelian
        DS = New DataSet
        Dim mymessage As String = ""
        'sqlstr = "select companycode,companyname,companyid from finance.company  order by companycode"
        sqlstr = "select employeenumber,employeename,kamid from shop.kam order by employeenumber asc;"
        If DbAdapter1.GetDataSet(sqlstr, DS, mymessage) Then
            DS.Tables(0).TableName = "KAM"

            Dim idx(0) As DataColumn
            idx(0) = DS.Tables(0).Columns("kamid")
            DS.Tables(0).PrimaryKey = idx
            DS.Tables(0).PrimaryKey = idx
            DS.Tables(0).Columns("kamid").AutoIncrement = True
            DS.Tables(0).Columns("kamid").AutoIncrementSeed = -1
            DS.Tables(0).Columns("kamid").AutoIncrementStep = -1

            'Binding Object

            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
            'BS.Sort = "ordernum asc"
            DataGridView1.AutoGenerateColumns = False
            DataGridView1.DataSource = BS
            CM = CType(Me.BindingContext(BS), CurrencyManager)
        End If
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim drv As DataRowView = BS.AddNew()
        Dim dr = drv.Row        
        DS.Tables(0).Rows.Add(dr)
        'Dim myform = New FormInputJurnal(BS, PosNameBS, PosNameCollection)
        Dim myform = New FormInputKAM(BS)
        If Not myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
            DS.Tables(0).Rows.Remove(dr)
        End If

        Me.DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        If Not IsNothing(BS.Current) Then
            Dim myform = New FormInputKAM(BS)
            If Not myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
                BS.CancelEdit()
            End If

        Else
            MessageBox.Show("No record to update.")
        End If

        Me.DataGridView1.Invalidate()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(BS.Current) Then
            If MessageBox.Show("Delete this record(s)", "Delete Record", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                For Each dsrow In DataGridView1.SelectedRows
                    BS.RemoveAt(CType(dsrow, DataGridViewRow).Index)
                Next
            End If
        Else
            MessageBox.Show("No record to delete.")
        End If
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click

        Dim myposition = CM.Position
        Me.Validate()
        BS.EndEdit()

        Dim ds2 = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            'Dim docnumber As String = String.Empty

            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If DbAdapter1.KAMTx(Me, mye) Then
                Dim myquery = From row As DataRow In DS.Tables(0).Rows
                              Where row.RowState = DataRowState.Added

                For Each rec In myquery.ToArray
                    rec.Delete()
                Next
                DS.Merge(ds2)
                DS.AcceptChanges()
                BS.Position = myposition
                MessageBox.Show("Saved!")

                'LoadData()
            Else
                MessageBox.Show(mye.message)
            End If
        Else
            MessageBox.Show("Nothing to save.")
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        ToolStripButton2.PerformClick()
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        LoadData()
    End Sub
End Class