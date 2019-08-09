Imports EPP.PublicClass
Public Class FormSystemParameter

    Dim ds As DataSet
    Dim sqlstr As String
    Dim BS As BindingSource
    Dim cbBS As BindingSource


    Private Sub FormCutoff_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Button1.Text = "&Save" Then
            Select Case MessageBox.Show("Save unsaved records?", "Unsaved records", MessageBoxButtons.YesNoCancel)
                Case Windows.Forms.DialogResult.Yes
                    Button1.PerformClick()
                Case Windows.Forms.DialogResult.Cancel
                    e.Cancel = True
            End Select
        End If
    End Sub
    Private Sub assignData()
        'ds.Tables(0).Rows(0).Item("tvalue") = DateTimePicker1.Value
        ds.Tables(0).Rows(0).Item("nvalue") = ComboBox1.SelectedValue
        ds.Tables(0).Rows(1).Item("nvalue") = Decimal.Parse(TextBox1.Text)
        ds.Tables(0).Rows(2).Item("cvalue") = TextBox2.Text
        ds.Tables(0).Rows(3).Item("cvalue") = TextBox3.Text
        ds.Tables(0).Rows(4).Item("bvalue") = CheckBox1.Checked
        ds.Tables(0).Rows(5).Item("cvalue") = TextBox4.Text
        ds.Tables(0).Rows(6).Item("nvalue") = TextBox5.Text
        ds.Tables(0).Rows(7).Item("bvalue") = RadioButton1.Checked
        ds.Tables(0).Rows(8).Item("cvalue") = TextBox6.Text
        ds.Tables(0).Rows(9).Item("cvalue") = TextBox7.Text
        ds.Tables(0).Rows(10).Item("cvalue") = TextBox8.Text
        ds.Tables(0).Rows(11).Item("cvalue") = TextBox9.Text
    End Sub
    Private Function getChanges() As DataSet
        Me.Validate()
        assignData()
        
        If IsNothing(BS) Then
            Return ds
        End If
        Me.Validate()
        BS.EndEdit()
        Return ds.GetChanges()
    End Function
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Validate()
        If Button1.Text <> "&OK" Then
            assignData()
            Dim ds2 = ds.GetChanges
            If Not IsNothing(ds2) Then
                Dim mymessage As String = String.Empty
                Dim ra As Integer
                Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                If DbAdapter1.systemparam(Me, mye) Then
                    Dim myquery = From row As DataRow In ds.Tables(0).Rows
                                  Where row.RowState = DataRowState.Added

                    For Each rec In myquery.ToArray
                        rec.Delete()
                    Next
                    ds.Merge(ds2)
                    ds.AcceptChanges()
                    'LoadData()
                    Button1.Text = "&OK"
                    MessageBox.Show("Saved.")

                    'Me.Close()
                Else
                    MessageBox.Show(mye.message)
                End If
            Else
                MessageBox.Show("Nothing to save.")
            End If
            Button1.Text = "&OK"
        End If
        Me.Close()
    End Sub

    Private Sub FormCutoff_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadData()
        Button1.Text = "&OK"
    End Sub

    Public Sub LoadData()
        'populate Data 
        DS = New DataSet
        Dim mymessage As String = ""
        sqlstr = "select * from shop.paramhd ph  where ph.parent = 1 order by ivalue;" &
                 "select s.*,s.cutoff::character varying as mydate from shop.cutoff s where date_part('Year',s.cutoff) >= " & Today.Year & " order by s.cutoff asc; "
        If DbAdapter1.GetDataSet(sqlstr, DS, mymessage) Then
            ds.Tables(0).TableName = "Param"
            Dim idx(0) As DataColumn
            idx(0) = ds.Tables(0).Columns("paramhdid")
            DS.Tables(0).PrimaryKey = idx
            'Binding Object
            BS = New BindingSource
            cbBS = New BindingSource

            BS.DataSource = ds.Tables(0)


            cbBS.DataSource = ds.Tables(1)
            ComboBox1.DisplayMember = "mydate"
            ComboBox1.ValueMember = "cutoffid"
            ComboBox1.DataSource = cbBS

            'BS.Sort = "ordernum asc"

            'DateTimePicker1.Value = ds.Tables(0).Rows(0).Item("tvalue")
            ComboBox1.SelectedValue = ds.Tables(0).Rows(0).Item("nvalue")
            TextBox1.Text = Format(ds.Tables(0).Rows(1).Item("nvalue"), "#,##0.00")
            TextBox2.Text = "" & ds.Tables(0).Rows(2).Item("cvalue")
            TextBox3.Text = "" & ds.Tables(0).Rows(3).Item("cvalue")
            CheckBox1.Checked = "" & ds.Tables(0).Rows(4).Item("bvalue")
            TextBox4.Text = "" & ds.Tables(0).Rows(5).Item("cvalue")
            TextBox5.Text = Format(ds.Tables(0).Rows(6).Item("nvalue"), "#,##0")
            TextBox6.Text = "" & ds.Tables(0).Rows(8).Item("cvalue")
            TextBox7.Text = "" & ds.Tables(0).Rows(9).Item("cvalue")
            TextBox8.Text = "" & ds.Tables(0).Rows(10).Item("cvalue")
            TextBox9.Text = "" & ds.Tables(0).Rows(11).Item("cvalue")
            RadioButton1.Checked = ds.Tables(0).Rows(7).Item("bvalue")
            RadioButton2.Checked = Not (RadioButton1.Checked)
        End If
    End Sub

    Private Sub TextBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Validated
        TextBox1.Text = String.Format("{0:#,##0.00}", Decimal.Parse(TextBox1.Text))
        TextBox5.Text = String.Format("{0:#,##0}", Decimal.Parse(TextBox5.Text))
    End Sub


    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating, TextBox5.Validating
        If Not IsNumeric(sender.text) Then
            MessageBox.Show("Invalid data.")
            e.Cancel = True
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged, TextBox1.TextChanged, TextBox2.TextChanged, TextBox3.TextChanged, TextBox4.TextChanged, TextBox5.TextChanged, RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, ComboBox1.SelectedValueChanged, TextBox6.TextChanged, TextBox7.TextChanged, TextBox8.TextChanged, TextBox9.TextChanged
        Button1.Text = "&Save"
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Button1.Text = "&OK"
        Me.Close()
    End Sub
End Class