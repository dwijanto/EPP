Public Class FormInputItemMustBeSold
    Private bs As New BindingSource
    Private ItemBS As New BindingSource
    Dim myrow As DataRowView
    Public Sub New(ByRef DS As DataSet, ByRef bs As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        myrow = bs.Current
        Me.bs = bs

        ItemBS.DataSource = DS.Tables(1)

        ComboBox1.DataSource = ItemBS
        ComboBox1.DisplayMember = "refno"
        ComboBox1.ValueMember = "itemid"

        Label2.Text = "" & myrow.Item("descriptionname")

        ComboBox1.DataBindings.Add("selectedvalue", Me.bs, "id")

    End Sub
    Public Sub New(ByRef datatable As DataTable, ByRef bs As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        myrow = bs.Current
        Me.bs = bs

        ItemBS.DataSource = datatable

        ComboBox1.DataSource = ItemBS
        ComboBox1.DisplayMember = "refno"
        ComboBox1.ValueMember = "itemid"
        Label2.Text = "" & myrow.Item("descriptionname")

        ComboBox1.DataBindings.Add("SelectedValue", Me.bs, "id")

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Me.validate()
        Dim mycombotext As String = String.Empty
        Try
            If Not Me.validate Then
                Me.DialogResult = DialogResult.None
                'bs.CancelEdit()
                Exit Sub
            End If
            If Not IsNothing(ComboBox1.SelectedItem) Then
                myrow.Item("refno") = ComboBox1.SelectedItem.row.item("refno")
                myrow.Item("descriptionname") = ComboBox1.SelectedItem.row.item("descriptionname")
                myrow.Item("id") = ComboBox1.SelectedValue
                mycombotext = myrow.Item("refno")
            End If

            bs.EndEdit()
        Catch ex As Exception
            MessageBox.Show(ex.Message & " " & ". Reference No: " & mycombotext)
            Me.DialogResult = DialogResult.None
            'bs.CancelEdit()
        End Try

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        bs.CancelEdit()
    End Sub

    Private Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        MyBase.Validate()
        If ComboBox1.SelectedIndex = -1 Then
            myret = False
        End If
        Return myret
    End Function

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

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged, ComboBox1.SelectedIndexChanged
        ''Debug.Print("hello")
        'Dim mycombo = DirectCast(sender, ComboBox)
        'Label2.Text = ""
        'If Not IsNothing(ComboBox1.SelectedItem) Then
        '    Dim myrow As DataRow = ComboBox1.SelectedItem.row
        '    Label2.Text = myrow.Item("descriptionname")
        'End If
    End Sub


    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted


        'Debug.Print("hello")
        Dim mycombo As ComboBox = DirectCast(sender, ComboBox)

        '1. Force the Combobox to commit the value 
        For Each binding As Binding In mycombo.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next
        Dim myrow As DataRowView = mycombo.SelectedItem
        Label2.Text = myrow.Item("descriptionname")

        'Label2.Text = ""
        'If Not IsNothing(ComboBox1.SelectedItem) Then
        '    Dim myrow As DataRow = ComboBox1.SelectedItem.row
        '    Label2.Text = myrow.Item("descriptionname")
        'End If
    End Sub
End Class