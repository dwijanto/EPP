Public Class FormInputSellingGroupEmployee

    Private bs As New BindingSource
    Private EmployeeBS As New BindingSource
    Dim myrow As DataRowView

    Public Sub New(ByRef DS As DataSet, ByRef bs As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        myrow = bs.Current
        Me.bs = bs

        EmployeeBS.DataSource = DS.Tables(4)

        ComboBox1.DataSource = EmployeeBS
        ComboBox1.DisplayMember = "employeenumber"
        ComboBox1.ValueMember = "employeenumber"

        Label2.Text = "" & myrow.Item("employeename")
        Label5.Text = "" & myrow.Item("department")

        ComboBox1.DataBindings.Add("selectedvalue", Me.bs, "employeenumber")
        CheckBox1.DataBindings.Add("checked", Me.bs, "canbuy")
    End Sub
    'Public Sub New(ByRef datatable As DataTable, ByRef bs As BindingSource)

    '    ' This call is required by the designer.
    '    InitializeComponent()
    '    ' Add any initialization after the InitializeComponent() call.
    '    myrow = bs.Current
    '    Me.bs = bs

    '    EmployeeBS.DataSource = datatable

    '    ComboBox1.DataSource = EmployeeBS
    '    ComboBox1.DisplayMember = "employeenumber"
    '    ComboBox1.ValueMember = "employeenumber"
    '    Label2.Text = "" & myrow.Item("sn") & " " & myrow.Item("givenname")
    '    Label5.Text = "" & myrow.Item("department")

    '    ComboBox1.DataBindings.Add("selectedvalue", Me.bs, "employeenumber")
    '    CheckBox1.DataBindings.Add("checked", Me.bs, "canbuy")

    'End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Me.validate()
        Dim mycombotext As String = String.Empty
        Try
            If Not Me.validate Then
                Me.DialogResult = DialogResult.None                
                Exit Sub
            End If
            bs.EndEdit()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.DialogResult = DialogResult.None      
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


    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim mycombo As ComboBox = DirectCast(sender, ComboBox)

        '1. Force the Combobox to commit the value 
        For Each binding As Binding In mycombo.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next
        Dim drv As DataRowView = mycombo.SelectedItem
        Label2.Text = "" & drv.Item("sn") & " " & drv.Item("givenname")
        Label5.Text = "" & drv.Item("department")
        myrow.Item("employeename") = Label2.Text
        myrow.Item("department") = Label5.Text
    End Sub
End Class