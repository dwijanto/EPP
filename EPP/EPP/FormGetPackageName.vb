Public Class FormGetPackageName
    Public Property packageid As Long
    Private BS As BindingSource
    Private Sub FormGetPackageName_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Public Sub New(ByRef BS As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()
        Me.BS = BS
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = Me.BS
        ' Add any initialization after the InitializeComponent() call.
        ToolStripComboBox1.SelectedIndex = 0
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.packageid = BS.Current.row.item("id")
    End Sub

    Private Sub ToolStripComboBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.Click

    End Sub

    Private Sub ToolStripTextBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.Click

    End Sub

    Private Sub ToolStripTextBox1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.Enter
        Debug.Print("enter")
    End Sub

    Private Sub ToolStripTextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ToolStripTextBox1.KeyUp
        Debug.Print("keyup")
        If e.KeyValue = Keys.Return Then
            'Me.packageid = BS.Current.row.item("id")
            'Button1.PerformClick()
            Dim myfilter As String = String.Empty
            Dim fields As String() = {"packagename", "refno"}

            myfilter = fields(ToolStripComboBox1.SelectedIndex) & " like " & ToolStripTextBox1.Text.ToString.Replace("'", "''")
            BS.Filter = myfilter
        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Button1.PerformClick()
    End Sub

    Private Sub ToolStripTextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripTextBox1.TextChanged
        Dim myfilter As String = String.Empty
        Dim fields As String() = {"packagename", "refno"}
        BS.Filter = ""
        If ToolStripComboBox1.SelectedIndex > -1 Then
            myfilter = fields(ToolStripComboBox1.SelectedIndex) & " like '*" & ToolStripTextBox1.Text.ToString.Replace("'", "''") & "*'"
            BS.Filter = myfilter
        End If

    End Sub
End Class