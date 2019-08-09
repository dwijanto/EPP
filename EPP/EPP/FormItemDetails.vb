Public Class FormItemDetails
    Dim bs As BindingSource
    Dim dr As DataRow
    Public Sub New(ByRef bs As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()
        Me.bs = bs
        dr = CType(bs.Current, DataRowView).Row
        ' Add any initialization after the InitializeComponent() call.
        TextBox1.Text = dr.Item("producttypename")
        TextBox2.Text = dr.Item("familyname")
        TextBox3.Text = dr.Item("brandname")
        TextBox4.Text = dr.Item("refno")
        TextBox5.Text = dr.Item("productname")
        TextBox6.Text = dr.Item("descriptionname")
        TextBox7.Text = String.Format("{0:#,##0.00}", dr.Item("retailprice"))
        TextBox8.Text = String.Format("{0:#,##0.00}", dr.Item("staffprice"))
        TextBox9.Text = String.Format("{0:#,##0.00}", dr.Item("validprice"))
        TextBox9.Visible = Not (dr.Item("staffprice") = dr.Item("validprice"))
        Label10.Visible = Not (dr.Item("staffprice") = dr.Item("validprice"))
    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class