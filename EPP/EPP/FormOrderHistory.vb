Imports EPP.PublicClass
Public Class FormOrderHistory
    'load data  based on hdid
    Public Property bs As BindingSource
    Dim DS As New DataSet
    Dim pohdid As Long
    Dim dr As DataRow
    Dim PODTLBS As BindingSource
    Public Sub New(ByVal bs As BindingSource)
        ' This call is required by the designer.
        InitializeComponent()
        Me.bs = bs
        dr = CType(bs.Current, DataRowView).Row
        pohdid = dr.Item("pohdid")
        ' Add any initialization after the InitializeComponent() call.

    End Sub


    Private Sub FormOrderHistory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Private Sub loaddata()
        Dim sqlstr = "select *,shop.validamount(ph.status,pd.qty,pd.confirmedqty,pd.staffprice) as amount,case ph.status when 1 then 'N/A' else confirmedqty::character varying end as confirmed, staffprice * confirmedqty as amount1 from shop.podtl pd " &
                     " left join shop.pohd ph on ph.pohdid = pd.pohdid" &
                     " left join shop.item i on i.itemid = pd.itemid " &
                     " left join shop.product p on p.productid = i.productid " &
                     " left join shop.brand  b on b.brandid = i.brandid " &
                     " left join shop.description d on d.descriptionid = i.descriptionid" &
                     " where pd.pohdid = " & pohdid & " order by podtlid;"

        Dim mymessage As String = String.Empty
        If DbAdapter1.GetDataSet(sqlstr, DS, mymessage) Then
            DS.Tables(0).TableName = "PODTL"

            PODTLBS = New BindingSource
            PODTLBS.DataSource = DS.Tables(0)
            TextBox1.Text = dr.Item("billrefno")
            TextBox2.Text = Format(dr.Item("orderdate"), "dd-MMM-yyyy")
            TextBox3.Text = Format(dr.Item("totalamount"), "#,##0.00")
            TextBox4.Text = dr.Item("billingto")
            TextBox5.Text = "" & dr.Item("billingtoname")
            TextBox6.Text = dr.Item("statusname")
            DataGridView1.AutoGenerateColumns = False
            DataGridView1.DataSource = PODTLBS


        Else
            MessageBox.Show(mymessage)
        End If

    End Sub

End Class