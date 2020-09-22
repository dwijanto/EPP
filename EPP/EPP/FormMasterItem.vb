Imports EPP.HelperClass
Imports EPP.PublicClass
Imports System.Text
Public Class FormMasterItem

    Private Property sqlstr As String
    Private DS As DataSet
    Private BS As BindingSource
    Private CM As CurrencyManager
    Private ProductTypeBS As BindingSource
    Private FamilyBS As BindingSource
    Private BrandBS As BindingSource
    Private DescriptionBS As BindingSource
    Private ItemPriceBS As BindingSource
    Private ProductBS As BindingSource


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
            Return Nothing
        End If
        Me.Validate()
        BS.EndEdit()
        DS.EnforceConstraints = False
        Return DS.GetChanges()
    End Function

    Private Sub FormCutoff_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadData()
    End Sub

    Public Sub LoadData()
        'populate Data For Pembelian
        DS = New DataSet
        Dim mymessage As String = ""
        sqlstr = "select producttypename,familyname,productname,brandname,refno,descriptionname,retailprice,staffprice,promotionprice,promotionstartdate,promotionenddate,itemid,i.producttypeid,i.familyid::bigint,i.productid,i.brandid::bigint,i.descriptionid from shop.item i" &
                 " left join shop.producttype pt on pt.producttypeid = i.producttypeid" &
                 " left join shop.family f on f.familyid = i.familyid" &
                 " left join shop.product p on p.productid = i.productid" &
                 " left join shop.brand b on b.brandid = i.brandid" &
                 " left join shop.description d on d.descriptionid = i.descriptionid" &
                 " left join shop.itemprice ip on ip.itempriceid = i.itemid order by producttypename,familyname,productname,brandname,refno,descriptionname;" &
                 " select producttypeid,producttypename from shop.producttype order by producttypename;" &
                 " select familyid,familyname from shop.family order by familyname;" &
                 " select brandid,brandname from shop.brand order by brandname;" &
                 " select descriptionid,descriptionname from shop.description order by descriptionname;" &
                 " select itempriceid from shop.itemprice;" &
                 " select productid,productname from shop.product order by productname;"

        If DbAdapter1.GetDataSet(sqlstr, DS, mymessage) Then
            DS.Tables(0).TableName = "Item"

            Dim idx(0) As DataColumn
            idx(0) = DS.Tables(0).Columns("itemid")
            DS.Tables(0).PrimaryKey = idx

            DS.Tables(0).Columns("itemid").AutoIncrement = True
            DS.Tables(0).Columns("itemid").AutoIncrementSeed = -1
            DS.Tables(0).Columns("itemid").AutoIncrementStep = -1

            Dim fkProductTypeId As ForeignKeyConstraint
            Dim fkFamilyId As ForeignKeyConstraint
            Dim fkBrandId As ForeignKeyConstraint
            Dim fkDescriptionId As ForeignKeyConstraint
            Dim fkProductId As ForeignKeyConstraint

            fkProductTypeId = New ForeignKeyConstraint("FK1", DS.Tables(1).Columns("producttypeid"), DS.Tables(0).Columns("producttypeid"))
            fkProductTypeId.UpdateRule = Rule.Cascade
            DS.Tables(0).Constraints.Add(fkProductTypeId)

            fkFamilyId = New ForeignKeyConstraint("FK2", DS.Tables(2).Columns("familyid"), DS.Tables(0).Columns("familyid"))
            fkFamilyId.UpdateRule = Rule.Cascade
            DS.Tables(0).Constraints.Add(fkFamilyId)

            fkBrandId = New ForeignKeyConstraint("FK3", DS.Tables(3).Columns("brandid"), DS.Tables(0).Columns("brandid"))
            fkBrandId.UpdateRule = Rule.Cascade
            DS.Tables(0).Constraints.Add(fkBrandId)

            fkDescriptionId = New ForeignKeyConstraint("FK4", DS.Tables(4).Columns("descriptionid"), DS.Tables(0).Columns("descriptionid"))
            fkDescriptionId.UpdateRule = Rule.Cascade
            DS.Tables(0).Constraints.Add(fkDescriptionId)

            fkProductId = New ForeignKeyConstraint("FK5", DS.Tables(6).Columns("productid"), DS.Tables(0).Columns("productid"))
            fkProductId.UpdateRule = Rule.Cascade
            DS.Tables(0).Constraints.Add(fkProductId)

            Dim idx1(0) As DataColumn
            idx1(0) = DS.Tables(1).Columns("producttypeid")
            DS.Tables(1).PrimaryKey = idx1

            DS.Tables(1).Columns("producttypeid").AutoIncrement = True
            DS.Tables(1).Columns("producttypeid").AutoIncrementSeed = -1
            DS.Tables(1).Columns("producttypeid").AutoIncrementStep = -1




            Dim idx2(0) As DataColumn
            idx2(0) = DS.Tables(2).Columns("familyid")
            DS.Tables(2).PrimaryKey = idx2

            DS.Tables(2).Columns("familyid").AutoIncrement = True
            DS.Tables(2).Columns("familyid").AutoIncrementSeed = -1
            DS.Tables(2).Columns("familyid").AutoIncrementStep = -1

            Dim idx3(0) As DataColumn
            idx3(0) = DS.Tables(3).Columns("brandid")
            DS.Tables(3).PrimaryKey = idx3

            DS.Tables(3).Columns("brandid").AutoIncrement = True
            DS.Tables(3).Columns("brandid").AutoIncrementSeed = -1
            DS.Tables(3).Columns("brandid").AutoIncrementStep = -1


            Dim idx4(0) As DataColumn
            idx4(0) = DS.Tables(4).Columns("descriptionid")
            DS.Tables(4).PrimaryKey = idx4

            DS.Tables(4).Columns("descriptionid").AutoIncrement = True
            DS.Tables(4).Columns("descriptionid").AutoIncrementSeed = -1
            DS.Tables(4).Columns("descriptionid").AutoIncrementStep = -1

            Dim idx5(0) As DataColumn
            idx5(0) = DS.Tables(5).Columns("itempriceid")
            DS.Tables(5).PrimaryKey = idx5

            Dim idx6(0) As DataColumn
            idx6(0) = DS.Tables(6).Columns("productid")
            DS.Tables(6).PrimaryKey = idx6

            DS.Tables(6).Columns("productid").AutoIncrement = True
            DS.Tables(6).Columns("productid").AutoIncrementSeed = -1
            DS.Tables(6).Columns("productid").AutoIncrementStep = -1

            'Binding Object

            BS = New BindingSource

            ProductTypeBS = New BindingSource
            FamilyBS = New BindingSource
            BrandBS = New BindingSource
            DescriptionBS = New BindingSource
            ItemPriceBS = New BindingSource
            ProductBS = New BindingSource


            BS.DataSource = DS.Tables(0)
            ProductTypeBS.DataSource = DS.Tables(1)
            FamilyBS.DataSource = DS.Tables(2)
            BrandBS.DataSource = DS.Tables(3)
            DescriptionBS.DataSource = DS.Tables(4)
            ItemPriceBS.DataSource = DS.Tables(5)
            ProductBS.DataSource = DS.Tables(6)


            'BS.Sort = "ordernum asc"
            DataGridView1.AutoGenerateColumns = False
            DataGridView1.DataSource = BS
            CM = CType(Me.BindingContext(BS), CurrencyManager)
        End If
    End Sub
    'Add
    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        If Not IsNothing(BS) Then
            BS.Sort = ""
        End If
        Dim drv As DataRowView = BS.AddNew()
        Dim dr = drv.Row
        'dr.Item("promotionaldate") = Date.Today
        'dr.Item("collectiondate") = Date.Today
        'dr.Item("emailnotificationdate") = Date.Today
        'dr.Item("chequesubmitdate") = Date.Today
        DS.Tables(0).Rows.Add(dr)
        'Dim myform = New FormInputJurnal(BS, PosNameBS, PosNameCollection)
        Dim myform = New FormInputMasterItem(DS, BS)
        If Not myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
            DS.Tables(0).Rows.Remove(dr)
        End If

        Me.DataGridView1.Invalidate()
    End Sub
    'Update
    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        If Not IsNothing(BS.Current) Then
            Dim myform = New FormInputMasterItem(DS, BS)
            If Not myform.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                BS.CancelEdit()       
            End If
            myform.Dispose()
        Else
            MessageBox.Show("No record to update.")
        End If

        Me.DataGridView1.Invalidate()
    End Sub
    'Delete
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
    'save
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click

        Dim myposition = CM.Position
        Me.Validate()
        BS.EndEdit()
        DS.EnforceConstraints = False
        Dim ds2 = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer

            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If DbAdapter1.ItemTx(Me, mye) Then
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
        Me.DataGridView1.Invalidate()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        ToolStripButton2.PerformClick()
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        LoadData()
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        If Not IsNothing(ItemPriceBS) Then
            If MessageBox.Show("Remove Price for selected item(s)", "Remove Price", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                For Each dsrow In DataGridView1.SelectedRows
                    'BS.RemoveAt(CType(dsrow, DataGridViewRow).Index)
                    BS.Position = CType(dsrow, DataGridViewRow).Index
                    Dim myrow = BS.Current
                    Dim myposition = ItemPriceBS.Find("itempriceid", myrow.Row.Item("itemid").ToString)
                    If myposition >= 0 Then
                        ItemPriceBS.RemoveAt(myposition)
                        myrow.Item("retailprice") = DBNull.Value
                        myrow.Item("staffprice") = DBNull.Value
                    End If
                    
                Next
                Me.DataGridView1.Invalidate()
            End If
            
            'DS.Tables(5).Rows.Find()
        End If
        

    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        Dim myfilter As New StringBuilder

        If ToolStripTextBox1.Text <> "" Then
            myfilter.Append("descriptionname like '*" & ToolStripTextBox1.Text.Replace("'", "''") & "*'")
            myfilter.Append(" or refno like '*" & ToolStripTextBox1.Text.Replace("'", "''") & "*'")
            myfilter.Append(" or productname like '*" & ToolStripTextBox1.Text.Replace("'", "''") & "*'")
            myfilter.Append(" or brandname like '*" & ToolStripTextBox1.Text.Replace("'", "''") & "*'")
            myfilter.Append(" or familyname like '*" & ToolStripTextBox1.Text.Replace("'", "''") & "*'")
            myfilter.Append(" or producttypename like '*" & ToolStripTextBox1.Text.Replace("'", "''") & "*'")
            BS.Filter = myfilter.ToString
        Else
            BS.Filter = ""
        End If
    End Sub

    Private Sub ToolStripButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton8.Click
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""

        Dim myresult As DialogResult
        myresult = MessageBox.Show("Display as splitted worksheets?", "Select View mode", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3)

        Select Case myresult
            Case Windows.Forms.DialogResult.Yes
                splittedworksheets(Me, e)
            Case DialogResult.No
                nonsplittedresult(Me, e)
                
        End Select

    End Sub

    Private Sub FormattingReport()

    End Sub

    Private Sub PivotTable()

    End Sub

    Private Sub nonsplittedresult(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim myQueryWorksheetlist As New List(Of QueryWorksheet)




        Dim mymessage As String = String.Empty

        'Dim mycutoff As Integer
        Dim sqlstr As String
        'Dim mycriteria As String = String.Empty


        sqlstr = "select producttypename,i.familyid::bigint,familyname,productname,i.brandid::bigint,brandname,refno,descriptionname,retailprice,staffprice,promotionprice,promotionstartdate,promotionenddate from shop.item i" &
                " left join shop.producttype pt on pt.producttypeid = i.producttypeid" &
                " left join shop.family f on f.familyid = i.familyid" &
                " left join shop.product p on p.productid = i.productid" &
                " left join shop.brand b on b.brandid = i.brandid" &
                " left join shop.description d on d.descriptionid = i.descriptionid" &
                " left join shop.itemprice ip on ip.itempriceid = i.itemid where not retailprice isnull order by producttypename,familyname,productname,brandname,refno,descriptionname;"

        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog

        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
            Dim reportname = "Product" '& GetCompanyName()
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback)
            myreport.Datasheet = 1
            myreport.Run(Me, e)
        End If
    End Sub

    Private Sub splittedworksheets(ByVal formMasterItem As FormMasterItem, ByVal e As EventArgs)
        Dim mymessage As String = String.Empty
        Dim myQueryWorksheetList As New List(Of QueryWorksheet)

        'Dim mycutoff As Integer

        Dim sqlstr1 As String
        Dim sqlstr2 As String
        Dim sqlstr3 As String
        'Dim mycriteria As String = String.Empty


        sqlstr1 = "select refno as ""Product Id"",brandname as ""Brand"",i.familyid::bigint as ""Family Id"",familyname as ""Family"",productname as ""Product"",descriptionname as ""Description"",retailprice as ""Retail Price (HK$ / unit)"",staffprice as ""Staff Price (HK$ / unit)"",promotionprice as ""Promotion Price"",promotionstartdate as ""Promotion start date"",promotionenddate as ""Promotion end date"" " &
                  " from shop.item i " &
                  " left join shop.producttype pt on pt.producttypeid = i.producttypeid" &
                  " left join shop.family f on f.familyid = i.familyid " &
                  " left join shop.product p on p.productid = i.productid " &
                  " left join shop.brand b on b.brandid = i.brandid " &
                  " left join shop.description d on d.descriptionid = i.descriptionid " &
                  " left join shop.itemprice ip on ip.itempriceid = i.itemid " &
                  " where(i.producttypeid = 1) and not retailprice isnull " &
                  " order by producttypename,productname,refno,familyname,productname,brandname,descriptionname"
        sqlstr2 = "select refno as ""Product Id"",brandname as ""Brand"",i.familyid::bigint as ""Family Id"",familyname as ""Family"",productname as ""Product"",descriptionname as ""Description"",retailprice as ""Retail Price (HK$ / unit)"",staffprice as ""Staff Price (HK$ / unit)"",promotionprice as ""Promotion Price"",promotionstartdate as ""Promotion start date"",promotionenddate as ""Promotion end date"" " &
                 " from shop.item i " &
                 " left join shop.producttype pt on pt.producttypeid = i.producttypeid" &
                 " left join shop.family f on f.familyid = i.familyid " &
                 " left join shop.product p on p.productid = i.productid " &
                 " left join shop.brand b on b.brandid = i.brandid " &
                 " left join shop.description d on d.descriptionid = i.descriptionid " &
                 " left join shop.itemprice ip on ip.itempriceid = i.itemid " &
                 " where(i.producttypeid = 2) and not retailprice isnull" &
                 " order by producttypename,productname,refno,familyname,productname,brandname,descriptionname"
        sqlstr3 = "select refno as ""Product Id"",brandname as ""Brand"",i.familyid::bigint as ""Family Id"",familyname as ""Family"",productname as ""Product"",descriptionname as ""Description"",retailprice as ""Retail Price (HK$ / unit)"",staffprice as ""Staff Price (HK$ / unit)"",promotionprice as ""Promotion Price"",promotionstartdate as ""Promotion start date"",promotionenddate as ""Promotion end date"" " &
                 " from shop.item i " &
                 " left join shop.producttype pt on pt.producttypeid = i.producttypeid" &
                 " left join shop.family f on f.familyid = i.familyid " &
                 " left join shop.product p on p.productid = i.productid " &
                 " left join shop.brand b on b.brandid = i.brandid " &
                 " left join shop.description d on d.descriptionid = i.descriptionid " &
                 " left join shop.itemprice ip on ip.itempriceid = i.itemid " &
                 " where(i.producttypeid = 3) and not retailprice isnull" &
                 " order by producttypename,productname,refno,familyname,productname,brandname,descriptionname"

        Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog

        DirectoryBrowser.Description = "Which directory do you want to use?"
        If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = DirectoryBrowser.SelectedPath 'Application.StartupPath & "\PrintOut"
            Dim reportname = "Product" '& GetCompanyName()
            Dim mycallback As FormatReportDelegate = AddressOf FormattingReportSplitted
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTableSplitted

            Dim myqueryworksheet = New QueryWorksheet With {.DataSheet = 1,
                                                            .SheetName = "Kitchen Electric",
                                                            .Sqlstr = sqlstr1,
                                                            .Location = "B7"}
            myQueryWorksheetList.Add(myqueryworksheet)
            myqueryworksheet = New QueryWorksheet With {.DataSheet = 2,
                                                        .SheetName = "Non-Food Electric",
                                                        .Sqlstr = sqlstr2,
                                                        .Location = "B7"}
            myQueryWorksheetList.Add(myqueryworksheet)
            myqueryworksheet = New QueryWorksheet With {.DataSheet = 3,
                                                        .SheetName = "Cookware",
                                                        .Sqlstr = sqlstr3,
                                                        .Location = "B7"}
            myQueryWorksheetList.Add(myqueryworksheet)

            Dim mytemplate = "\templates\Staff Price list.xltx"
            Dim myreport As New ExportToExcelFile(Me, myQueryWorksheetList, filename, reportname, mycallback, PivotCallback, mytemplate)
            myreport.Datasheet = 1
            myreport.Run(Me, e)
        End If
    End Sub

    Private Sub FormattingReportSplitted(ByRef sender As Object, ByRef e As EventArgs)
        Dim osheet = sender
        osheet.cells(3, 8).value = String.Format("Printed Date: {0:MMM-yyyy}", Date.Today.Date)
        osheet.columns("B:B").ColumnWidth = 16.75
    End Sub

    Private Sub PivotTableSplitted()

    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub
End Class