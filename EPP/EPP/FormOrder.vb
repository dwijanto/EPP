Imports EPP.PublicClass
Imports System.Text

Public Class FormOrder
    Implements IMessageFilter

    Dim WithEvents Timer1 As New Timer

    Dim myProduct As New List(Of Product)
    Dim myOrder As New List(Of Order)

    'Dim SelectedProduct As New List(Of Product)
    'Dim ProductBS As New BindingSource
    'Dim SelectedBS As New BindingSource

    Dim cm As CurrencyManager
    Dim cm1 As CurrencyManager
    Dim cm2 As CurrencyManager
    Dim mytotal As Decimal = 0
    Dim mytotalqty As Integer = 0
    Dim mytotalpromotion As Decimal = 0
    Dim mytotalqtypromotion As Decimal = 0

    Dim DS As DataSet

    Dim ProductTypeBS As BindingSource
    Dim ProductFamilyBS As BindingSource
    Dim BrandFamilyBS As BindingSource
    Dim ShoppingCartBS As BindingSource
    Dim ItemBS As BindingSource
    Dim OrderBS As New BindingSource
    Dim POHDBS As BindingSource
    Dim PODTBS As BindingSource
    Dim PackageItemBS As BindingSource

    Dim ProductTypeFilter As String = String.Empty
    Dim FamilyFilter As String = String.Empty
    Dim BrandFilter As String = String.Empty
    Dim familysb As New StringBuilder
    Dim sb As New StringBuilder
    Dim sqlstrSB As StringBuilder
    Dim mycutoff, mycollection, mynotification, mysubmit As Date
    Dim TimerClosing As Boolean = False

    Dim clrSelectedText As Color = Color.White
    Dim clrHighlight As Color = Color.CadetBlue
    Dim imagepath As String
    Dim imageext As String()

    Public Function GetMenuDesc() As String
        'Label1.Text = "Welcome, " & HelperClass1.UserInfo.DisplayName
        Return "App.Version: " & My.Application.Info.Version.ToString & " :: Server: " & DbAdapter1.ConnectionStringDict.Item("HOST") & ", Database: " & DbAdapter1.ConnectionStringDict.Item("DATABASE") & ", Userid: " & HelperClass1.UserId
    End Function

    Private Sub FormOrder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim sqlstr As String = String.Empty
        If Not IsNothing(ShoppingCartBS) Then
            If TimerClosing Then
                sqlstr = "delete from shop.shoppingcart where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber)
                DbAdapter1.ExecuteNonQuery(sqlstr)
                Exit Sub
            End If

            If DS.Tables(9).Rows(0).Item("cutoff") < Date.Now Then
                For Each drv As DataRowView In ShoppingCartBS.List
                    ShoppingCartBS.RemoveCurrent()
                Next
                sqlstr = "delete from shop.shoppingcart where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber)
                DbAdapter1.ExecuteNonQuery(sqlstr)
                Exit Sub
            End If
            If ShoppingCartBS.Count = 0 Then
                sqlstr = "delete from shop.shoppingcart where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber)
                DbAdapter1.ExecuteNonQuery(sqlstr)
            ElseIf Not IsNothing(getChanges) Then
              
                Select Case MessageBox.Show("You have some items in your shopping cart, Do you want to keep it ?", "Unsaved records", MessageBoxButtons.YesNoCancel)
                    Case Windows.Forms.DialogResult.Yes
                        saveShoppingCart()
                    Case Windows.Forms.DialogResult.No
                        sqlstr = "delete from shop.shoppingcart where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber)
                        DbAdapter1.ExecuteNonQuery(sqlstr)
                    Case Windows.Forms.DialogResult.Cancel
                        e.Cancel = True
                End Select

            End If
        End If
    End Sub

    Private Function getChanges() As DataSet
        If IsNothing(ShoppingCartBS) Then
            Return DS
        End If
        Me.Validate()
        ShoppingCartBS.EndEdit()
        Return DS.GetChanges()
    End Function
    Private Sub FormOrder_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsNothing(DbAdapter1) Then
            DbAdapter1 = New DbAdapter
            dbtools1.Userid = DbAdapter1.userid
            dbtools1.Password = DbAdapter1.password
            HelperClass1 = New HelperClass
        End If

        Label1.Text = "Welcome, " & HelperClass1.UserInfo.DisplayName
        Me.Text = "e-Staff Purchase System - " & GetMenuDesc()
        TextBox8.Text = HelperClass1.UserInfo.DisplayName
        TextBox9.Text = HelperClass1.UserInfo.telephonenumber
        TextBox4.Text = HelperClass1.UserInfo.Department
        TextBox5.Text = HelperClass1.UserInfo.email
        TextBox6.Text = HelperClass1.UserInfo.title
        TextBox7.Text = HelperClass1.UserInfo.employeenumber

        'Record Login
        Try
            loglogin(DbAdapter1.userid)
        Catch ex As Exception
        End Try

        loaddata()

        'With DataGridView1.ColumnHeadersDefaultCellStyle
        '    .BackColor = SystemColors.GradientActiveCaption 'Color.DarkBlue
        '    .ForeColor = Color.Black 'Color.LemonChiffon
        '    .Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Bold)
        '    'New Font(Microsoft Sans Serif, 8.25pt, style=Bold)
        'End With
        'DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(255, 255, 192)
        'DataGridView1.BackgroundColor = SystemColors.InactiveCaption
        'DataGridView1.RowsDefaultCellStyle.BackColor = Color.FromArgb(192, 255, 255)
        'DataGridView1.EnableHeadersVisualStyles = False
        'DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.EnableResizing
        'DataGridView1.ColumnHeadersHeight = 25
    End Sub

    Private Sub loglogin(ByVal userid As String)
        Dim applicationname As String = "eStaff User"
        Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        Dim computername As String = My.Computer.Name
        Dim time_stamp As DateTime = Now
        DbAdapter1.loglogin(applicationname, userid, username, computername, time_stamp)
    End Sub

    Private Sub loaddataori()
        'populate Data  
        DS = New DataSet

        Dim mymessage As String = ""
        Dim sqlstr As String = String.Empty
        sqlstrSB = New StringBuilder

        'sqlstr = "(select distinct 'Promotional Items' as producttypename,0 as producttypeid from shop.itemprice ip" &
        '        " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
        '        " union all" &
        '        " (select producttypename,producttypeid from shop.producttype order by producttypeid asc);" &
        '        " (select distinct 0 as producttypeid,i.familyid,f.familyname from shop.itemprice ip" &
        '        " left join shop.item i on i.itemid = ip.itempriceid" &
        '        " left join shop.family f on f.familyid = i.familyid" &
        '        " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
        '        " union all" &
        '        " (select distinct i.producttypeid,i.familyid,f.familyname from shop.item i" &
        '        " left join shop.family f on f.familyid = i.familyid" &
        '        " order by producttypeid,f.familyname);" &
        '        " (select distinct 0,0,i.brandid,b.brandname  from shop.item i" &
        '        " left join shop.brand b on b.brandid = i.brandid" &
        '        " left join shop.producttype p on p.producttypeid = i.producttypeid " &
        '        " order by brandname);" &
        '        " select *,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion from shop.shoppingcart sp" &
        '        " left join shop.item i on i.itemid = sp.itemid " &
        '        " left join shop.brand b on b.brandid = i.brandid " &
        '        " left join shop.product p on p.productid = i.productid " &
        '        " left join shop.description d on d.descriptionid = i.descriptionid" &
        '        " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '        " where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ";" &
        '        " select *,1 as quantity,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice  from shop.item i" &
        '        " left join shop.producttype pt on pt.producttypeid = i.producttypeid" &
        '        " left join shop.family f on f.familyid = i.familyid" &
        '        " left join shop.brand  b on b.brandid = i.brandid" &
        '        " left join shop.product  p on p.productid = i.productid" &
        '        " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '        " inner join shop.itemprice ip on ip.itempriceid = i.itemid;" &
        '        " select pd.pohdid,billrefno,orderdate,billingto,e.sn || ' ' || e.givenname as billingtoname,sum(shop.validamount(ph.status,pd.qty,pd.confirmedqty,pd.staffprice)) as totalamount , sum(pd.qty * pd.staffprice) as totalamountold,s.statusname,py.chequenumber,py.bankcode,ph.pohdid from shop.pohd ph" &
        '        " left join shop.employee e on e.employeenumber = ph.billingto" &
        '        " left join shop.status s on s.statusid = ph.status " &
        '        " left join shop.podtl pd on pd.pohdid = ph.pohdid" &
        '        " left join shop.payment py on py.paymentid = ph.pohdid" &
        '        " where ph.employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) &
        '        " group by pd.pohdid,billrefno,orderdate,billingto,billingtoname,statusname,ph.pohdid,py.chequenumber,py.bankcode order by ph.stamp desc;" &
        '        " select * from shop.pohd where pohdid = 0;" &
        '        " select * from shop.podtl where pohdid = 0;" &
        '        " select nvalue from shop.paramhd where paramname = 'Quota Amount';" &
        '        " select  c.cutoff,ph.nvalue from shop.paramhd ph left join shop.cutoff c on c.cutoffid =  ph.nvalue where paramname = 'Current Cutoff';" &
        '        " select nvalue from shop.paramhd where paramname = 'Quota Qty';" &
        '        " select bvalue from shop.paramhd where paramname = 'Quota Default';"

        'sqlstr = "(select distinct 'PROMOTIONAL ITEMS' as producttypename,0 as producttypeid from shop.itemprice ip" &
        '        " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
        '        " union all" &
        '        " (select producttypename,producttypeid from shop.producttype order by producttypeid asc);" &
        '        " (select distinct 0 as producttypeid,i.familyid,f.familyname from shop.itemprice ip" &
        '        " left join shop.item i on i.itemid = ip.itempriceid" &
        '        " left join shop.family f on f.familyid = i.familyid" &
        '        " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
        '        " union all" &
        '        " (select distinct i.producttypeid,i.familyid,f.familyname from shop.item i" &
        '        " left join shop.family f on f.familyid = i.familyid" &
        '        " order by producttypeid,f.familyname);" &
        '        " (select distinct 0,0,i.brandid,b.brandname  from shop.item i" &
        '        " left join shop.brand b on b.brandid = i.brandid" &
        '        " left join shop.producttype p on p.producttypeid = i.producttypeid " &
        '        " order by brandname);" &
        '        " select *,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion from shop.shoppingcart sp" &
        '        " left join shop.item i on i.itemid = sp.itemid " &
        '        " left join shop.brand b on b.brandid = i.brandid " &
        '        " left join shop.product p on p.productid = i.productid " &
        '        " left join shop.description d on d.descriptionid = i.descriptionid" &
        '        " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '        " where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ";" &
        '        " select *,1 as quantity,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice  from shop.item i" &
        '        " left join shop.producttype pt on pt.producttypeid = i.producttypeid" &
        '        " left join shop.family f on f.familyid = i.familyid" &
        '        " left join shop.brand  b on b.brandid = i.brandid" &
        '        " left join shop.product  p on p.productid = i.productid" &
        '        " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '        " inner join shop.itemprice ip on ip.itempriceid = i.itemid;" &
        '        " select pd.pohdid,billrefno,orderdate,billingto,e.sn || ' ' || e.givenname as billingtoname,sum(shop.validamount(ph.status,pd.qty,pd.confirmedqty,pd.staffprice)) as totalamount , sum(pd.qty * pd.staffprice) as totalamountold,s.statusname,py.chequenumber,py.bankcode,ph.pohdid from shop.pohd ph" &
        '        " left join shop.employee e on e.employeenumber = ph.billingto" &
        '        " left join shop.status s on s.statusid = ph.status " &
        '        " left join shop.podtl pd on pd.pohdid = ph.pohdid" &
        '        " left join shop.payment py on py.paymentid = ph.pohdid" &
        '        " where ph.employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) &
        '        " group by pd.pohdid,billrefno,orderdate,billingto,billingtoname,statusname,ph.pohdid,py.chequenumber,py.bankcode order by ph.stamp desc;" &
        '        " select * from shop.pohd where pohdid = 0;" &
        '        " select * from shop.podtl where pohdid = 0;" &
        '        " select nvalue from shop.paramhd where paramname = 'Quota Amount';" &
        '        " select  c.cutoff,ph.nvalue,c.emailnotificationdate,c.chequesubmitdate,c.collectiondate from shop.paramhd ph left join shop.cutoff c on c.cutoffid =  ph.nvalue where paramname = 'Current Cutoff';" &
        '        " select nvalue from shop.paramhd where paramname = 'Quota Qty';" &
        '        " select bvalue from shop.paramhd where paramname = 'Quota Default';" &
        '        " select cvalue from shop.paramhd where paramname = 'Special User';" &
        '        " select * from shop.employee where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ";"

        'sqlstr = "(select distinct 'PROMOTIONAL ITEMS' as producttypename,0 as producttypeid from shop.itemprice ip" &
        '       " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
        '       " union all" &
        '       " (select producttypename,producttypeid from shop.producttype order by producttypeid asc);" &
        '       " (select distinct 0 as producttypeid,i.familyid,f.familyname from shop.itemprice ip" &
        '       " left join shop.item i on i.itemid = ip.itempriceid" &
        '       " left join shop.family f on f.familyid = i.familyid" &
        '       " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
        '       " union all" &
        '       " (select distinct i.producttypeid,i.familyid,f.familyname from shop.item i" &
        '       " left join shop.family f on f.familyid = i.familyid" &
        '       " order by producttypeid,f.familyname);" &
        '       " (select distinct 0,0,i.brandid,b.brandname  from shop.item i" &
        '       " left join shop.brand b on b.brandid = i.brandid" &
        '       " left join shop.producttype p on p.producttypeid = i.producttypeid " &
        '       " order by brandname);" &
        '       " select sp.*,i.*,b.*,p.*,d.*,ip.itempriceid,ip.retailprice,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as staffprice,ip.promotionprice,ip.promotionstartdate,ip.promotionenddate,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice from shop.shoppingcart sp" &
        '       " left join shop.item i on i.itemid = sp.itemid " &
        '       " left join shop.brand b on b.brandid = i.brandid " &
        '       " left join shop.product p on p.productid = i.productid " &
        '       " left join shop.description d on d.descriptionid = i.descriptionid" &
        '       " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '       " where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ";" &
        '       " select *,1 as quantity,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice  from shop.item i" &
        '       " left join shop.producttype pt on pt.producttypeid = i.producttypeid" &
        '       " left join shop.family f on f.familyid = i.familyid" &
        '       " left join shop.brand  b on b.brandid = i.brandid" &
        '       " left join shop.product  p on p.productid = i.productid" &
        '       " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '       " inner join shop.itemprice ip on ip.itempriceid = i.itemid;" &
        '       " select pd.pohdid,billrefno,orderdate,billingto,e.sn || ' ' || e.givenname as billingtoname,sum(shop.validamount(ph.status,pd.qty,pd.confirmedqty,pd.staffprice)) as totalamount , sum(pd.qty * pd.staffprice) as totalamountold,s.statusname,py.chequenumber,py.bankcode,ph.pohdid from shop.pohd ph" &
        '       " left join shop.employee e on e.employeenumber = ph.billingto" &
        '       " left join shop.status s on s.statusid = ph.status " &
        '       " left join shop.podtl pd on pd.pohdid = ph.pohdid" &
        '       " left join shop.payment py on py.paymentid = ph.pohdid" &
        '       " where ph.employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) &
        '       " group by pd.pohdid,billrefno,orderdate,billingto,billingtoname,statusname,ph.pohdid,py.chequenumber,py.bankcode order by ph.stamp desc;" &
        '       " select * from shop.pohd where pohdid = 0;" &
        '       " select * from shop.podtl where pohdid = 0;" &
        '       " select nvalue from shop.paramhd where paramname = 'Quota Amount';" &
        '       " select  c.cutoff,ph.nvalue,c.emailnotificationdate,c.chequesubmitdate,c.collectiondate from shop.paramhd ph left join shop.cutoff c on c.cutoffid =  ph.nvalue where paramname = 'Current Cutoff';" &
        '       " select nvalue from shop.paramhd where paramname = 'Quota Qty';" &
        '       " select bvalue from shop.paramhd where paramname = 'Quota Default';" &
        '       " select cvalue from shop.paramhd where paramname = 'Special User';" &
        '       " select * from shop.employee where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ";" &
        '      " select cvalue from shop.paramhd where paramname = 'imagepath';" &
        '        " select cvalue from shop.paramhd where paramname = 'imageext';"

        ''ProductType 0
        sqlstrSB.Append("(select distinct 'PROMOTIONAL ITEMS' as producttypename,0 as producttypeid from shop.itemprice ip" &
               " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
               " union all" &
               " (select producttypename,producttypeid from shop.producttype order by producttypeid asc);")
        ''ProductType With Promo
        'sqlstrSB.Append("(select distinct 'PROMOTIONAL ITEMS' as producttypename,0 as producttypeid from shop.itemprice ip" &
        '      " where shop.validpromotion(promotionstartdate,promotionenddate," & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ",ip.itempriceid))" &
        '      " union all" &
        '      " (select producttypename,producttypeid from shop.producttype order by producttypeid asc);")

        'Family 1
        sqlstrSB.Append(" (select distinct 0 as producttypeid,i.familyid,f.familyname from shop.itemprice ip" &
               " left join shop.item i on i.itemid = ip.itempriceid" &
               " left join shop.family f on f.familyid = i.familyid" &
               " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
               " union all" &
               " (select distinct i.producttypeid,i.familyid,f.familyname from shop.item i" &
               " left join shop.family f on f.familyid = i.familyid" &
               " order by producttypeid,f.familyname);")
        'Brand 2
        sqlstrSB.Append(" (select distinct 0,0,i.brandid,b.brandname  from shop.item i" &
               " left join shop.brand b on b.brandid = i.brandid" &
               " left join shop.producttype p on p.producttypeid = i.producttypeid " &
               " order by brandname);")
        'Shopping Cart 3
        sqlstrSB.Append(" select sp.*,i.*,b.*,p.*,d.*,ip.itempriceid,ip.retailprice,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as staffprice,ip.promotionprice,ip.promotionstartdate,ip.promotionenddate,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice from shop.shoppingcart sp" &
               " left join shop.item i on i.itemid = sp.itemid " &
               " left join shop.brand b on b.brandid = i.brandid " &
               " left join shop.product p on p.productid = i.productid " &
               " left join shop.description d on d.descriptionid = i.descriptionid" &
               " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
               " where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ";")
        'Item
        'sqlstrSB.Append(" select *,1 as quantity,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice  from shop.item i" &
        '       " left join shop.producttype pt on pt.producttypeid = i.producttypeid" &
        '       " left join shop.family f on f.familyid = i.familyid" &
        '       " left join shop.brand  b on b.brandid = i.brandid" &
        '       " left join shop.product  p on p.productid = i.productid" &
        '       " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '       " inner join shop.itemprice ip on ip.itempriceid = i.itemid;")
        'Item
        'sqlstrSB.Append(" (select i.*,pt.*,f.*,b.*,p.*,d.*,ip.*,1 as quantity,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice,null::bigint as packageid" &
        '                " from shop.item i " &
        '                " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
        '                " left join shop.family f on f.familyid = i.familyid " &
        '                " left join shop.brand  b on b.brandid = i.brandid " &
        '                " left join shop.product  p on p.productid = i.productid " &
        '                " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '                " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '                " where not itemid in (select itemid as id from shop.itempackage i " &
        '                " left join shop.packagename p on p.id = i.packageid  where active" &
        '                " except all " &
        '                " select i.id from shop.itemmustbesold i " &
        '                " left join shop.itempackage ip on ip.itemid = i.id" &
        '                " left join shop.packagename p on p.id = ip.packageid  " &
        '                " where p.active)" &
        '                " union all" &
        '                " select i.*,pt.*,f.*,b.*,p.*,d.descriptionid,pn.packagename as descriptionname,foo.mainitem as itempriceid,foo.retailprice,foo.staffprice,foo.promotionprice,null::date as promotionstartdate,null::date promotionenddate,1 as quantity,null::boolean as promotion,foo.validprice,foo.packageid as packageid" &
        '                " from shop.packagename pn" &
        '                " left join shop.item i on i.itemid = pn.mainitem" &
        '                " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
        '                " left join shop.family f on f.familyid = i.familyid " &
        '                " left join shop.brand  b on b.brandid = i.brandid " &
        '                " left join shop.product  p on p.productid = i.productid " &
        '                " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '                " left join (select mainitem,packageid,sum(ip.retailprice) as retailprice,sum(ip.staffprice) as staffprice,sum(ip.promotionprice) as promotionprice,sum(shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate)) as validprice from shop.itempackage i " &
        '                " left join shop.packagename p on p.id = i.packageid" &
        '                " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '                " where active " &
        '                " group by mainitem,packageid) as foo on foo.mainitem = i.itemid" &
        '                " where pn.active);")
        'Item Add Package in Promotional Item 4
        sqlstrSB.Append(" (select i.*,pt.*,f.*,b.*,p.*,d.descriptionid,pn.packagename as descriptionname,foo.mainitem as itempriceid,foo.retailprice,foo.staffprice,foo.promotionprice,null::date as promotionstartdate,null::date promotionenddate,1 as quantity,true as promotion,foo.validprice,foo.packageid as packageid" &
                       " from shop.packagename pn" &
                       " left join shop.item i on i.itemid = pn.mainitem" &
                       " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
                       " left join shop.family f on f.familyid = i.familyid " &
                       " left join shop.brand  b on b.brandid = i.brandid " &
                       " left join shop.product  p on p.productid = i.productid " &
                       " left join shop.description  d on d.descriptionid = i.descriptionid" &
                       " left join (select mainitem,packageid,(sum(ip.retailprice) ) as retailprice,(sum(ip.staffprice) ) as staffprice,sum(ip.promotionprice) as promotionprice,sum(shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate)) as validprice from shop.itempackage i " &
                       " left join shop.packagename p on p.id = i.packageid" &
                       " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
                       " where active " &
                       " group by mainitem,packageid) as foo on foo.mainitem = i.itemid" &
                       " where pn.active and pn.promotionalitem)" &
                        " union all" &
                        "(select i.*,pt.*,f.*,b.*,p.*,d.*,ip.*,1 as quantity,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice,null::bigint as packageid" &
                       " from shop.item i " &
                       " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
                       " left join shop.family f on f.familyid = i.familyid " &
                       " left join shop.brand  b on b.brandid = i.brandid " &
                       " left join shop.product  p on p.productid = i.productid " &
                       " left join shop.description  d on d.descriptionid = i.descriptionid" &
                       " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
                       " where not itemid in (select itemid as id from shop.itempackage i " &
                       " left join shop.packagename p on p.id = i.packageid  where active" &
                       " except all " &
                       " select i.id from shop.itemmustbesold i " &
                       " left join shop.itempackage ip on ip.itemid = i.id" &
                       " left join shop.packagename p on p.id = ip.packageid  " &
                       " where p.active))" &
                       " union all" &
                       " (select i.*,pt.*,f.*,b.*,p.*,d.descriptionid,pn.packagename as descriptionname,foo.mainitem as itempriceid,foo.retailprice,foo.staffprice,foo.promotionprice,null::date as promotionstartdate,null::date promotionenddate,1 as quantity,null::boolean as promotion,foo.validprice,foo.packageid as packageid" &
                       " from shop.packagename pn" &
                       " left join shop.item i on i.itemid = pn.mainitem" &
                       " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
                       " left join shop.family f on f.familyid = i.familyid " &
                       " left join shop.brand  b on b.brandid = i.brandid " &
                       " left join shop.product  p on p.productid = i.productid " &
                       " left join shop.description  d on d.descriptionid = i.descriptionid" &
                       " left join (select mainitem,packageid,sum(ip.retailprice) as retailprice,sum(ip.staffprice) as staffprice,sum(ip.promotionprice) as promotionprice,sum(shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate)) as validprice from shop.itempackage i " &
                       " left join shop.packagename p on p.id = i.packageid" &
                       " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
                       " where active " &
                       " group by mainitem,packageid) as foo on foo.mainitem = i.itemid" &
                       " where pn.active and not pn.promotionalitem);")
        ''Item Add Package in Promotional Item Group
        'sqlstrSB.Append(" (select i.*,pt.*,f.*,b.*,p.*,d.descriptionid,pn.packagename as descriptionname,foo.mainitem as itempriceid,foo.retailprice,foo.staffprice,foo.promotionprice,null::date as promotionstartdate,null::date promotionenddate,1 as quantity,case promotion when true then true else null end as promotion,foo.validprice,foo.packageid as packageid" &
        '              " from shop.packagename pn" &
        '              " left join shop.item i on i.itemid = pn.mainitem" &
        '              " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
        '              " left join shop.family f on f.familyid = i.familyid " &
        '              " left join shop.brand  b on b.brandid = i.brandid " &
        '              " left join shop.product  p on p.productid = i.productid " &
        '              " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '              " left join (select mainitem,packageid,(sum(ip.retailprice) ) as retailprice,(sum(ip.staffprice) ) as staffprice,sum(ip.promotionprice) as promotionprice,sum(shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate," & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ",ip.itempriceid)) as validprice , shop.validpromotion(promotionstartdate,promotionenddate," & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ",ip.itempriceid)as promotion  from shop.itempackage i " &
        '              " left join shop.packagename p on p.id = i.packageid" &
        '              " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '              " where active " &
        '              " group by mainitem,packageid,promotion) as foo on foo.mainitem = i.itemid" &
        '              " where pn.active and pn.promotionalitem)" &
        '               " union all" &
        '               "(select i.*,pt.*,f.*,b.*,p.*,d.*,ip.*,1 as quantity,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice,null::bigint as packageid" &
        '              " from shop.item i " &
        '              " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
        '              " left join shop.family f on f.familyid = i.familyid " &
        '              " left join shop.brand  b on b.brandid = i.brandid " &
        '              " left join shop.product  p on p.productid = i.productid " &
        '              " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '              " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '              " where not itemid in (select itemid as id from shop.itempackage i " &
        '              " left join shop.packagename p on p.id = i.packageid  where active" &
        '              " except all " &
        '              " select i.id from shop.itemmustbesold i " &
        '              " left join shop.itempackage ip on ip.itemid = i.id" &
        '              " left join shop.packagename p on p.id = ip.packageid  " &
        '              " where p.active))" &
        '              " union all" &
        '              " (select i.*,pt.*,f.*,b.*,p.*,d.descriptionid,pn.packagename as descriptionname,foo.mainitem as itempriceid,foo.retailprice,foo.staffprice,foo.promotionprice,null::date as promotionstartdate,null::date promotionenddate,1 as quantity,null::boolean as promotion,foo.validprice,foo.packageid as packageid" &
        '              " from shop.packagename pn" &
        '              " left join shop.item i on i.itemid = pn.mainitem" &
        '              " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
        '              " left join shop.family f on f.familyid = i.familyid " &
        '              " left join shop.brand  b on b.brandid = i.brandid " &
        '              " left join shop.product  p on p.productid = i.productid " &
        '              " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '              " left join (select mainitem,packageid,sum(ip.retailprice) as retailprice,sum(ip.staffprice) as staffprice,sum(ip.promotionprice) as promotionprice,sum(shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate)) as validprice from shop.itempackage i " &
        '              " left join shop.packagename p on p.id = i.packageid" &
        '              " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '              " where active " &
        '              " group by mainitem,packageid) as foo on foo.mainitem = i.itemid" &
        '              " where pn.active and not pn.promotionalitem);")
        'History
        sqlstrSB.Append(String.Format(" select pd.pohdid,billrefno,orderdate,billingto,e.sn || ' ' || e.givenname as billingtoname,sum(shop.validamount(ph.status,pd.qty,pd.confirmedqty,pd.staffprice)) as totalamount , sum(pd.qty * pd.staffprice) as totalamountold,s.statusname,py.chequenumber,py.bankcode,ph.pohdid from shop.pohd ph" &
               " left join shop.employee e on e.employeenumber = ph.billingto" &
               " left join shop.status s on s.statusid = ph.status " &
               " left join shop.podtl pd on pd.pohdid = ph.pohdid" &
               " left join shop.payment py on py.paymentid = ph.pohdid" &
               " where ph.employeenumber = {0}", DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber)) &
               " group by pd.pohdid,billrefno,orderdate,billingto,billingtoname,statusname,ph.pohdid,py.chequenumber,py.bankcode order by ph.stamp desc;")
        sqlstrSB.Append(" select * from shop.pohd where pohdid = 0;")
        sqlstrSB.Append(" select * from shop.podtl where pohdid = 0;")
        sqlstrSB.Append(" select nvalue from shop.paramhd where paramname = 'Quota Amount';")
        sqlstrSB.Append(" select  c.cutoff,ph.nvalue,c.emailnotificationdate,c.chequesubmitdate,c.collectiondate from shop.paramhd ph left join shop.cutoff c on c.cutoffid =  ph.nvalue where paramname = 'Current Cutoff';")
        sqlstrSB.Append(" select nvalue from shop.paramhd where paramname = 'Quota Qty';")
        sqlstrSB.Append(" select bvalue from shop.paramhd where paramname = 'Quota Default';")
        sqlstrSB.Append(" select cvalue from shop.paramhd where paramname = 'Special User';")
        sqlstrSB.Append(" select * from shop.employee where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ";")
        sqlstrSB.Append(" select cvalue from shop.paramhd where paramname = 'imagepath';")
        sqlstrSB.Append(" select cvalue from shop.paramhd where paramname = 'imageext';")
        sqlstrSB.Append(" select itemid,packageid,p.packagename,ip.retailprice,ip.staffprice,ip.promotionprice,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice from shop.itempackage i " &
                        " left join shop.packagename p on p.id = i.packageid" &
                        " inner join shop.itemprice ip on ip.itempriceid = i.itemid where active ;")
        sqlstr = sqlstrSB.ToString
        Try


            If DbAdapter1.GetDataSet(sqlstr, DS, mymessage) Then
                DS.Tables(0).TableName = "ProductType"
                DS.Tables(1).TableName = "ProductFamily"
                DS.Tables(2).TableName = "All Brand"
                DS.Tables(3).TableName = "ShoppingCart"
                DS.Tables(4).TableName = "Items"
                DS.Tables(5).TableName = "Orders"
                DS.Tables(6).TableName = "POHD"
                DS.Tables(7).TableName = "PODTL"
                DS.Tables(8).TableName = "Quota Amount"
                DS.Tables(9).TableName = "Current Cutoff"
                DS.Tables(12).TableName = "Special User"
                DS.Tables(13).TableName = "Employee"
                DS.Tables(14).TableName = "imagepath"
                DS.Tables(15).TableName = "imageext"
                DS.Tables(16).TableName = "PackageItem"
                mycutoff = DS.Tables(9).Rows(0).Item("cutoff")
                mycollection = DS.Tables(9).Rows(0).Item("collectiondate")
                mynotification = DS.Tables(9).Rows(0).Item("emailnotificationdate")
                mysubmit = DS.Tables(9).Rows(0).Item("chequesubmitdate")
                imagepath = DS.Tables(14).Rows(0).Item(0)
                imageext = DS.Tables(15).Rows(0).Item(0).ToString.Split(",")

                DS.Tables(6).Columns("pohdid").AutoIncrement = True
                DS.Tables(6).Columns("pohdid").AutoIncrementSeed = -1
                DS.Tables(6).Columns("pohdid").AutoIncrementStep = -1



                Dim idx(0) As DataColumn
                idx(0) = DS.Tables(0).Columns("producttypeid")
                DS.Tables(0).PrimaryKey = idx

                Dim idx1(1) As DataColumn
                idx1(0) = DS.Tables(1).Columns("producttypeid")
                idx1(1) = DS.Tables(1).Columns("familyid")
                DS.Tables(1).PrimaryKey = idx1

                Dim idx3(0) As DataColumn
                idx3(0) = DS.Tables(3).Columns("refno")
                DS.Tables(3).PrimaryKey = idx3

                Dim idx4(0) As DataColumn
                idx4(0) = DS.Tables(4).Columns("itemid")
                DS.Tables(4).PrimaryKey = idx4

                Dim idx6(0) As DataColumn
                idx6(0) = DS.Tables(6).Columns("pohdid")
                DS.Tables(6).PrimaryKey = idx6

                Dim relation As New DataRelation("relProducttypefamily",
                                 DS.Tables(0).Columns("producttypeid"),
                                 DS.Tables(1).Columns("producttypeid"))
                DS.Relations.Add(relation)

                Dim relation6 As New DataRelation("relPOHDDetail", DS.Tables(6).Columns("pohdid"),
                                                  DS.Tables(7).Columns("pohdid"))
                DS.Relations.Add(relation6)

                'Binding Object


                ProductTypeBS = New BindingSource
                ProductFamilyBS = New BindingSource
                BrandFamilyBS = New BindingSource
                ShoppingCartBS = New BindingSource
                ItemBS = New BindingSource
                OrderBS = New BindingSource
                POHDBS = New BindingSource
                PODTBS = New BindingSource
                PackageItemBS = New BindingSource

                ProductTypeBS.DataSource = DS.Tables(0)
                ProductFamilyBS.DataSource = ProductTypeBS
                ProductFamilyBS.DataMember = "relProducttypefamily"
                ItemBS.DataSource = DS.Tables("Items")
                ListBox1.DataSource = ProductTypeBS
                ListBox1.ValueMember = "producttypeid"
                ListBox1.DisplayMember = "producttypename"
                ShoppingCartBS.DataSource = DS.Tables("ShoppingCart")

                ItemBS.DataSource = DS.Tables("Items")
                OrderBS.DataSource = DS.Tables("Orders")
                OrderBS.Sort = "pohdid desc"
                POHDBS.DataSource = DS.Tables("POHD")
                PODTBS.DataSource = DS.Tables("PODTL")
                PackageItemBS.DataSource = DS.Tables("PackageItem")

                'BS.Sort = "ordernum asc"
                DataGridView1.AutoGenerateColumns = False
                DataGridView1.DataSource = ItemBS
                cm = CType(Me.BindingContext(ItemBS), CurrencyManager)
                CType(DataGridView1.Columns("qtycombobox"), DataGridViewComboBoxColumn).DataSource = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}

                DataGridView2.AutoGenerateColumns = False
                DataGridView2.DataSource = ShoppingCartBS


                cm1 = CType(Me.BindingContext(ShoppingCartBS), CurrencyManager)
                'getTotalShoppingCart will count if item is package. 
                mytotal = getTotalShoppingCart(True)

                showtotal(mytotal)
                'Label8.Text = String.Format("Total Amount: {0:#,##0.00} HKD", mytotal)

                DataGridView3.AutoGenerateColumns = False
                DataGridView3.DataSource = OrderBS
                cm2 = CType(Me.BindingContext(OrderBS), CurrencyManager)
                DataGridView3.DefaultCellStyle.Font = New Font("Microsoft Sans Serif", 8.25, System.Drawing.FontStyle.Regular)
            Else
                MessageBox.Show(mymessage)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub loaddata()
        'populate Data  
        DS = New DataSet

        Dim mymessage As String = ""
        Dim sqlstr As String = String.Empty
        sqlstrSB = New StringBuilder

        'sqlstr = "(select distinct 'Promotional Items' as producttypename,0 as producttypeid from shop.itemprice ip" &
        '        " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
        '        " union all" &
        '        " (select producttypename,producttypeid from shop.producttype order by producttypeid asc);" &
        '        " (select distinct 0 as producttypeid,i.familyid,f.familyname from shop.itemprice ip" &
        '        " left join shop.item i on i.itemid = ip.itempriceid" &
        '        " left join shop.family f on f.familyid = i.familyid" &
        '        " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
        '        " union all" &
        '        " (select distinct i.producttypeid,i.familyid,f.familyname from shop.item i" &
        '        " left join shop.family f on f.familyid = i.familyid" &
        '        " order by producttypeid,f.familyname);" &
        '        " (select distinct 0,0,i.brandid,b.brandname  from shop.item i" &
        '        " left join shop.brand b on b.brandid = i.brandid" &
        '        " left join shop.producttype p on p.producttypeid = i.producttypeid " &
        '        " order by brandname);" &
        '        " select *,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion from shop.shoppingcart sp" &
        '        " left join shop.item i on i.itemid = sp.itemid " &
        '        " left join shop.brand b on b.brandid = i.brandid " &
        '        " left join shop.product p on p.productid = i.productid " &
        '        " left join shop.description d on d.descriptionid = i.descriptionid" &
        '        " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '        " where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ";" &
        '        " select *,1 as quantity,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice  from shop.item i" &
        '        " left join shop.producttype pt on pt.producttypeid = i.producttypeid" &
        '        " left join shop.family f on f.familyid = i.familyid" &
        '        " left join shop.brand  b on b.brandid = i.brandid" &
        '        " left join shop.product  p on p.productid = i.productid" &
        '        " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '        " inner join shop.itemprice ip on ip.itempriceid = i.itemid;" &
        '        " select pd.pohdid,billrefno,orderdate,billingto,e.sn || ' ' || e.givenname as billingtoname,sum(shop.validamount(ph.status,pd.qty,pd.confirmedqty,pd.staffprice)) as totalamount , sum(pd.qty * pd.staffprice) as totalamountold,s.statusname,py.chequenumber,py.bankcode,ph.pohdid from shop.pohd ph" &
        '        " left join shop.employee e on e.employeenumber = ph.billingto" &
        '        " left join shop.status s on s.statusid = ph.status " &
        '        " left join shop.podtl pd on pd.pohdid = ph.pohdid" &
        '        " left join shop.payment py on py.paymentid = ph.pohdid" &
        '        " where ph.employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) &
        '        " group by pd.pohdid,billrefno,orderdate,billingto,billingtoname,statusname,ph.pohdid,py.chequenumber,py.bankcode order by ph.stamp desc;" &
        '        " select * from shop.pohd where pohdid = 0;" &
        '        " select * from shop.podtl where pohdid = 0;" &
        '        " select nvalue from shop.paramhd where paramname = 'Quota Amount';" &
        '        " select  c.cutoff,ph.nvalue from shop.paramhd ph left join shop.cutoff c on c.cutoffid =  ph.nvalue where paramname = 'Current Cutoff';" &
        '        " select nvalue from shop.paramhd where paramname = 'Quota Qty';" &
        '        " select bvalue from shop.paramhd where paramname = 'Quota Default';"

        'sqlstr = "(select distinct 'PROMOTIONAL ITEMS' as producttypename,0 as producttypeid from shop.itemprice ip" &
        '        " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
        '        " union all" &
        '        " (select producttypename,producttypeid from shop.producttype order by producttypeid asc);" &
        '        " (select distinct 0 as producttypeid,i.familyid,f.familyname from shop.itemprice ip" &
        '        " left join shop.item i on i.itemid = ip.itempriceid" &
        '        " left join shop.family f on f.familyid = i.familyid" &
        '        " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
        '        " union all" &
        '        " (select distinct i.producttypeid,i.familyid,f.familyname from shop.item i" &
        '        " left join shop.family f on f.familyid = i.familyid" &
        '        " order by producttypeid,f.familyname);" &
        '        " (select distinct 0,0,i.brandid,b.brandname  from shop.item i" &
        '        " left join shop.brand b on b.brandid = i.brandid" &
        '        " left join shop.producttype p on p.producttypeid = i.producttypeid " &
        '        " order by brandname);" &
        '        " select *,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion from shop.shoppingcart sp" &
        '        " left join shop.item i on i.itemid = sp.itemid " &
        '        " left join shop.brand b on b.brandid = i.brandid " &
        '        " left join shop.product p on p.productid = i.productid " &
        '        " left join shop.description d on d.descriptionid = i.descriptionid" &
        '        " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '        " where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ";" &
        '        " select *,1 as quantity,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice  from shop.item i" &
        '        " left join shop.producttype pt on pt.producttypeid = i.producttypeid" &
        '        " left join shop.family f on f.familyid = i.familyid" &
        '        " left join shop.brand  b on b.brandid = i.brandid" &
        '        " left join shop.product  p on p.productid = i.productid" &
        '        " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '        " inner join shop.itemprice ip on ip.itempriceid = i.itemid;" &
        '        " select pd.pohdid,billrefno,orderdate,billingto,e.sn || ' ' || e.givenname as billingtoname,sum(shop.validamount(ph.status,pd.qty,pd.confirmedqty,pd.staffprice)) as totalamount , sum(pd.qty * pd.staffprice) as totalamountold,s.statusname,py.chequenumber,py.bankcode,ph.pohdid from shop.pohd ph" &
        '        " left join shop.employee e on e.employeenumber = ph.billingto" &
        '        " left join shop.status s on s.statusid = ph.status " &
        '        " left join shop.podtl pd on pd.pohdid = ph.pohdid" &
        '        " left join shop.payment py on py.paymentid = ph.pohdid" &
        '        " where ph.employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) &
        '        " group by pd.pohdid,billrefno,orderdate,billingto,billingtoname,statusname,ph.pohdid,py.chequenumber,py.bankcode order by ph.stamp desc;" &
        '        " select * from shop.pohd where pohdid = 0;" &
        '        " select * from shop.podtl where pohdid = 0;" &
        '        " select nvalue from shop.paramhd where paramname = 'Quota Amount';" &
        '        " select  c.cutoff,ph.nvalue,c.emailnotificationdate,c.chequesubmitdate,c.collectiondate from shop.paramhd ph left join shop.cutoff c on c.cutoffid =  ph.nvalue where paramname = 'Current Cutoff';" &
        '        " select nvalue from shop.paramhd where paramname = 'Quota Qty';" &
        '        " select bvalue from shop.paramhd where paramname = 'Quota Default';" &
        '        " select cvalue from shop.paramhd where paramname = 'Special User';" &
        '        " select * from shop.employee where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ";"

        'sqlstr = "(select distinct 'PROMOTIONAL ITEMS' as producttypename,0 as producttypeid from shop.itemprice ip" &
        '       " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
        '       " union all" &
        '       " (select producttypename,producttypeid from shop.producttype order by producttypeid asc);" &
        '       " (select distinct 0 as producttypeid,i.familyid,f.familyname from shop.itemprice ip" &
        '       " left join shop.item i on i.itemid = ip.itempriceid" &
        '       " left join shop.family f on f.familyid = i.familyid" &
        '       " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
        '       " union all" &
        '       " (select distinct i.producttypeid,i.familyid,f.familyname from shop.item i" &
        '       " left join shop.family f on f.familyid = i.familyid" &
        '       " order by producttypeid,f.familyname);" &
        '       " (select distinct 0,0,i.brandid,b.brandname  from shop.item i" &
        '       " left join shop.brand b on b.brandid = i.brandid" &
        '       " left join shop.producttype p on p.producttypeid = i.producttypeid " &
        '       " order by brandname);" &
        '       " select sp.*,i.*,b.*,p.*,d.*,ip.itempriceid,ip.retailprice,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as staffprice,ip.promotionprice,ip.promotionstartdate,ip.promotionenddate,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice from shop.shoppingcart sp" &
        '       " left join shop.item i on i.itemid = sp.itemid " &
        '       " left join shop.brand b on b.brandid = i.brandid " &
        '       " left join shop.product p on p.productid = i.productid " &
        '       " left join shop.description d on d.descriptionid = i.descriptionid" &
        '       " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '       " where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ";" &
        '       " select *,1 as quantity,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice  from shop.item i" &
        '       " left join shop.producttype pt on pt.producttypeid = i.producttypeid" &
        '       " left join shop.family f on f.familyid = i.familyid" &
        '       " left join shop.brand  b on b.brandid = i.brandid" &
        '       " left join shop.product  p on p.productid = i.productid" &
        '       " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '       " inner join shop.itemprice ip on ip.itempriceid = i.itemid;" &
        '       " select pd.pohdid,billrefno,orderdate,billingto,e.sn || ' ' || e.givenname as billingtoname,sum(shop.validamount(ph.status,pd.qty,pd.confirmedqty,pd.staffprice)) as totalamount , sum(pd.qty * pd.staffprice) as totalamountold,s.statusname,py.chequenumber,py.bankcode,ph.pohdid from shop.pohd ph" &
        '       " left join shop.employee e on e.employeenumber = ph.billingto" &
        '       " left join shop.status s on s.statusid = ph.status " &
        '       " left join shop.podtl pd on pd.pohdid = ph.pohdid" &
        '       " left join shop.payment py on py.paymentid = ph.pohdid" &
        '       " where ph.employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) &
        '       " group by pd.pohdid,billrefno,orderdate,billingto,billingtoname,statusname,ph.pohdid,py.chequenumber,py.bankcode order by ph.stamp desc;" &
        '       " select * from shop.pohd where pohdid = 0;" &
        '       " select * from shop.podtl where pohdid = 0;" &
        '       " select nvalue from shop.paramhd where paramname = 'Quota Amount';" &
        '       " select  c.cutoff,ph.nvalue,c.emailnotificationdate,c.chequesubmitdate,c.collectiondate from shop.paramhd ph left join shop.cutoff c on c.cutoffid =  ph.nvalue where paramname = 'Current Cutoff';" &
        '       " select nvalue from shop.paramhd where paramname = 'Quota Qty';" &
        '       " select bvalue from shop.paramhd where paramname = 'Quota Default';" &
        '       " select cvalue from shop.paramhd where paramname = 'Special User';" &
        '       " select * from shop.employee where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ";" &
        '      " select cvalue from shop.paramhd where paramname = 'imagepath';" &
        '        " select cvalue from shop.paramhd where paramname = 'imageext';"

        ''ProductType
        'sqlstrSB.Append("(select distinct 'PROMOTIONAL ITEMS' as producttypename,0 as producttypeid from shop.itemprice ip" &
        '       " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
        '       " union all" &
        '       " (select producttypename,producttypeid from shop.producttype order by producttypeid asc);")
        ''ProductType With Promo
        sqlstrSB.Append("(select distinct 'PROMOTIONAL ITEMS' as producttypename,0 as producttypeid from shop.itemprice ip" &
              " where shop.validpromotion(promotionstartdate,promotionenddate," & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ",ip.itempriceid))" &
              " union all" &
              " (select producttypename,producttypeid from shop.producttype order by producttypeid asc);")

        'Family
        'sqlstrSB.Append(" (select distinct 0 as producttypeid,i.familyid,f.familyname from shop.itemprice ip" &
        '       " left join shop.item i on i.itemid = ip.itempriceid" &
        '       " left join shop.family f on f.familyid = i.familyid" &
        '       " where ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date)" &
        '       " union all" &
        '       " (select distinct i.producttypeid,i.familyid,f.familyname from shop.item i" &
        '       " left join shop.family f on f.familyid = i.familyid" &
        '       " order by producttypeid,f.familyname);")
        sqlstrSB.Append(" (select distinct 0 as producttypeid,i.familyid,f.familyname from shop.itemprice ip" &
               " left join shop.item i on i.itemid = ip.itempriceid" &
               " left join shop.family f on f.familyid = i.familyid" &
               " where shop.validpromotion(promotionstartdate,promotionenddate," & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ",ip.itempriceid) )" &
               " union all" &
               " (select distinct i.producttypeid,i.familyid,f.familyname from shop.item i" &
               " left join shop.family f on f.familyid = i.familyid" &
               " order by producttypeid,f.familyname);")
        'Brand
        sqlstrSB.Append(" (select distinct 0,0,i.brandid,b.brandname  from shop.item i" &
               " left join shop.brand b on b.brandid = i.brandid" &
               " left join shop.producttype p on p.producttypeid = i.producttypeid " &
               " order by brandname);")
        'Shopping Cart
        'sqlstrSB.Append(" select sp.*,i.*,b.*,p.*,d.*,ip.itempriceid,ip.retailprice,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as staffprice,ip.promotionprice,ip.promotionstartdate,ip.promotionenddate,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice from shop.shoppingcart sp" &
        '       " left join shop.item i on i.itemid = sp.itemid " &
        '       " left join shop.brand b on b.brandid = i.brandid " &
        '       " left join shop.product p on p.productid = i.productid " &
        '       " left join shop.description d on d.descriptionid = i.descriptionid" &
        '       " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '       " where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ";")
        sqlstrSB.Append(" select sp.*,i.*,b.*,p.*,d.*,ip.itempriceid,ip.retailprice,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate," & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ",ip.itempriceid) as staffprice,ip.promotionprice,ip.promotionstartdate,ip.promotionenddate,shop.validpromotion(promotionstartdate,promotionenddate," & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ",ip.itempriceid)  as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate," & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ",ip.itempriceid) as validprice from shop.shoppingcart sp" &
                        " left join shop.item i on i.itemid = sp.itemid " &
                        " left join shop.brand b on b.brandid = i.brandid " &
                        " left join shop.product p on p.productid = i.productid " &
                        " left join shop.description d on d.descriptionid = i.descriptionid" &
                        " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
                        " where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ";")
        'Item
        'sqlstrSB.Append(" select *,1 as quantity,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice  from shop.item i" &
        '       " left join shop.producttype pt on pt.producttypeid = i.producttypeid" &
        '       " left join shop.family f on f.familyid = i.familyid" &
        '       " left join shop.brand  b on b.brandid = i.brandid" &
        '       " left join shop.product  p on p.productid = i.productid" &
        '       " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '       " inner join shop.itemprice ip on ip.itempriceid = i.itemid;")
        'Item
        'sqlstrSB.Append(" (select i.*,pt.*,f.*,b.*,p.*,d.*,ip.*,1 as quantity,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice,null::bigint as packageid" &
        '                " from shop.item i " &
        '                " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
        '                " left join shop.family f on f.familyid = i.familyid " &
        '                " left join shop.brand  b on b.brandid = i.brandid " &
        '                " left join shop.product  p on p.productid = i.productid " &
        '                " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '                " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '                " where not itemid in (select itemid as id from shop.itempackage i " &
        '                " left join shop.packagename p on p.id = i.packageid  where active" &
        '                " except all " &
        '                " select i.id from shop.itemmustbesold i " &
        '                " left join shop.itempackage ip on ip.itemid = i.id" &
        '                " left join shop.packagename p on p.id = ip.packageid  " &
        '                " where p.active)" &
        '                " union all" &
        '                " select i.*,pt.*,f.*,b.*,p.*,d.descriptionid,pn.packagename as descriptionname,foo.mainitem as itempriceid,foo.retailprice,foo.staffprice,foo.promotionprice,null::date as promotionstartdate,null::date promotionenddate,1 as quantity,null::boolean as promotion,foo.validprice,foo.packageid as packageid" &
        '                " from shop.packagename pn" &
        '                " left join shop.item i on i.itemid = pn.mainitem" &
        '                " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
        '                " left join shop.family f on f.familyid = i.familyid " &
        '                " left join shop.brand  b on b.brandid = i.brandid " &
        '                " left join shop.product  p on p.productid = i.productid " &
        '                " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '                " left join (select mainitem,packageid,sum(ip.retailprice) as retailprice,sum(ip.staffprice) as staffprice,sum(ip.promotionprice) as promotionprice,sum(shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate)) as validprice from shop.itempackage i " &
        '                " left join shop.packagename p on p.id = i.packageid" &
        '                " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '                " where active " &
        '                " group by mainitem,packageid) as foo on foo.mainitem = i.itemid" &
        '                " where pn.active);")
        'Item Add Package in Promotional Item
        'sqlstrSB.Append(" (select i.*,pt.*,f.*,b.*,p.*,d.descriptionid,pn.packagename as descriptionname,foo.mainitem as itempriceid,foo.retailprice,foo.staffprice,foo.promotionprice,null::date as promotionstartdate,null::date promotionenddate,1 as quantity,true as promotion,foo.validprice,foo.packageid as packageid" &
        '               " from shop.packagename pn" &
        '               " left join shop.item i on i.itemid = pn.mainitem" &
        '               " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
        '               " left join shop.family f on f.familyid = i.familyid " &
        '               " left join shop.brand  b on b.brandid = i.brandid " &
        '               " left join shop.product  p on p.productid = i.productid " &
        '               " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '               " left join (select mainitem,packageid,(sum(ip.retailprice) ) as retailprice,(sum(ip.staffprice) ) as staffprice,sum(ip.promotionprice) as promotionprice,sum(shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate)) as validprice from shop.itempackage i " &
        '               " left join shop.packagename p on p.id = i.packageid" &
        '               " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '               " where active " &
        '               " group by mainitem,packageid) as foo on foo.mainitem = i.itemid" &
        '               " where pn.active and pn.promotionalitem)" &
        '                " union all" &
        '                "(select i.*,pt.*,f.*,b.*,p.*,d.*,ip.*,1 as quantity,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice,null::bigint as packageid" &
        '               " from shop.item i " &
        '               " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
        '               " left join shop.family f on f.familyid = i.familyid " &
        '               " left join shop.brand  b on b.brandid = i.brandid " &
        '               " left join shop.product  p on p.productid = i.productid " &
        '               " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '               " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '               " where not itemid in (select itemid as id from shop.itempackage i " &
        '               " left join shop.packagename p on p.id = i.packageid  where active" &
        '               " except all " &
        '               " select i.id from shop.itemmustbesold i " &
        '               " left join shop.itempackage ip on ip.itemid = i.id" &
        '               " left join shop.packagename p on p.id = ip.packageid  " &
        '               " where p.active))" &
        '               " union all" &
        '               " (select i.*,pt.*,f.*,b.*,p.*,d.descriptionid,pn.packagename as descriptionname,foo.mainitem as itempriceid,foo.retailprice,foo.staffprice,foo.promotionprice,null::date as promotionstartdate,null::date promotionenddate,1 as quantity,null::boolean as promotion,foo.validprice,foo.packageid as packageid" &
        '               " from shop.packagename pn" &
        '               " left join shop.item i on i.itemid = pn.mainitem" &
        '               " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
        '               " left join shop.family f on f.familyid = i.familyid " &
        '               " left join shop.brand  b on b.brandid = i.brandid " &
        '               " left join shop.product  p on p.productid = i.productid " &
        '               " left join shop.description  d on d.descriptionid = i.descriptionid" &
        '               " left join (select mainitem,packageid,sum(ip.retailprice) as retailprice,sum(ip.staffprice) as staffprice,sum(ip.promotionprice) as promotionprice,sum(shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate)) as validprice from shop.itempackage i " &
        '               " left join shop.packagename p on p.id = i.packageid" &
        '               " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
        '               " where active " &
        '               " group by mainitem,packageid) as foo on foo.mainitem = i.itemid" &
        '               " where pn.active and not pn.promotionalitem);")
        'Item Add Package in Promotional Item Group
        sqlstrSB.Append(" (select i.*,pt.*,f.*,b.*,p.*,d.descriptionid,pn.packagename as descriptionname,foo.mainitem as itempriceid,foo.retailprice,foo.staffprice,foo.promotionprice,null::date as promotionstartdate,null::date promotionenddate,1 as quantity,case promotion when true then true else null end as promotion,foo.validprice,foo.packageid as packageid" &
                      " from shop.packagename pn" &
                      " left join shop.item i on i.itemid = pn.mainitem" &
                      " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
                      " left join shop.family f on f.familyid = i.familyid " &
                      " left join shop.brand  b on b.brandid = i.brandid " &
                      " left join shop.product  p on p.productid = i.productid " &
                      " left join shop.description  d on d.descriptionid = i.descriptionid" &
                      " left join (select mainitem,packageid,(sum(ip.retailprice) ) as retailprice,(sum(ip.staffprice) ) as staffprice,sum(ip.promotionprice) as promotionprice,sum(shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate," & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ",ip.itempriceid)) as validprice , shop.validpromotion(promotionstartdate,promotionenddate," & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ",ip.itempriceid)as promotion  from shop.itempackage i " &
                      " left join shop.packagename p on p.id = i.packageid" &
                      " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
                      " where active " &
                      " group by mainitem,packageid,promotion) as foo on foo.mainitem = i.itemid" &
                      " where pn.active and pn.promotionalitem and promotion and not foo.retailprice isnull)" &
                       " union all" &
                       "(select i.*,pt.*,f.*,b.*,p.*,d.*,ip.*,1 as quantity,shop.validpromotion(promotionstartdate,promotionenddate," & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ",ip.itempriceid) as promotion, shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate," & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ",ip.itempriceid) as validprice,null::bigint as packageid" &
                      " from shop.item i " &
                      " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
                      " left join shop.family f on f.familyid = i.familyid " &
                      " left join shop.brand  b on b.brandid = i.brandid " &
                      " left join shop.product  p on p.productid = i.productid " &
                      " left join shop.description  d on d.descriptionid = i.descriptionid" &
                      " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
                      " where not itemid in (select itemid as id from shop.itempackage i " &
                      " left join shop.packagename p on p.id = i.packageid  where active" &
                      " except all " &
                      " select i.id from shop.itemmustbesold i " &
                      " left join shop.itempackage ip on ip.itemid = i.id" &
                      " left join shop.packagename p on p.id = ip.packageid  " &
                      " where p.active))" &
                      " union all" &
                      " (select i.*,pt.*,f.*,b.*,p.*,d.descriptionid,pn.packagename as descriptionname,foo.mainitem as itempriceid,foo.retailprice,foo.staffprice,foo.promotionprice,null::date as promotionstartdate,null::date promotionenddate,1 as quantity,null::boolean as promotion,foo.validprice,foo.packageid as packageid" &
                      " from shop.packagename pn" &
                      " left join shop.item i on i.itemid = pn.mainitem" &
                      " left join shop.producttype pt on pt.producttypeid = i.producttypeid " &
                      " left join shop.family f on f.familyid = i.familyid " &
                      " left join shop.brand  b on b.brandid = i.brandid " &
                      " left join shop.product  p on p.productid = i.productid " &
                      " left join shop.description  d on d.descriptionid = i.descriptionid" &
                      " left join (select mainitem,packageid,sum(ip.retailprice) as retailprice,sum(ip.staffprice) as staffprice,sum(ip.promotionprice) as promotionprice,sum(shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate," & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ",ip.itempriceid)) as validprice from shop.itempackage i " &
                      " left join shop.packagename p on p.id = i.packageid" &
                      " inner join shop.itemprice ip on ip.itempriceid = i.itemid" &
                      " where active " &
                      " group by mainitem,packageid) as foo on foo.mainitem = i.itemid" &
                      " where pn.active and not pn.promotionalitem and not foo.retailprice isnull);")
                'History
        sqlstrSB.Append(String.Format(" select pd.pohdid,billrefno,orderdate,billingto,e.sn || ' ' || e.givenname as billingtoname,sum(shop.validamount(ph.status,pd.qty,pd.confirmedqty,pd.staffprice)) as totalamount , sum(pd.qty * pd.staffprice) as totalamountold,s.statusname,py.chequenumber,py.bankcode,ph.pohdid from shop.pohd ph" &
               " left join shop.employee e on e.employeenumber = ph.billingto" &
               " left join shop.status s on s.statusid = ph.status " &
               " left join shop.podtl pd on pd.pohdid = ph.pohdid" &
               " left join shop.payment py on py.paymentid = ph.pohdid" &
               " where ph.employeenumber = {0}", DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber)) &
               " group by pd.pohdid,billrefno,orderdate,billingto,billingtoname,statusname,ph.pohdid,py.chequenumber,py.bankcode order by ph.stamp desc;")
        sqlstrSB.Append(" select * from shop.pohd where pohdid = 0;")
        sqlstrSB.Append(" select * from shop.podtl where pohdid = 0;")
        sqlstrSB.Append(" select nvalue from shop.paramhd where paramname = 'Quota Amount';")
        sqlstrSB.Append(" select  c.cutoff,ph.nvalue,c.emailnotificationdate,c.chequesubmitdate,c.collectiondate from shop.paramhd ph left join shop.cutoff c on c.cutoffid =  ph.nvalue where paramname = 'Current Cutoff';")
        sqlstrSB.Append(" select nvalue from shop.paramhd where paramname = 'Quota Qty';")
        sqlstrSB.Append(" select bvalue from shop.paramhd where paramname = 'Quota Default';")
        sqlstrSB.Append(" select cvalue from shop.paramhd where paramname = 'Special User';")
        sqlstrSB.Append(" select * from shop.employee where employeenumber = " & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ";")
        sqlstrSB.Append(" select cvalue from shop.paramhd where paramname = 'imagepath';")
        sqlstrSB.Append(" select cvalue from shop.paramhd where paramname = 'imageext';")
        'sqlstrSB.Append(" select itemid,packageid,p.packagename,ip.retailprice,ip.staffprice,ip.promotionprice,ip.promotionstartdate <= shop.getcurrentcutoff()::date and ip.promotionenddate >= shop.getcurrentcutoff()::date as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate) as validprice from shop.itempackage i " &
        '                " left join shop.packagename p on p.id = i.packageid" &
        '                " inner join shop.itemprice ip on ip.itempriceid = i.itemid where active ;")

        sqlstrSB.Append(" select itemid,packageid,p.packagename,ip.retailprice,ip.staffprice,ip.promotionprice,shop.validpromotion(promotionstartdate,promotionenddate," & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ",ip.itempriceid) as promotion,shop.validprice(staffprice,promotionprice,promotionstartdate,promotionenddate," & DJLib.Dbtools.escapeString(HelperClass1.UserInfo.employeenumber) & ",ip.itempriceid) as validprice from shop.itempackage i " &
                       " left join shop.packagename p on p.id = i.packageid" &
                       " inner join shop.itemprice ip on ip.itempriceid = i.itemid where active ;")
        sqlstrSB.Append("select distinct se.employeenumber,sp.packageid,canbuy  from shop.sellinggroupemployee se" &
                        " left join shop.sellinggrouppackage sp on sp.sellinggroupid = se.sellinggroupid")

        sqlstr = sqlstrSB.ToString
        Try

        
        If DbAdapter1.GetDataSet(sqlstr, DS, mymessage) Then
            DS.Tables(0).TableName = "ProductType"
            DS.Tables(1).TableName = "ProductFamily"
            DS.Tables(2).TableName = "All Brand"
            DS.Tables(3).TableName = "ShoppingCart"
            DS.Tables(4).TableName = "Items"
            DS.Tables(5).TableName = "Orders"
            DS.Tables(6).TableName = "POHD"
            DS.Tables(7).TableName = "PODTL"
            DS.Tables(8).TableName = "Quota Amount"
            DS.Tables(9).TableName = "Current Cutoff"
            DS.Tables(12).TableName = "Special User"
            DS.Tables(13).TableName = "Employee"
            DS.Tables(14).TableName = "imagepath"
            DS.Tables(15).TableName = "imageext"
            DS.Tables(16).TableName = "PackageItem"
            DS.Tables(17).TableName = "SellingGroupEmployee"
            mycutoff = DS.Tables(9).Rows(0).Item("cutoff")
            mycollection = DS.Tables(9).Rows(0).Item("collectiondate")
            mynotification = DS.Tables(9).Rows(0).Item("emailnotificationdate")
            mysubmit = DS.Tables(9).Rows(0).Item("chequesubmitdate")
            imagepath = DS.Tables(14).Rows(0).Item(0)
            imageext = DS.Tables(15).Rows(0).Item(0).ToString.Split(",")

            DS.Tables(6).Columns("pohdid").AutoIncrement = True
            DS.Tables(6).Columns("pohdid").AutoIncrementSeed = -1
            DS.Tables(6).Columns("pohdid").AutoIncrementStep = -1



            Dim idx(0) As DataColumn
            idx(0) = DS.Tables(0).Columns("producttypeid")
            DS.Tables(0).PrimaryKey = idx

            Dim idx1(1) As DataColumn
            idx1(0) = DS.Tables(1).Columns("producttypeid")
            idx1(1) = DS.Tables(1).Columns("familyid")
            DS.Tables(1).PrimaryKey = idx1

            Dim idx3(0) As DataColumn
            idx3(0) = DS.Tables(3).Columns("refno")
            DS.Tables(3).PrimaryKey = idx3

            Dim idx4(0) As DataColumn
            idx4(0) = DS.Tables(4).Columns("itemid")
            DS.Tables(4).PrimaryKey = idx4

            Dim idx6(0) As DataColumn
            idx6(0) = DS.Tables(6).Columns("pohdid")
            DS.Tables(6).PrimaryKey = idx6

            Dim idx17(1) As DataColumn
            idx17(0) = DS.Tables(17).Columns("employeenumber")
            idx17(1) = DS.Tables(17).Columns("packageid")
            DS.Tables(17).PrimaryKey = idx17

            Dim relation As New DataRelation("relProducttypefamily",
                             DS.Tables(0).Columns("producttypeid"),
                             DS.Tables(1).Columns("producttypeid"))
            DS.Relations.Add(relation)

            Dim relation6 As New DataRelation("relPOHDDetail", DS.Tables(6).Columns("pohdid"),
                                              DS.Tables(7).Columns("pohdid"))
            DS.Relations.Add(relation6)

                    'Binding Object


            ProductTypeBS = New BindingSource
            ProductFamilyBS = New BindingSource
            BrandFamilyBS = New BindingSource
            ShoppingCartBS = New BindingSource
            ItemBS = New BindingSource
            OrderBS = New BindingSource
            POHDBS = New BindingSource
            PODTBS = New BindingSource
            PackageItemBS = New BindingSource

            ProductTypeBS.DataSource = DS.Tables(0)
            ProductFamilyBS.DataSource = ProductTypeBS
            ProductFamilyBS.DataMember = "relProducttypefamily"
            ItemBS.DataSource = DS.Tables("Items")
            ListBox1.DataSource = ProductTypeBS
            ListBox1.ValueMember = "producttypeid"
            ListBox1.DisplayMember = "producttypename"
            ShoppingCartBS.DataSource = DS.Tables("ShoppingCart")

            ItemBS.DataSource = DS.Tables("Items")
            OrderBS.DataSource = DS.Tables("Orders")
            OrderBS.Sort = "pohdid desc"
            POHDBS.DataSource = DS.Tables("POHD")
            PODTBS.DataSource = DS.Tables("PODTL")
            PackageItemBS.DataSource = DS.Tables("PackageItem")

            'BS.Sort = "ordernum asc"
            DataGridView1.AutoGenerateColumns = False
            DataGridView1.DataSource = ItemBS
            cm = CType(Me.BindingContext(ItemBS), CurrencyManager)
            CType(DataGridView1.Columns("qtycombobox"), DataGridViewComboBoxColumn).DataSource = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}

            DataGridView2.AutoGenerateColumns = False
            DataGridView2.DataSource = ShoppingCartBS


            cm1 = CType(Me.BindingContext(ShoppingCartBS), CurrencyManager)
            'getTotalShoppingCart will count if item is package. 
            mytotal = getTotalShoppingCart(True)

            showtotal(mytotal)
                    'Label8.Text = String.Format("Total Amount: {0:#,##0.00} HKD", mytotal)

            DataGridView3.AutoGenerateColumns = False
            DataGridView3.DataSource = OrderBS
            cm2 = CType(Me.BindingContext(OrderBS), CurrencyManager)
            DataGridView3.DefaultCellStyle.Font = New Font("Microsoft Sans Serif", 8.25, System.Drawing.FontStyle.Regular)
        Else
            MessageBox.Show(mymessage)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    'Private Sub FormOrder_Load1(ByVal sender As Object, ByVal e As System.EventArgs)
    '    HelperClass1 = New HelperClass
    '    DbAdapter1 = New DbAdapter
    '    Me.Text = GetMenuDesc()
    '    Me.Location = New Point(300, 10)
    '    dbtools1.Userid = DbAdapter1.userid
    '    dbtools1.Password = DbAdapter1.password


    '    For i = 0 To 100
    '        myProduct.Add(New Product With {.brand = String.Format("Brand-{0:000}", i),
    '                                        .description = String.Format("Description-{0:000}", i),
    '                                        .modelno = String.Format("ModelNo-{0:000}", i),
    '                                        .product = String.Format("Product-{0:000}", i),
    '                                        .retailprice = 0.5 + i,
    '                                        .staffprice = 0.1 + i,
    '                                        .qty = 1})

    '    Next

    '    For i = 1 To 5
    '        myOrder.Add(New Order With {.datetime = CDate("2013-02-01 12:00").AddDays(-i),
    '                                    .referenceno = String.Format("1234567{0}", 10 - i),
    '                                    .totalamount = 85.0 + i})
    '    Next


    '    ProductBS.DataSource = myProduct
    '    SelectedBS.DataSource = SelectedProduct
    '    OrderBS.DataSource = myOrder

    '    DataGridView1.AutoGenerateColumns = False
    '    DataGridView2.AutoGenerateColumns = False
    '    DataGridView3.AutoGenerateColumns = False


    '    DataGridView1.DataSource = ProductBS
    '    DataGridView2.DataSource = SelectedBS
    '    DataGridView3.DataSource = OrderBS
    '    '        DataGridView1.Columns("qtycombobox").datasource = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}
    '    CType(DataGridView1.Columns(6), DataGridViewComboBoxColumn).DataSource = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}
    'End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        'DS.Tables(3) -> ShoppingCart
        If e.ColumnIndex = 7 Then
            Me.Validate()
            If DS.Tables("Employee").Rows.Count > 0 Then
                If Not IsDBNull(DS.Tables("Employee").Rows(0).Item("blockenddate")) Then
                    If DS.Tables("Employee").Rows(0).Item("blockenddate") > Today.Date Then
                        MessageBox.Show(String.Format("Sorry, you cannot place an order at this moment. Please come back again after {0:dd-MMM-yyyy}. Thank you.", DS.Tables("Employee").Rows(0).Item("blockenddate")))
                        Exit Sub
                    End If
                End If


            End If
            If DS.Tables(9).Rows(0).Item("cutoff") < Date.Now Then
                'MessageBox.Show(String.Format("You cannot place order after the cutoff time({0:dd-MMM-yyyy HH:mm:ss tt}) until the new cutoff date is set.{1}You may keep your item(s) in your shopping cart for next order.{1}Please come back again later. Thank you.", DS.Tables(9).Rows(0).Item("tvalue"), vbCrLf), "Cutoff Time")
                'MessageBox.Show(String.Format("You cannot place order after the cutoff time({0:dd-MMM-yyyy HH:mm:ss tt}) until the new cutoff date is set.{1}Your shopping cart will be removed.{1}Please come back again later. Thank you.", DS.Tables(9).Rows(0).Item("cutoff"), vbCrLf), "Cutoff Time")
                Dim ordercollectiondate As Date = DS.Tables(9).Rows(0).Item("collectiondate")
                MessageBox.Show(String.Format("You just missed the cut off time. No purchases made at this time.{0}Please come back again on {1:dd-MMM-yyyy}. Thank you.", vbCrLf, ordercollectiondate.AddDays(7)), "Cutoff Time")
                'Delete ShoppingCart
                For Each datarowv As DataRowView In ShoppingCartBS.List
                    ShoppingCartBS.RemoveCurrent()
                Next
                'clear total    
                mytotal = 0
                showtotal(mytotal)
                Exit Sub
            End If

            DataGridView2.DataSource = Nothing
            Dim dr = CType(ItemBS.Current, DataRowView).Row

            'find item if not avail then create
            Dim myobj(0) As Object
            myobj(0) = dr.Item("refno")
            Dim myresult = DS.Tables(3).Rows.Find(myobj)

            Dim newdr As DataRow

            If IsNothing(myresult) Then
                Dim drv As DataRowView = ShoppingCartBS.AddNew()
                newdr = drv.Row
                newdr.Item("qty") = dr.Item("quantity")
                newdr.Item("itemid") = dr.Item("itemid")
                newdr.Item("brandname") = dr.Item("brandname")
                newdr.Item("productname") = dr.Item("productname")
                newdr.Item("refno") = dr.Item("refno")
                newdr.Item("descriptionname") = dr.Item("descriptionname")
                newdr.Item("retailprice") = dr.Item("retailprice")
                newdr.Item("staffprice") = dr.Item("validprice")
                newdr.Item("promotion") = dr.Item("promotion")
                newdr.Item("packageid") = dr.Item("packageid")
                newdr.Item("employeenumber") = HelperClass1.UserInfo.employeenumber
                DS.Tables(3).Rows.Add(newdr)
            Else
                newdr = myresult
                newdr.Item("qty") = newdr.Item("qty") + dr.Item("quantity")

            End If



            'mytotal = mytotal + newdr.Item("qty") * newdr.Item("staffprice")
            'mytotalqty = mytotalqty + newdr.Item("qty")
            mytotal = getTotalAmount(ShoppingCartBS)
            mytotalqty = getTotalQty(ShoppingCartBS)
            mytotalpromotion = getTotalAmountPromotion(ShoppingCartBS)
            mytotalqtypromotion = getTotalQtyPromotion(ShoppingCartBS)
            Label8.Text = String.Format("Total Amount: {0:#,##0.00} HKD", mytotal)
            DataGridView2.DataSource = DS.Tables(3)
        ElseIf e.ColumnIndex = 8 Then
            'Dim mydetail As New FormItemDetails(ItemBS)
            Dim mydetail As New FormItemDetailsImage(ItemBS, imagepath, imageext)
            mydetail.Show()


        End If

    End Sub

    Private Function getTotalAmount(ByVal shoppingcartbs As BindingSource) As Double
        Dim myret As Double = 0
        If Not IsNothing(shoppingcartbs) Then
            For i = 0 To shoppingcartbs.Count - 1
                Dim drv As DataRowView = shoppingcartbs.Item(i)
                Dim dr As DataRow = drv.Row
                myret = myret + dr.Item("qty") * dr.Item("staffprice")
            Next
        End If
        Return myret
    End Function
    Private Function getTotalAmountPromotion(ByVal shoppingcartbs As BindingSource) As Double
        Dim myret As Double = 0
        If Not IsNothing(shoppingcartbs) Then
            For i = 0 To shoppingcartbs.Count - 1
                Dim drv As DataRowView = shoppingcartbs.Item(i)
                Dim dr As DataRow = drv.Row
                If Not IsDBNull(dr.Item("promotion")) Then
                    myret = myret + dr.Item("qty") * dr.Item("staffprice")
                End If
            Next
        End If
        Return myret
    End Function

    Private Function getTotalQty(ByVal shoppingcartbs As BindingSource) As Integer
        Dim myret As Integer = 0
        If Not IsNothing(shoppingcartbs) Then
            For i = 0 To shoppingcartbs.Count - 1
                Dim drv As DataRowView = shoppingcartbs.Item(i)
                Dim dr As DataRow = drv.Row
                myret = myret + dr.Item("qty")
            Next
        End If
        Return myret
    End Function
    Private Function getTotalQtyPromotion(ByVal shoppingcartbs As BindingSource) As Integer
        Dim myret As Integer = 0
        If Not IsNothing(shoppingcartbs) Then
            For i = 0 To shoppingcartbs.Count - 1
                Dim drv As DataRowView = shoppingcartbs.Item(i)
                Dim dr As DataRow = drv.Row
                If Not IsDBNull(dr.Item("promotion")) Then
                    myret = myret + dr.Item("qty")
                End If
            Next
        End If
        Return myret
    End Function
    Private Sub DataGridView1_CellClick1(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        'DS.Tables(3) -> ShoppingCart
        If e.ColumnIndex = 7 Then
            Me.Validate()
            'Check for cutoff



            DataGridView2.DataSource = Nothing
            Dim drv As DataRowView = ShoppingCartBS.AddNew()
            Dim newdr As DataRow = drv.Row

            Dim dr = CType(ItemBS.Current, DataRowView).Row
            newdr.Item("itemid") = dr.Item("itemid")
            newdr.Item("brandname") = dr.Item("brandname")
            newdr.Item("productname") = dr.Item("productname")
            newdr.Item("refno") = dr.Item("refno")
            newdr.Item("descriptionname") = dr.Item("descriptionname")
            newdr.Item("retailprice") = dr.Item("retailprice")
            newdr.Item("staffprice") = dr.Item("validprice")
            newdr.Item("qty") = dr.Item("quantity")
            newdr.Item("employeenumber") = HelperClass1.UserInfo.employeenumber
            DS.Tables(3).Rows.Add(newdr)

            mytotal = mytotal + newdr.Item("qty") * newdr.Item("staffprice")
            mytotalqty = mytotalqty + newdr.Item("qty")

            Label8.Text = String.Format("Total Amount: {0:#,##0.00} HKD", mytotal)
            DataGridView2.DataSource = DS.Tables(3)
        ElseIf e.ColumnIndex = 8 Then
            Dim mydetail As New FormItemDetails(ItemBS)
            mydetail.Show()


        End If

    End Sub
    'Private Sub DataGridView1_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
    '    MessageBox.Show(e.ColumnIndex & " " & e.RowIndex)
    'End Sub


    'Private Sub Button1_Click1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '    If Not IsNothing(SelectedBS.Current) Then
    '        If MessageBox.Show("Delete selected Record(s)?", "Question", System.Windows.Forms.MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
    '            'DS.Tables(0).Rows.Remove(CType(bs.Current, DataRowView).Row)
    '            For Each dsrow In DataGridView2.SelectedRows
    '                mytotal = mytotal - (CType(dsrow, DataGridViewRow).DataGridView.SelectedCells.Item(4).FormattedValue * CType(dsrow, DataGridViewRow).DataGridView.SelectedCells.Item(5).FormattedValue)
    '                Label8.Text = String.Format("Total Amount: {0:#,##0.00} HKD", mytotal)
    '                SelectedBS.RemoveAt(CType(dsrow, DataGridViewRow).Index)
    '            Next
    '        End If
    '    Else
    '        MessageBox.Show("No record to delete.")
    '    End If
    'End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not IsNothing(ShoppingCartBS.Current) Then
            If MessageBox.Show("Delete selected Record(s)?", "Shopping Cart Delete Record(s)", System.Windows.Forms.MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                'DS.Tables(0).Rows.Remove(CType(bs.Current, DataRowView).Row)
                For Each dsrow In DataGridView2.SelectedRows
                    ShoppingCartBS.RemoveAt(CType(dsrow, DataGridViewRow).Index)
                Next
                Label8.Text = String.Format("Total Amount: {0:#,##0.00} HKD", getTotalShoppingCart(True))
            End If
        Else
            MessageBox.Show("No record to delete.", "Shopping Cart Delete Record(s)")
        End If
    End Sub
    Private Sub Button2_ClickOld(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Working with DataSet.Tables("Order"), OrderBS
        'check shoppingcart.
        If Not IsNothing(ShoppingCartBS) Then
            If ShoppingCartBS.Count = 0 Then
                MessageBox.Show("Your Shopping Cart is Empty. Cannot do checkout", "Checkout")
                Exit Sub
            End If
        End If


        If DS.Tables("Employee").Rows.Count > 0 Then
            If Not IsDBNull(DS.Tables("Employee").Rows(0).Item("blockenddate")) Then
                If DS.Tables("Employee").Rows(0).Item("blockenddate") > Today.Date Then
                    MessageBox.Show(String.Format("Sorry, you cannot place an order at this moment. Please come back again after {0:dd-MMM-yyyy}. Thank you.", DS.Tables("Employee").Rows(0).Item("blockenddate")))
                    Exit Sub
                End If
            End If


        End If

        'Check Cutoff Date
        'If order date > cutoffdate then reject
        If DS.Tables(9).Rows(0).Item("cutoff") < Date.Now Then
            'MessageBox.Show(String.Format("You cannot place order after the cutoff time({0:dd-MMM-yyyy HH:mm:ss tt}) until the new cutoff date is set.{1}You may keep your item(s) in your shopping cart for next order.{1}Please come back again later. Thank you.", DS.Tables(9).Rows(0).Item("tvalue"), vbCrLf), "Cutoff Time")
            'MessageBox.Show(String.Format("You cannot place order after the cutoff time({0:dd-MMM-yyyy HH:mm:ss tt}) until the new cutoff date is set.{1}Your shopping cart will be removed.{1}Please come back again later. Thank you.", DS.Tables(9).Rows(0).Item("cutoff"), vbCrLf), "Cutoff Time")
            Dim ordercollectiondate As Date = DS.Tables(9).Rows(0).Item("collectiondate")
            MessageBox.Show(String.Format("You just missed the cut off time. No purchases made at this time.{0}Please come back again on {1:dd-MMM-yyyy}. Thank you.", vbCrLf, ordercollectiondate.AddDays(7)), "Cutoff Time")
            'Delete ShoppingCart
            For Each drv As DataRowView In ShoppingCartBS.List
                ShoppingCartBS.RemoveCurrent()
            Next
            'clear total    
            mytotal = 0
            showtotal(mytotal)
            Exit Sub
        End If

        'check zero price
        Dim zeroprice As Boolean = False
        For Each drv As DataRowView In ShoppingCartBS.List
            Dim dr As DataRow = drv.Row
            If IsNothing(dr.Item("staffprice")) Then
                zeroprice = True
                ShoppingCartBS.RemoveCurrent()
            End If
        Next
        If zeroprice Then
            MessageBox.Show("Some item(s) with null price is removed.")
            Exit Sub
        End If
        Dim billingto As String = HelperClass1.UserInfo.employeenumber
        Dim billingtoname As String = HelperClass1.UserInfo.DisplayName
        If Not IsNothing(ShoppingCartBS) Then
            If ShoppingCartBS.Count > 0 Then

                If Not IsNothing(HelperClass1.UserInfo.title) Then
                    Dim mytitle As String = HelperClass1.UserInfo.title.ToString
                    'If mytitle.ToString.Contains("KEY ACCOUNT MANAGER") Then
                    'If mytitle.ToString.ToLower.Contains("key account") Or DS.Tables(12).Rows(0).item("ayam,ctang").Contains(HelperClass1.UserInfo.userid) Then
                    If mytitle.ToString.ToLower.Contains("key account") Or DS.Tables(12).Rows(0).Item(0).Contains(HelperClass1.UserInfo.userid) Then
                        Dim myFormBillingTo As New FormBillingTo(HelperClass1.UserInfo)
                        If Not myFormBillingTo.ShowDialog() = Windows.Forms.DialogResult.OK Then
                            Exit Sub
                        Else
                            billingto = CType(myFormBillingTo.bs.Current, DataRowView).Row.Item("employeenumber")
                            billingtoname = CType(myFormBillingTo.bs.Current, DataRowView).Row.Item("surname").ToString & " " & CType(myFormBillingTo.bs.Current, DataRowView).Row.Item("givenname").ToString


                        End If
                    End If
                End If
                'save employee
                Dim mymessage As String = String.Empty
                If DbAdapter1.inputupdateemployee(HelperClass1.UserInfo, mymessage) Then
                Else
                    MessageBox.Show(mymessage)
                    Exit Sub
                End If

                'Check Quota
                If DS.Tables(11).Rows(0).Item(0) Then
                    'Quota Limit For Amount
                    Dim totalpurchase As Decimal
                    Dim QuotaLimit As Decimal

                    If Not ValidQuotaAmount(billingto, mytotal - mytotalpromotion, totalpurchase, QuotaLimit) Then
                        MessageBox.Show(String.Format("Employee Quota Amount : {4:#,##0.00} {2}{2}Purchase Order Amount during this year : {1:#,##0.00} {2}New Order Amount: {3:#,##0.00} {2}{2}Total Order Amount: {0:#,##0.00} ", mytotal + totalpurchase, totalpurchase, vbCrLf, mytotal, QuotaLimit), "Quota Amount Over The Limit")
                        Exit Sub
                    End If
                Else
                    Dim totalpurchaseqty As Decimal

                    Dim QuotaLimitqty As Decimal
                    If Not ValidQuotaQty(billingto, mytotalqty - mytotalqtypromotion, totalpurchaseqty, QuotaLimitqty) Then
                        MessageBox.Show(String.Format("Employee Quota Qty: {4:#,##0} {2}{2}Purchase Order Qty during this year : {1:#,##0} {2}New Order Qty: {3:#,##0} {2}{2}Total Order Qty: {0:#,##0} ", mytotalqty + totalpurchaseqty, totalpurchaseqty, vbCrLf, mytotalqty, QuotaLimitqty), "Quota Qty Over The Limit")
                        Exit Sub
                    End If
                End If

                'If MessageBox.Show("Continue to Check out? After this process, you cannot cancel/modify your order.", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = vbCancel Then
                '    Exit Sub
                'End If
                If MessageBox.Show("Continue to Check out?", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = vbCancel Then
                    Exit Sub
                End If
                'create POHD
                Dim pohddrv As DataRowView = POHDBS.AddNew
                Dim pohddr As DataRow = pohddrv.Row
                pohddr.BeginEdit()
                pohddr.Item("orderdate") = Date.Today.Date
                pohddr.Item("employeenumber") = HelperClass1.UserInfo.employeenumber
                pohddr.Item("billingto") = billingto
                pohddr.Item("status") = 1
                pohddr.Item("qtyconfirmation") = False
                pohddr.Item("cutoffid") = DS.Tables(9).Rows(0).Item("nvalue")
                pohddr.EndEdit()
                DS.Tables(6).Rows.Add(pohddr)

                'For Each drv As DataRowView In ShoppingCartBS.List
                Dim i As Integer = 0
                For Each drv As DataRow In DS.Tables(3).Rows
                    'create PODTL, check packageid

                    Dim dr As DataRow = drv
                    Dim obj(0) As Object

                    If Not IsDBNull(dr.Item("packageid")) Then
                        PackageItemBS.Filter = ("packageid = " & dr.Item("packageid"))
                        For Each mydata As DataRowView In PackageItemBS.List
                            'Debug.Print(mydata.Row.Item(0))

                            Dim podtldrv As DataRowView = PODTBS.AddNew
                            Dim podtldr As DataRow = podtldrv.Row
                            podtldr.BeginEdit()
                            podtldr.Item("pohdid") = pohddr.Item("pohdid")
                            podtldr.Item("itemid") = mydata.Row.Item("itemid")
                            podtldr.Item("qty") = dr.Item("qty")
                            podtldr.Item("retailprice") = mydata.Row.Item("retailprice")
                            podtldr.Item("staffprice") = mydata.Row.Item("validprice") 'mydata.Row.Item("staffprice")

                            'podtldr.Item("staffprice") = dr.Item("validprice")
                            podtldr.Item("promotionalitem") = mydata.Row.Item("promotion")
                            podtldr.EndEdit()
                            DS.Tables(7).Rows.Add(podtldr)
                        Next
                    Else
                        Dim podtldrv As DataRowView = PODTBS.AddNew
                        Dim podtldr As DataRow = podtldrv.Row
                        podtldr.BeginEdit()
                        podtldr.Item("pohdid") = pohddr.Item("pohdid")
                        podtldr.Item("itemid") = dr.Item("itemid")
                        podtldr.Item("qty") = dr.Item("qty")
                        podtldr.Item("retailprice") = dr.Item("retailprice")
                        podtldr.Item("staffprice") = dr.Item("staffprice")

                        'podtldr.Item("staffprice") = dr.Item("validprice")
                        podtldr.Item("promotionalitem") = dr.Item("promotion")
                        podtldr.EndEdit()
                        DS.Tables(7).Rows.Add(podtldr)
                    End If



                Next
                'Delete ShoppingCart
                For Each drv As DataRowView In ShoppingCartBS.List
                    ShoppingCartBS.RemoveCurrent()
                Next


                Dim ra As Integer
                'Dim docnumber As String = String.Empty
                Dim ds2 = DS.GetChanges
                'Remove ShoppingCart from DataBase
                Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                If DbAdapter1.ShoppingCartTx(Me, mye) Then
                    DS.Tables(3).AcceptChanges()
                End If
                Dim myrefno As String = String.Empty
                Dim mypohdid As Long = 0
                If DbAdapter1.CreateTx(Me, mye) Then
                    myrefno = mye.dataset.Tables(6).Rows(0).Item("billrefno")
                    mypohdid = mye.dataset.Tables(6).Rows(0).Item("pohdid")
                    Dim myquery = From row As DataRow In DS.Tables(6).Rows
                              Where row.RowState = DataRowState.Added
                    For Each rec In myquery.ToArray
                        rec.Delete()
                    Next
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                Else
                    MessageBox.Show(mye.message)
                    Exit Sub
                End If
                'Save TablePOHD and PODTL return ReferenceNumber from pohd to display

                'DS.AcceptChanges()

                'Add Dataset.Tables("Order")
                Dim myorderview As DataRowView = OrderBS.AddNew
                Dim Orderdr As DataRow = myorderview.Row
                Orderdr.Item("pohdid") = mypohdid
                Orderdr.Item("billrefno") = myrefno
                Orderdr.Item("orderdate") = Date.Today.Date
                Orderdr.Item("billingto") = billingto
                Orderdr.Item("billingtoname") = billingtoname
                Orderdr.Item("totalamount") = mytotal
                Orderdr.Item("statusname") = "New Order"

                DS.Tables(5).Rows.Add(Orderdr)
                DS.Tables(5).AcceptChanges()
                cm2.Position = 0
                mytotal = 0
                mytotalqty = 0
                showtotal(mytotal)


                TabControl1.SelectedTab = TabPage2
                'MessageBox.Show("Your reference number is :    " & myrefno & "." & vbCrLf &
                '                "You will receive order confirmation later. " &
                '                "Please refer your purchasing schedule email notification for more details. " & vbCrLf &
                '                "Thank you for shopping.")
                MessageBox.Show("Your reference number is :    " & myrefno & "." & vbCrLf &
                                "The confirmation order will be sent out on " & Format(mynotification, "dd-MMM-yyyy") &
                                ". Please refer to the purchasing schedule email notification for more details. " & vbCrLf &
                                "Thank you for shopping.")
            Else
                MessageBox.Show("Your Shopping Cart is Empty. Cannot do checkout", "Checkout")
            End If
        Else
            MessageBox.Show("Your Shopping Cart is Empty. Cannot do checkout", "Checkout")
        End If
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim needToLoad As Boolean = False
        'Working with DataSet.Tables("Order"), OrderBS
        'check shoppingcart.
        If Not IsNothing(ShoppingCartBS) Then
            If ShoppingCartBS.Count = 0 Then
                MessageBox.Show("Your Shopping Cart is Empty. Cannot do checkout", "Checkout")
                Exit Sub
            End If
        End If


        If DS.Tables("Employee").Rows.Count > 0 Then
            If Not IsDBNull(DS.Tables("Employee").Rows(0).Item("blockenddate")) Then
                If DS.Tables("Employee").Rows(0).Item("blockenddate") > Today.Date Then
                    MessageBox.Show(String.Format("Sorry, you cannot place an order at this moment. Please come back again after {0:dd-MMM-yyyy}. Thank you.", DS.Tables("Employee").Rows(0).Item("blockenddate")))
                    Exit Sub
                End If
            End If


        End If

        'Check Cutoff Date
        'If order date > cutoffdate then reject
        If DS.Tables(9).Rows(0).Item("cutoff") < Date.Now Then
            'MessageBox.Show(String.Format("You cannot place order after the cutoff time({0:dd-MMM-yyyy HH:mm:ss tt}) until the new cutoff date is set.{1}You may keep your item(s) in your shopping cart for next order.{1}Please come back again later. Thank you.", DS.Tables(9).Rows(0).Item("tvalue"), vbCrLf), "Cutoff Time")
            'MessageBox.Show(String.Format("You cannot place order after the cutoff time({0:dd-MMM-yyyy HH:mm:ss tt}) until the new cutoff date is set.{1}Your shopping cart will be removed.{1}Please come back again later. Thank you.", DS.Tables(9).Rows(0).Item("cutoff"), vbCrLf), "Cutoff Time")
            Dim ordercollectiondate As Date = DS.Tables(9).Rows(0).Item("collectiondate")
            MessageBox.Show(String.Format("You just missed the cut off time. No purchases made at this time.{0}Please come back again on {1:dd-MMM-yyyy}. Thank you.", vbCrLf, ordercollectiondate.AddDays(7)), "Cutoff Time")
            'Delete ShoppingCart
            For Each drv As DataRowView In ShoppingCartBS.List
                ShoppingCartBS.RemoveCurrent()
            Next
            'clear total    
            mytotal = 0
            showtotal(mytotal)
            Exit Sub
        End If

        'check zero price
        Dim zeroprice As Boolean = False
        For Each drv As DataRowView In ShoppingCartBS.List
            Dim dr As DataRow = drv.Row
            If IsNothing(dr.Item("staffprice")) Then
                zeroprice = True
                ShoppingCartBS.RemoveCurrent()
            End If
        Next
        If zeroprice Then
            MessageBox.Show("Some item(s) with null price is removed.")
            Exit Sub
        End If
        Dim billingto As String = HelperClass1.UserInfo.employeenumber
        Dim billingtoname As String = HelperClass1.UserInfo.DisplayName
        If Not IsNothing(ShoppingCartBS) Then
            If ShoppingCartBS.Count > 0 Then

                If Not IsNothing(HelperClass1.UserInfo.title) Then
                    Dim mytitle As String = HelperClass1.UserInfo.title.ToString
                    'If mytitle.ToString.Contains("KEY ACCOUNT MANAGER") Then
                    'If mytitle.ToString.ToLower.Contains("key account") Or DS.Tables(12).Rows(0).item("ayam,ctang").Contains(HelperClass1.UserInfo.userid) Then
                    If mytitle.ToString.ToLower.Contains("key account") Or DS.Tables(12).Rows(0).Item(0).Contains(HelperClass1.UserInfo.userid) Then
                        Dim myFormBillingTo As New FormBillingTo(HelperClass1.UserInfo)
                        If Not myFormBillingTo.ShowDialog() = Windows.Forms.DialogResult.OK Then
                            Exit Sub
                        Else
                            billingto = CType(myFormBillingTo.bs.Current, DataRowView).Row.Item("employeenumber")
                            billingtoname = CType(myFormBillingTo.bs.Current, DataRowView).Row.Item("surname").ToString & " " & CType(myFormBillingTo.bs.Current, DataRowView).Row.Item("givenname").ToString


                        End If
                    End If
                End If
                'save employee
                Dim mymessage As String = String.Empty
                If DbAdapter1.inputupdateemployee(HelperClass1.UserInfo, mymessage) Then
                Else
                    MessageBox.Show(mymessage)
                    Exit Sub
                End If





                'Check Quota
                If DS.Tables(11).Rows(0).Item(0) Then
                    'Quota Limit For Amount
                    Dim totalpurchase As Decimal
                    Dim QuotaLimit As Decimal

                    If Not ValidQuotaAmount(billingto, mytotal - mytotalpromotion, totalpurchase, QuotaLimit) Then
                        MessageBox.Show(String.Format("Employee Quota Amount : {4:#,##0.00} {2}{2}Purchase Order Amount during this year : {1:#,##0.00} {2}New Order Amount: {3:#,##0.00} {2}{2}Total Order Amount: {0:#,##0.00} ", mytotal + totalpurchase, totalpurchase, vbCrLf, mytotal, QuotaLimit), "Quota Amount Over The Limit")
                        Exit Sub
                    End If
                Else
                    Dim totalpurchaseqty As Decimal

                    Dim QuotaLimitqty As Decimal
                    If Not ValidQuotaQty(billingto, mytotalqty - mytotalqtypromotion, totalpurchaseqty, QuotaLimitqty) Then
                        MessageBox.Show(String.Format("Employee Quota Qty: {4:#,##0} {2}{2}Purchase Order Qty during this year : {1:#,##0} {2}New Order Qty: {3:#,##0} {2}{2}Total Order Qty: {0:#,##0} ", mytotalqty + totalpurchaseqty, totalpurchaseqty, vbCrLf, mytotalqty, QuotaLimitqty), "Quota Qty Over The Limit")
                        Exit Sub
                    End If
                End If

                'If MessageBox.Show("Continue to Check out? After this process, you cannot cancel/modify your order.", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = vbCancel Then
                '    Exit Sub
                'End If



                'Check For Valid Package Promo 

                'for each shopping cart
                ' check valid shop.promotion() if not valid mark row error.
                'if HasError the exit
                For Each mydrv As DataRowView In ShoppingCartBS.List
                    mydrv.Row.RowError = ""
                    If Not IsDBNull(mydrv.Item("packageid")) Then

                        'if Selling Group Item -> do checking
                        If DbAdapter1.IsSellingGroup(mydrv.Item("packageid")) Then
                            Dim mycol(1) As Object
                            mycol(0) = billingto
                            mycol(1) = mydrv.Item("packageid")
                            Dim results As DataRow = DS.Tables(17).Rows.Find(mycol)
                            If Not IsNothing(results) Then
                                If Not results.Item("canbuy") Then
                                    'mydrv.Row.RowError = "Sorry, you cannot buy this item."
                                    If MessageBox.Show("This promotional Item is not for you. Click Button OK if you want to buy it with staff price.", "Promo Item", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                                        PackageItemBS.Filter = ("packageid = " & mydrv.Item("packageid"))
                                        Dim mytotalp As Decimal = 0
                                        For Each mydata As DataRowView In PackageItemBS.List
                                            'Debug.Print(mydata.Row.Item(0))
                                            mydata.Row.Item("validprice") = mydata.Row.Item("staffprice")
                                            mydata.Row.Item("promotion") = False
                                            mytotalp = mytotalp + mydata.Row.Item("staffprice")

                                        Next
                                        needToLoad = True
                                        mydrv.Item("staffprice") = mytotalp
                                        DataGridView2.Invalidate()
                                    Else
                                        mydrv.Row.RowError = "Sorry, you cannot buy this item."
                                    End If
                                    'mydrv.Row.RowError = "Sorry, you cannot buy this item."
                                Else
                                    'Check Maxqty
                                    Dim result As Integer = 0
                                    'result = DbAdapter1.isEligibleQty(mydrv.Item("packageid"), mydrv.Item("qty"), billingto, result)
                                    If DbAdapter1.isEligibleQty(mydrv.Item("packageid"), mydrv.Item("qty"), billingto, result) < 0 Then
                                        mydrv.Row.RowError = String.Format("Maximum quantity for this Promotional item is {0}.", result)
                                    End If
                                End If

                            Else
                                'mydrv.Row.RowError = String.Format("This billingto {0} {1}, is not eligible for this promotional items. Please contact Administrator if eligible!", billingto, billingtoname)
                            End If
                        End If

                    End If
                Next

                If DS.Tables(3).HasErrors Then
                    MessageBox.Show("There are some errors. Please check.")
                    Exit Sub
                End If


                If MessageBox.Show("Continue to Check out?", "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = vbCancel Then
                    If needToLoad Then
                        Me.loaddata()
                    End If
                    Exit Sub
                End If
                'create POHD
                mytotal = getTotalShoppingCart(True)

                Dim pohddrv As DataRowView = POHDBS.AddNew
                Dim pohddr As DataRow = pohddrv.Row
                pohddr.BeginEdit()
                pohddr.Item("orderdate") = Date.Today.Date
                pohddr.Item("employeenumber") = HelperClass1.UserInfo.employeenumber
                pohddr.Item("billingto") = billingto
                pohddr.Item("status") = 1
                pohddr.Item("qtyconfirmation") = False
                pohddr.Item("cutoffid") = DS.Tables(9).Rows(0).Item("nvalue")
                pohddr.EndEdit()
                DS.Tables(6).Rows.Add(pohddr)

                'For Each drv As DataRowView In ShoppingCartBS.List
                Dim i As Integer = 0
                For Each drv As DataRow In DS.Tables(3).Rows
                    'create PODTL, check packageid

                    Dim dr As DataRow = drv
                    Dim obj(0) As Object

                    If Not IsDBNull(dr.Item("packageid")) Then
                        PackageItemBS.Filter = ("packageid = " & dr.Item("packageid"))
                        For Each mydata As DataRowView In PackageItemBS.List
                            'Debug.Print(mydata.Row.Item(0))

                            Dim podtldrv As DataRowView = PODTBS.AddNew
                            Dim podtldr As DataRow = podtldrv.Row
                            podtldr.BeginEdit()
                            podtldr.Item("pohdid") = pohddr.Item("pohdid")
                            podtldr.Item("itemid") = mydata.Row.Item("itemid")
                            podtldr.Item("qty") = dr.Item("qty")
                            podtldr.Item("retailprice") = mydata.Row.Item("retailprice")
                            podtldr.Item("staffprice") = mydata.Row.Item("validprice") 'mydata.Row.Item("staffprice")

                            'podtldr.Item("staffprice") = dr.Item("validprice")
                            podtldr.Item("promotionalitem") = mydata.Row.Item("promotion")
                            podtldr.EndEdit()
                            DS.Tables(7).Rows.Add(podtldr)
                        Next
                    Else
                        Dim podtldrv As DataRowView = PODTBS.AddNew
                        Dim podtldr As DataRow = podtldrv.Row
                        podtldr.BeginEdit()
                        podtldr.Item("pohdid") = pohddr.Item("pohdid")
                        podtldr.Item("itemid") = dr.Item("itemid")
                        podtldr.Item("qty") = dr.Item("qty")
                        podtldr.Item("retailprice") = dr.Item("retailprice")
                        podtldr.Item("staffprice") = dr.Item("staffprice")

                        'podtldr.Item("staffprice") = dr.Item("validprice")
                        podtldr.Item("promotionalitem") = dr.Item("promotion")
                        podtldr.EndEdit()
                        DS.Tables(7).Rows.Add(podtldr)
                    End If



                Next
                'Delete ShoppingCart
                For Each drv As DataRowView In ShoppingCartBS.List
                    ShoppingCartBS.RemoveCurrent()
                Next


                Dim ra As Integer
                'Dim docnumber As String = String.Empty
                Dim ds2 = DS.GetChanges
                'Remove ShoppingCart from DataBase
                Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                If DbAdapter1.ShoppingCartTx(Me, mye) Then
                    DS.Tables(3).AcceptChanges()
                End If
                Dim myrefno As String = String.Empty
                Dim mypohdid As Long = 0
                If DbAdapter1.CreateTx(Me, mye) Then
                    myrefno = mye.dataset.Tables(6).Rows(0).Item("billrefno")
                    mypohdid = mye.dataset.Tables(6).Rows(0).Item("pohdid")
                    Dim myquery = From row As DataRow In DS.Tables(6).Rows
                              Where row.RowState = DataRowState.Added
                    For Each rec In myquery.ToArray
                        rec.Delete()
                    Next
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                Else
                    MessageBox.Show(mye.message)
                    Exit Sub
                End If
                'Save TablePOHD and PODTL return ReferenceNumber from pohd to display

                'DS.AcceptChanges()

                'Add Dataset.Tables("Order")
                Dim myorderview As DataRowView = OrderBS.AddNew
                Dim Orderdr As DataRow = myorderview.Row
                Orderdr.Item("pohdid") = mypohdid
                Orderdr.Item("billrefno") = myrefno
                Orderdr.Item("orderdate") = Date.Today.Date
                Orderdr.Item("billingto") = billingto
                Orderdr.Item("billingtoname") = billingtoname

                Orderdr.Item("totalamount") = mytotal
                Orderdr.Item("statusname") = "New Order"

                DS.Tables(5).Rows.Add(Orderdr)
                DS.Tables(5).AcceptChanges()
                cm2.Position = 0
                mytotal = 0
                mytotalqty = 0
                showtotal(mytotal)


                TabControl1.SelectedTab = TabPage2
                'MessageBox.Show("Your reference number is :    " & myrefno & "." & vbCrLf &
                '                "You will receive order confirmation later. " &
                '                "Please refer your purchasing schedule email notification for more details. " & vbCrLf &
                '                "Thank you for shopping.")
                MessageBox.Show("Your reference number is :    " & myrefno & "." & vbCrLf &
                                "The confirmation order will be sent out on " & Format(mynotification, "dd-MMM-yyyy") &
                                ". Please refer to the purchasing schedule email notification for more details. " & vbCrLf &
                                "Thank you for shopping.")
                'Refresh
                If needtoload Then
                    Me.loaddata()
                End If

            Else
                MessageBox.Show("Your Shopping Cart is Empty. Cannot do checkout", "Checkout")
            End If
        Else
            MessageBox.Show("Your Shopping Cart is Empty. Cannot do checkout", "Checkout")
        End If
    End Sub

    Private Sub DataGridView3_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        If e.ColumnIndex = 8 Then
            Dim mydr As DataRow = CType(OrderBS.Current, DataRowView).Row
            If mydr.Item("statusname") = "Completed" Then
                MessageBox.Show("Order is completed. You cannot modify the payment information.")
                Exit Sub
            ElseIf mydr.Item("statusname") = "Rejected Order" Then
                MessageBox.Show("Order is Rejected. You cannot pay for rejected order.")
                Exit Sub
            ElseIf mydr.Item("statusname") = "New Order" Then
                MessageBox.Show("Payment can be done after you received order confirmation.")
                Exit Sub
            ElseIf mydr.Item("statusname") = "Accepted Order" Then
            Else
                MessageBox.Show("Only Accepted Order allowed.")
                Exit Sub
            End If
            Dim mycheck As New FormPayment(OrderBS)
            If mycheck.ShowDialog() = Windows.Forms.DialogResult.OK Then
                DS.Tables(5).AcceptChanges()
            End If
        ElseIf e.ColumnIndex = 7 Then
            Dim pohdid As Long = CType(OrderBS.Current, DataRowView).Row.Item("pohdid")
            Dim mydetail As New FormOrderHistory(OrderBS)
            mydetail.ShowDialog()

        ElseIf e.ColumnIndex = 9 Then 'Cancel order
            Dim mydr As DataRow = CType(OrderBS.Current, DataRowView).Row
            If Not mydr.Item("statusname") = "New Order" Then
                MessageBox.Show("Only New Order allowed")
                Exit Sub
            Else
                If Not DS.Tables(9).Rows(0).Item("cutoff") < Date.Now Then
                    If MessageBox.Show("Do you want to cancel this order?", "Cancel Order", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                        'update this status current row to cancel
                        Dim sqlstr = "update shop.pohd set status = 6 where pohdid = " & mydr.Item("pohdid")
                        Dim mymessage As String = String.Empty
                        Dim ra As Long
                        If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, mymessage) Then
                            MessageBox.Show(mymessage)
                        Else
                            mydr.Item("statusname") = "Cancelled"
                            mydr.Item("totalamount") = "0"
                        End If
                    End If
                Else
                    MessageBox.Show("The cutoff is in progress, you cannot cancel the order. Please contact Administrator!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
                
            End If
        End If

    End Sub

    Private Sub ListBox1_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles ListBox1.DrawItem
        Dim brush As Brush
        Dim drv As DataRowView = ListBox1.Items(e.Index)
        e.DrawBackground()
        Dim myfont As Font = e.Font
        If drv.Row.Item(0).ToString = "PROMOTIONAL ITEMS" Then
            myfont = New Font(e.Font.Name, e.Font.Size, FontStyle.Bold)
            'e.Graphics.FillRectangle(Brushes.LightGreen, e.Bounds)
            'e.Graphics.DrawString(drv.Row.Item(0).ToString, e.Font, Brushes.Red, New System.Drawing.PointF(e.Bounds.X, e.Bounds.Y))
        End If

        If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
            brush = SystemBrushes.HighlightText 'New SolidBrush(clrSelectedText)
            'e.Graphics.FillRectangle(New SolidBrush(clrHighlight), e.Bounds)
            e.Graphics.FillRectangle(SystemBrushes.Highlight, e.Bounds)
            'e.Graphics.DrawString(drv.Row.Item(0).ToString, e.Font, New SolidBrush(clrSelectedText), e.Bounds)
        Else
            If drv.Row.Item(0).ToString = "PROMOTIONAL ITEMS" Then
                brush = Brushes.Red                
            Else
                brush = SystemBrushes.WindowText 'Brushes.Black
            End If

            'e.Graphics.DrawString(drv.Row.Item(0).ToString, e.Font, Brushes.Black, e.Bounds)
        End If


        'e.Graphics.DrawString(drv.Row.Item(0).ToString, e.Font, Brushes.Black, e.Bounds)
        e.Graphics.DrawString(drv.Row.Item(0).ToString, myfont, brush, e.Bounds)
        e.DrawFocusRectangle()
    End Sub




    Private Sub ListBox1_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        CheckedListBox1.DataSource = Nothing
        sb.Clear()
        ItemBS.Filter = ""
        Try
            If IsNumeric(ListBox1.SelectedValue) Then
                'MessageBox.Show(ListBox1.SelectedValue)
                'Find DS table(0) with selectedvalue -> datarow
                Dim myobj(0) As Object
                myobj(0) = ListBox1.SelectedValue
                Dim arrRows() As DataRow
                Dim myresult = DS.Tables(0).Rows.Find(myobj)

                arrRows = myresult.GetChildRows("relProducttypefamily")

                Dim dt As New DataTable
                If Not IsNothing(arrRows) Then

                    dt.Columns.Add("producttypeid", GetType(Integer))
                    dt.Columns.Add("familyid", GetType(Integer))
                    dt.Columns.Add("familyname", GetType(String))
                    'dt.Rows.Add(0, 0, "All Family")
                    For i = 0 To arrRows.Count - 1
                        dt.Rows.Add(arrRows(i).Item(0), arrRows(i).Item(1), arrRows(i).Item(2))
                    Next
                End If
                ProductTypeFilter = "promotion = true"
                If ListBox1.SelectedValue > 0 Then
                    ProductTypeFilter = "producttypeid = " & ListBox1.SelectedValue
                End If

                sb.Append(ProductTypeFilter)

                'assign to checkedlistbox1
                CheckedListBox1.DataSource = Nothing
                CheckedListBox1.DataSource = dt
                Me.CheckedListBox1.DisplayMember = "familyname"
                Me.CheckedListBox1.ValueMember = "familyid"

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    'Private Sub ListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
    '    sb.Clear()
    '    ItemBS.Filter = ""
    '    Try
    '        If IsNumeric(ListBox1.SelectedValue) Then

    '            Dim dv As DataView = TryCast(ProductFamilyBS.List, DataView)
    '            If Not IsNothing(dv) Then
    '                Dim Fdt As New DataTable
    '                Fdt = dv.ToTable(True, {"familyid", "familyname"})
    '                CheckedListBox1.DataSource = Fdt
    '                Me.CheckedListBox1.DisplayMember = "familyname"
    '                Me.CheckedListBox1.ValueMember = "familyid"
    '            End If

    '            ProductTypeFilter = "promotion = true"
    '            If ListBox1.SelectedValue > 0 Then
    '                ProductTypeFilter = "producttypeid = " & ListBox1.SelectedValue
    '            End If

    '            sb.Append(ProductTypeFilter)
    '            sb.Clear()
    '            sb.Append(ProductTypeFilter)
    '            sb.Append(IIf(FamilyFilter.Length > 0, " and ", "") & FamilyFilter)
    '            sb.Append(IIf(BrandFilter.Length > 0, " and ", "") & BrandFilter.ToString)
    '            ItemBS.Filter = sb.ToString
    '        End If
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged

        familysb.Clear()


        If IsNumeric(CheckedListBox1.SelectedIndex) Then
            FamilyFilter = ""
            'If CheckedListBox1.SelectedIndex > 0 Then
            For Each itemchecked In CheckedListBox1.CheckedItems

                FamilyFilter = "familyid = " & CType(itemchecked, DataRowView).Row.Item(1)
                familysb.Append(IIf(familysb.Length > 0, " or ", "(") & FamilyFilter)
            Next
            familysb.Append(IIf(familysb.Length > 0, ")", ""))
            'End If
            sb.Clear()
            sb.Append(ProductTypeFilter)
            sb.Append(IIf(familysb.Length > 0, " and ", "") & familysb.ToString)
        End If

        ItemBS.Filter = sb.ToString
        Dim dt As New DataTable
        Try
            CheckedListBox2.DataSource = Nothing
            CheckedListBox2.Items.Clear()
            DataGridView1.Invalidate()
            Dim dv As DataView = TryCast(ItemBS.List, DataView)
            If Not IsNothing(dv) Then
                dt = dv.ToTable(True, {"brandid", "brandname"})
                CheckedListBox2.DataSource = dt
                Me.CheckedListBox2.DisplayMember = "brandname"
                Me.CheckedListBox2.ValueMember = "brandid"
            Else
                ListBox1.SelectedValue = 0
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckedListBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox2.SelectedIndexChanged
        Dim Brandsb As New StringBuilder
        Brandsb.Clear()


        If IsNumeric(CheckedListBox2.SelectedIndex) Then
            BrandFilter = ""
            Try
                For Each itemchecked In CheckedListBox2.CheckedItems

                    BrandFilter = "brandid = " & CType(itemchecked, DataRowView).Row.Item(0)
                    Brandsb.Append(IIf(Brandsb.Length > 0, " or ", "(") & BrandFilter)
                Next
            Catch ex As Exception

            End Try

            Brandsb.Append(IIf(Brandsb.Length > 0, ")", ""))

            sb.Clear()
            sb.Append(ProductTypeFilter)
            'sb.Append(IIf(FamilyFilter.Length > 0, " and ", "") & FamilyFilter)
            sb.Append(IIf(familysb.Length > 0, " and ", "") & familysb.ToString)
            sb.Append(IIf(Brandsb.Length > 0, " and ", "") & Brandsb.ToString)
        End If

        ItemBS.Filter = sb.ToString
        'Dim dt As New DataTable
        'Try
        '    'CheckedListBox2.DataSource = Nothing
        '    'CheckedListBox2.Items.Clear()
        '    DataGridView1.Invalidate()
        '    Dim dv As DataView = TryCast(ItemBS.List, DataView)
        '    If Not IsNothing(dv) Then
        '        dt = dv.ToTable(True, {"brandid", "brandname"})
        '        CheckedListBox2.DataSource = dt
        '        Me.CheckedListBox2.DisplayMember = "brandname"
        '        Me.CheckedListBox2.ValueMember = "brandid"
        '    Else
        '        ListBox1.SelectedValue = 0
        '    End If

        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim myfilter As New StringBuilder

        If TextBox1.Text <> "" Then
            myfilter.Append("descriptionname like '*" & TextBox1.Text & "*'")
            myfilter.Append(" or refno like '*" & TextBox1.Text & "*'")
            myfilter.Append(" or productname like '*" & TextBox1.Text & "*'")
            myfilter.Append(" or brandname like '*" & TextBox1.Text & "*'")
            ItemBS.Filter = myfilter.ToString
        Else
            sb.Clear()
            sb.Append(ProductTypeFilter)
            sb.Append(IIf(FamilyFilter.Length > 0, " and ", "") & FamilyFilter)
            sb.Append(IIf(BrandFilter.Length > 0, " and ", "") & BrandFilter.ToString)
            ItemBS.Filter = sb.ToString

        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim myfilter As New StringBuilder
        Try
            If TextBox1.Text <> "" Then
                myfilter.Append("descriptionname like '*" & TextBox1.Text.Replace("'", "''") & "*'")
                myfilter.Append(" or refno like '*" & TextBox1.Text.Replace("'", "''") & "*'")
                myfilter.Append(" or productname like '*" & TextBox1.Text.Replace("'", "''") & "*'")
                myfilter.Append(" or brandname like '*" & TextBox1.Text.Replace("'", "''") & "*'")
                ItemBS.Filter = myfilter.ToString

            Else
                sb.Clear()
                sb.Append(ProductTypeFilter)
                sb.Append(IIf(FamilyFilter.Length > 0, " and ", "") & FamilyFilter)
                'sb.Append(IIf(familysb.Length > 0, " and ", "") & familysb.ToString)
                sb.Append(IIf(BrandFilter.Length > 0, " and ", "") & BrandFilter.ToString)
                ItemBS.Filter = sb.ToString

            End If
            ListBox1.Enabled = TextBox1.Text = ""
            CheckedListBox1.Enabled = TextBox1.Text = ""
            CheckedListBox2.Enabled = TextBox1.Text = ""

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.ColumnIndex)
    End Sub

    Private Sub saveShoppingCart()
        Dim myposition = cm.Position
        Me.Validate()
        ShoppingCartBS.EndEdit()

        Dim ds2 = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            'Dim docnumber As String = String.Empty

            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If DbAdapter1.ShoppingCartTx(Me, mye) Then
                'Dim myquery = From row As DataRow In DS.Tables("ShoppingCart").Rows
                '              Where row.RowState = DataRowState.Added

                'For Each rec In myquery.ToArray
                '    rec.Delete()
                'Next
                ''DS.Merge(ds2)
                'DS.AcceptChanges()
                ''BS.Position = myposition
                'MessageBox.Show("Saved!")

                ''LoadData()
            Else
                MessageBox.Show(mye.message)
            End If
        Else
            MessageBox.Show("Nothing to save.")
        End If
    End Sub
    Private Function getTotalShoppingCart1(ByVal amount As Boolean) As Decimal
        mytotal = 0
        mytotalqty = 0
        For Each drv As DataRowView In ShoppingCartBS.List
            Dim dr As DataRow = drv.Row
            Dim packagename As String = String.Empty
            If Not IsDBNull(dr.Item("packageid")) Then
                PackageItemBS.Filter = ("packageid = " & dr.Item("packageid"))
                Dim subtotal As Decimal = 0
                For Each mydata As DataRowView In PackageItemBS.List
                    packagename = mydata.Row.Item("packagename")
                    subtotal = subtotal + (1 * mydata.Row.Item("staffprice"))
                    mytotal = mytotal + (dr.Item("qty") * mydata.Row.Item("staffprice"))
                    mytotalqty = mytotalqty + dr.Item("qty")
                Next
                dr.Item("descriptionname") = packagename
                dr.Item("staffprice") = subtotal
            Else
                packagename = ""
                mytotal = mytotal + (dr.Item("qty") * dr.Item("staffprice"))
                mytotalqty = mytotalqty + dr.Item("qty")
            End If

        Next
        ShoppingCartBS.EndEdit()
        If amount Then
            Return mytotal
        Else
            Return mytotalqty
        End If
    End Function
    Private Function getTotalShoppingCart(ByVal amount As Boolean) As Decimal
        mytotal = 0
        mytotalqty = 0
        For Each drv As DataRowView In ShoppingCartBS.List
            Dim dr As DataRow = drv.Row
            Dim packagename As String = String.Empty
            If Not IsDBNull(dr.Item("packageid")) Then
                PackageItemBS.Filter = ("packageid = " & dr.Item("packageid"))
                Dim subtotal As Decimal = 0
                For Each mydata As DataRowView In PackageItemBS.List
                    packagename = mydata.Row.Item("packagename")
                    'subtotal = subtotal + (1 * mydata.Row.Item("staffprice"))
                    'mytotal = mytotal + (dr.Item("qty") * mydata.Row.Item("staffprice"))
                    subtotal = subtotal + (1 * mydata.Row.Item("validprice"))
                    mytotal = mytotal + (dr.Item("qty") * mydata.Row.Item("validprice"))
                    mytotalqty = mytotalqty + dr.Item("qty")
                Next
                dr.Item("descriptionname") = packagename
                dr.Item("staffprice") = subtotal
            Else
                packagename = ""
                mytotal = mytotal + (dr.Item("qty") * dr.Item("staffprice"))
                mytotalqty = mytotalqty + dr.Item("qty")
            End If
            
        Next
        ShoppingCartBS.EndEdit()
        If amount Then
            Return mytotal
        Else
            Return mytotalqty
        End If
    End Function

    Private Function ValidQuotaAmount(ByVal billingto As String, ByVal mytotal As Decimal, ByRef totalpurchase As Decimal, ByRef QuotaLimit As Decimal) As Boolean

        QuotaLimit = DS.Tables(8).Rows(0).Item("nvalue")

        totalpurchase = DbAdapter1.getTotalPurchase(billingto, Date.Today.Year)
        Dim openingbalance As Double = DbAdapter1.getopeningbalanceamount(billingto, Date.Today.Year)
        totalpurchase = totalpurchase + openingbalance
        Dim myret As Boolean = False
        If QuotaLimit >= totalpurchase + mytotal Then
            myret = True
        End If
        Return myret
    End Function

    Private Function ValidQuotaQty(ByVal billingto As String, ByVal mytotalqty As Integer, ByRef totalpurchaseqty As Decimal, ByRef QuotaLimitqty As Decimal) As Boolean
        QuotaLimitqty = DS.Tables(10).Rows(0).Item("nvalue")
        Dim openingbalance As Integer = DbAdapter1.getopeningbalanceqty(billingto, Date.Today.Year)
        totalpurchaseqty = DbAdapter1.getTotalPurchaseqty(billingto, Date.Today.Year)
        Dim myret As Boolean = False
        totalpurchaseqty = totalpurchaseqty + openingbalance
        If QuotaLimitqty >= totalpurchaseqty + mytotalqty Then
            myret = True
        End If
        Return myret
    End Function

    Private Sub showtotal(ByVal mytotal As Decimal)
        Label8.Text = String.Format("Total Amount: {0:#,##0.00} HKD", mytotal)
    End Sub

    Public Function PreFilterMessage(ByRef m As System.Windows.Forms.Message) As Boolean Implements System.Windows.Forms.IMessageFilter.PreFilterMessage
        If (m.Msg >= &H100 And m.Msg <= &H109) Or (m.Msg >= &H200 And m.Msg <= &H20E) Then
            Timer1.Stop()
            Timer1.Start()
        End If
        Return False
    End Function

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Timer1.Interval = My.Settings.TimerInterval '1800000 '30 * 60 * 1000
        Application.AddMessageFilter(Me)
        ' Add any initialization after the InitializeComponent() call.
        If My.Settings.TimerOn Then
            Timer1.Start()
        End If

    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        'MessageBox.Show("Time is up!")
        TimerClosing = True
        Me.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim myprocess As New Process
        myprocess.StartInfo.FileName = Application.StartupPath & "\doc\e-staffUserGuide.pdf"
        myprocess.Start()
    End Sub



End Class

Public Class Order
    Public Property referenceno As String
    Public Property datetime As DateTime
    Public Property totalamount As Double
End Class

Public Class Orders
    Implements IEnumerable, IEnumerator
    Private _Orders() As Order
    Private _OrderList As List(Of Order)
    Private _Index As Integer
    Private _Icount As Integer

    Public Sub New(ByVal _Orders() As Order)
        If IsNothing(_Orders) Then
        Else
            _Icount = _Orders.Count - 1
            _Index = -1
            Me._Orders = _Orders
        End If
    End Sub

    Public Sub New(ByVal _OrderList As List(Of Order))
        If IsNothing(_OrderList) Then
        Else
            _Icount = _OrderList.Count - 1
            _Index = -1
            Me._Orders = _Orders
        End If
    End Sub
    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return Me
    End Function

    Public ReadOnly Property Current As Object Implements System.Collections.IEnumerator.Current
        Get
            Return _OrderList(_Index)
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements System.Collections.IEnumerator.MoveNext
        If _Index < _Icount Then
            _Index += 1
            MoveNext = True
        Else
            MoveNext = False
        End If
    End Function

    Public Sub Reset() Implements System.Collections.IEnumerator.Reset
        _Index = -1
    End Sub
End Class

Public Class Product
    Public Property brand As String
    Public Property modelno As String
    Public Property product As String
    Public Property description As String
    Public Property retailprice As Double
    Public Property staffprice As Double
    Public Property qty As Integer
End Class

Public Class Products
    Implements IEnumerable, IEnumerator
    Private _Products() As Product
    Private _ProductList As List(Of Product)
    Private _Index As Integer
    Private _Icount As Integer

    Public Sub New(ByVal _Products() As Product)
        If IsNothing(_Products) Then
        Else
            _Icount = _Products.Count - 1
            _Index = -1
            Me._Products = _Products
        End If
    End Sub

    Public Sub New(ByVal _ProductList As List(Of Product))
        If IsNothing(_ProductList) Then
        Else
            _Icount = _ProductList.Count - 1
            _Index = -1
            Me._Products = _Products
        End If
    End Sub

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return Me
    End Function

    Public ReadOnly Property Current As Object Implements System.Collections.IEnumerator.Current
        Get
            Return _ProductList(_Index)
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements System.Collections.IEnumerator.MoveNext
        If _Index < _Icount Then
            _Index += 1
            MoveNext = True
        Else
            MoveNext = False
        End If
    End Function

    Public Sub Reset() Implements System.Collections.IEnumerator.Reset
        _Index = -1
    End Sub
End Class
