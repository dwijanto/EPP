Imports System.Reflection
Imports EPP.PublicClass
Public Class FormAdmin
    Dim myform As New FormMaster

    Public Sub LoadMe() Handles Me.Load
        Me.FormAdmin_Load(Me, New EventArgs)
        AddHandler ImportPriceListToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler CutOffCalendarToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler KAMTableToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler OrderItemToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler EmployeeTableToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler CutOffLogisticsToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportAvailableQtyToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler SendEmailOrderScheduleToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler SystemParameterToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ImportPriceListIFinanceToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ChequeReminder.Click, AddressOf ToolStripMenuItem_Click
        AddHandler SendOrderCollectionEmailToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler SetCurrentCutoffToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ChequePaymentToolStripMenuItem1.Click, AddressOf ToolStripMenuItem_Click
        AddHandler TransactionReportToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler CheckDepositReportToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler GenerateIFinanceOrdersToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler MasterItemToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler MasterFamilyToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler MasterProductToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler MasterProductTypeToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler MasterDescriptionToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler ItemMustBeSoldToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
        AddHandler PackageItemToolStripMenuItem1.Click, AddressOf ToolStripMenuItem_Click
        AddHandler SellingGroupToolStripMenuItem.Click, AddressOf ToolStripMenuItem_Click
    End Sub
    Public Function GetMenuDesc() As String
        Label1.Text = "Welcome, " & HelperClass1.UserInfo.DisplayName
        Return "App.Version: " & My.Application.Info.Version.ToString & " :: Server: " & DbAdapter1.ConnectionStringDict.Item("HOST") & ", Database: " & DbAdapter1.ConnectionStringDict.Item("DATABASE") & ", Userid: " & HelperClass1.UserId

    End Function
    Private Sub FormAdmin_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HelperClass1 = New HelperClass
            DbAdapter1 = New DbAdapter
            Me.Text = GetMenuDesc()
            Me.Location = New Point(300, 10)
            dbtools1.Userid = DbAdapter1.userid
            dbtools1.Password = DbAdapter1.password

            'Record Login
            Try
                loglogin(DbAdapter1.userid)
            Catch ex As Exception
            End Try


        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Close()
        End Try
    End Sub
    Private Sub loglogin(ByVal userid As String)
        Dim applicationname As String = "eStaff Admin"
        Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        Dim computername As String = My.Computer.Name
        Dim time_stamp As DateTime = Now
        DbAdapter1.loglogin(applicationname, userid, username, computername, time_stamp)
    End Sub

    Private Sub ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim ctrl As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        Dim assembly1 As Assembly = Assembly.GetAssembly(GetType(FormAdmin))
        Dim frm As Form = CType(assembly1.CreateInstance(assembly1.GetName.Name.ToString & "." & ctrl.Tag.ToString, True), Form)
        Dim inMemory As Boolean = False
        For i = 0 To My.Application.OpenForms.Count - 1
            If My.Application.OpenForms.Item(i).Name = frm.Name Then
                ExecuteForm(My.Application.OpenForms.Item(i))
                inMemory = True
            End If
        Next
        If Not inMemory Then
            ExecuteForm(frm)
        End If
    End Sub

    Private Sub ExecuteForm(ByVal obj As Windows.Forms.Form)
        With obj
            .WindowState = FormWindowState.Normal
            .StartPosition = FormStartPosition.CenterScreen
            .Show()
            .Focus()
        End With
    End Sub

    Private Sub FormMenu_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not e.CloseReason = CloseReason.ApplicationExitCall Then
            If MessageBox.Show("Are you sure?", "Exit", MessageBoxButtons.OKCancel) = DialogResult.OK Then
                Me.CloseOpenForm()
                HelperClass1.fadeout(Me)
                DbAdapter1.Dispose()
                HelperClass1.Dispose()
            Else
                e.Cancel = True
            End If
        End If
    End Sub
    Private Sub CloseOpenForm()
        For i = 1 To (My.Application.OpenForms.Count - 1)
            My.Application.OpenForms.Item(1).Close()
        Next
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    'Private Sub SystemParameterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SystemParameterToolStripMenuItem.Click
    '    Dim myform As New FormSystemParameter
    '    myform.Show()
    'End Sub

    Private Sub MasterBrandToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MasterBrandToolStripMenuItem.Click
        Dim sqlstr = "Select brandid,brandname from shop.brand order by brandname;"
        Dim InitDataGridview As New Action(AddressOf Me.initDataGridviewBrand)
        Dim myDbAdapterTx As New dbAdapterTx(AddressOf DbAdapter1.BrandTx)
        Dim myUpdateAction As New Action(AddressOf Me.UpdateBrand)
        Dim myAddAction As New Action(AddressOf Me.AddBrand)
        myform = New FormMaster With {.Sqlstr = sqlstr,
                                      .TableName = "Brand",
                                      .FieldId = "brandid",
                                      .InitDataGridView = InitDataGridview,
                                      .UpdateAction = myUpdateAction,
                                      .AddAction = myAddAction,
                                      .myDbAdapterTx = myDbAdapterTx
                                     }
        myform.Show()
    End Sub
    Private Function initDataGridviewBrand(ByRef sender As Object, ByRef e As System.EventArgs) As Boolean
        Dim myobj = CType(sender, DataGridView)
        myform.DS.Tables(0).TableName = "brandid"
        Dim idx(0) As DataColumn
        idx(0) = myform.DS.Tables(0).Columns("brandid") 'familyid
        myform.DS.Tables(0).PrimaryKey = idx

        myform.DS.Tables(0).Columns("brandid").AutoIncrement = True
        myform.DS.Tables(0).Columns("brandid").AutoIncrementSeed = -1
        myform.DS.Tables(0).Columns("brandid").AutoIncrementStep = -1

        'Binding Object

        myform.BS = New BindingSource

        myform.BS.DataSource = myform.DS.Tables(0)

        'BS.Sort = "ordernum asc"
        myobj.AutoGenerateColumns = False
        myobj.DataSource = myform.BS
        myobj.Columns(0).DataPropertyName = "brandid"
        myobj.Columns(1).DataPropertyName = "brandname"
        myobj.Columns(0).HeaderText = "Brand Id"
        myobj.Columns(1).HeaderText = "Brand Name"
        myform.CM = CType(Me.BindingContext(myform.BS), CurrencyManager)
        Return True
    End Function

    Private Function UpdateBrand(ByRef sender As Object, ByRef e As EventArgs) As Boolean
        Dim myformInput = New FormInputBrand(myform.BS)

        If Not myformInput.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            myform.BS.CancelEdit()
        End If
        myformInput.Dispose()
        Return True
    End Function

    Private Function AddBrand(ByRef sender As Object, ByRef e As EventArgs) As Boolean
        myform.BS.Sort = ""

        Dim drv As DataRowView = myform.BS.AddNew()
        Dim dr = drv.Row

        myform.DS.Tables(0).Rows.Add(dr)
        Dim myforminput = New FormInputBrand(myform.BS)
        If Not myforminput.ShowDialog() = Windows.Forms.DialogResult.OK Then
            myform.DS.Tables(0).Rows.Remove(dr)
        End If
        myforminput.Dispose()
        Return True
    End Function

    


    Private Sub CheckDepositReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckDepositReportToolStripMenuItem.Click

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        My.Settings.TimerOn = False
    End Sub

    Private Sub ItemMustBeSoldToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ItemMustBeSoldToolStripMenuItem.Click

    End Sub


    Private Sub SellingGroupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SellingGroupToolStripMenuItem.Click

    End Sub


    Private Sub ImportPriceListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportPriceListToolStripMenuItem.Click

    End Sub


    Private Sub SendEmailOrderScheduleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SendEmailOrderScheduleToolStripMenuItem.Click

    End Sub
End Class