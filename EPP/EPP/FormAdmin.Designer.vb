<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormAdmin
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormAdmin))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CheckDepositReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GenerateIFinanceOrdersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TransactionReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ActionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OrderPreparationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportPriceListToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SetCurrentCutoffToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SendEmailOrderScheduleToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CutOffLogisticsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportAvailableQtyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ChequeReminder = New System.Windows.Forms.ToolStripMenuItem()
        Me.SendOrderCollectionEmailToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportPriceListIFinanceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CutOffCalendarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SystemParameterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EmployeeTableToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.KAMTableToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterItemToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterProductTypeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterFamilyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterBrandToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterProductToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterDescriptionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PackageItemToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PackageItemToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ItemMustBeSoldToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SellingGroupToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OrderToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OrderItemToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ChequePaymentToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReportToolStripMenuItem, Me.ActionToolStripMenuItem, Me.ToolsToolStripMenuItem, Me.ExitToolStripMenuItem, Me.OrderToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(643, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ReportToolStripMenuItem
        '
        Me.ReportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CheckDepositReportToolStripMenuItem, Me.GenerateIFinanceOrdersToolStripMenuItem, Me.TransactionReportToolStripMenuItem})
        Me.ReportToolStripMenuItem.Name = "ReportToolStripMenuItem"
        Me.ReportToolStripMenuItem.Size = New System.Drawing.Size(59, 20)
        Me.ReportToolStripMenuItem.Text = "Reports"
        '
        'CheckDepositReportToolStripMenuItem
        '
        Me.CheckDepositReportToolStripMenuItem.Name = "CheckDepositReportToolStripMenuItem"
        Me.CheckDepositReportToolStripMenuItem.Size = New System.Drawing.Size(211, 22)
        Me.CheckDepositReportToolStripMenuItem.Tag = "FormChequeDepositReport"
        Me.CheckDepositReportToolStripMenuItem.Text = "Cheque Deposit Report"
        '
        'GenerateIFinanceOrdersToolStripMenuItem
        '
        Me.GenerateIFinanceOrdersToolStripMenuItem.Name = "GenerateIFinanceOrdersToolStripMenuItem"
        Me.GenerateIFinanceOrdersToolStripMenuItem.Size = New System.Drawing.Size(211, 22)
        Me.GenerateIFinanceOrdersToolStripMenuItem.Tag = "FormIFinanceOrder"
        Me.GenerateIFinanceOrdersToolStripMenuItem.Text = "Generate I-Finance Orders"
        '
        'TransactionReportToolStripMenuItem
        '
        Me.TransactionReportToolStripMenuItem.Name = "TransactionReportToolStripMenuItem"
        Me.TransactionReportToolStripMenuItem.Size = New System.Drawing.Size(211, 22)
        Me.TransactionReportToolStripMenuItem.Tag = "FormTxReport"
        Me.TransactionReportToolStripMenuItem.Text = "Transaction Report"
        '
        'ActionToolStripMenuItem
        '
        Me.ActionToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OrderPreparationToolStripMenuItem, Me.CutOffLogisticsToolStripMenuItem, Me.ImportAvailableQtyToolStripMenuItem, Me.ChequeReminder, Me.SendOrderCollectionEmailToolStripMenuItem, Me.ImportPriceListIFinanceToolStripMenuItem})
        Me.ActionToolStripMenuItem.Name = "ActionToolStripMenuItem"
        Me.ActionToolStripMenuItem.Size = New System.Drawing.Size(59, 20)
        Me.ActionToolStripMenuItem.Text = "Actions"
        '
        'OrderPreparationToolStripMenuItem
        '
        Me.OrderPreparationToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImportPriceListToolStripMenuItem, Me.SetCurrentCutoffToolStripMenuItem, Me.SendEmailOrderScheduleToolStripMenuItem})
        Me.OrderPreparationToolStripMenuItem.Name = "OrderPreparationToolStripMenuItem"
        Me.OrderPreparationToolStripMenuItem.Size = New System.Drawing.Size(372, 22)
        Me.OrderPreparationToolStripMenuItem.Text = "Order Preparation"
        '
        'ImportPriceListToolStripMenuItem
        '
        Me.ImportPriceListToolStripMenuItem.Name = "ImportPriceListToolStripMenuItem"
        Me.ImportPriceListToolStripMenuItem.Size = New System.Drawing.Size(223, 22)
        Me.ImportPriceListToolStripMenuItem.Tag = "FormImportPrice"
        Me.ImportPriceListToolStripMenuItem.Text = "Import Price List"
        '
        'SetCurrentCutoffToolStripMenuItem
        '
        Me.SetCurrentCutoffToolStripMenuItem.Name = "SetCurrentCutoffToolStripMenuItem"
        Me.SetCurrentCutoffToolStripMenuItem.Size = New System.Drawing.Size(223, 22)
        Me.SetCurrentCutoffToolStripMenuItem.Tag = "FormSetCurrentCutoff"
        Me.SetCurrentCutoffToolStripMenuItem.Text = "Set Current Cutoff"
        '
        'SendEmailOrderScheduleToolStripMenuItem
        '
        Me.SendEmailOrderScheduleToolStripMenuItem.Name = "SendEmailOrderScheduleToolStripMenuItem"
        Me.SendEmailOrderScheduleToolStripMenuItem.Size = New System.Drawing.Size(223, 22)
        Me.SendEmailOrderScheduleToolStripMenuItem.Tag = "FormSendEmailOrderScheduled"
        Me.SendEmailOrderScheduleToolStripMenuItem.Text = "Send Email Order Scheduled"
        '
        'CutOffLogisticsToolStripMenuItem
        '
        Me.CutOffLogisticsToolStripMenuItem.Name = "CutOffLogisticsToolStripMenuItem"
        Me.CutOffLogisticsToolStripMenuItem.Size = New System.Drawing.Size(372, 22)
        Me.CutOffLogisticsToolStripMenuItem.Tag = "FormSendQtyManually"
        Me.CutOffLogisticsToolStripMenuItem.Text = "Cut Off -> Send Qty to GSHK Logistics"
        '
        'ImportAvailableQtyToolStripMenuItem
        '
        Me.ImportAvailableQtyToolStripMenuItem.Name = "ImportAvailableQtyToolStripMenuItem"
        Me.ImportAvailableQtyToolStripMenuItem.Size = New System.Drawing.Size(372, 22)
        Me.ImportAvailableQtyToolStripMenuItem.Tag = "FormImportAvailableQty"
        Me.ImportAvailableQtyToolStripMenuItem.Text = "Import Available Qty && Send Order Confirmation"
        '
        'ChequeReminder
        '
        Me.ChequeReminder.Name = "ChequeReminder"
        Me.ChequeReminder.Size = New System.Drawing.Size(372, 22)
        Me.ChequeReminder.Tag = "FormChequeReminder"
        Me.ChequeReminder.Text = "Send Payment Reminder Email Notification to  customer"
        '
        'SendOrderCollectionEmailToolStripMenuItem
        '
        Me.SendOrderCollectionEmailToolStripMenuItem.Name = "SendOrderCollectionEmailToolStripMenuItem"
        Me.SendOrderCollectionEmailToolStripMenuItem.Size = New System.Drawing.Size(372, 22)
        Me.SendOrderCollectionEmailToolStripMenuItem.Tag = "FormOrderCollection"
        Me.SendOrderCollectionEmailToolStripMenuItem.Text = "Send Order Collection Email"
        '
        'ImportPriceListIFinanceToolStripMenuItem
        '
        Me.ImportPriceListIFinanceToolStripMenuItem.Name = "ImportPriceListIFinanceToolStripMenuItem"
        Me.ImportPriceListIFinanceToolStripMenuItem.Size = New System.Drawing.Size(372, 22)
        Me.ImportPriceListIFinanceToolStripMenuItem.Tag = "FormImportPrice2"
        Me.ImportPriceListIFinanceToolStripMenuItem.Text = "Import Price List I-Finance"
        Me.ImportPriceListIFinanceToolStripMenuItem.Visible = False
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CutOffCalendarToolStripMenuItem, Me.SystemParameterToolStripMenuItem, Me.EmployeeTableToolStripMenuItem, Me.KAMTableToolStripMenuItem, Me.MasterToolStripMenuItem, Me.PackageItemToolStripMenuItem, Me.SellingGroupToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(78, 20)
        Me.ToolsToolStripMenuItem.Text = "Parameters"
        '
        'CutOffCalendarToolStripMenuItem
        '
        Me.CutOffCalendarToolStripMenuItem.Name = "CutOffCalendarToolStripMenuItem"
        Me.CutOffCalendarToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.CutOffCalendarToolStripMenuItem.Tag = "FormCutoff"
        Me.CutOffCalendarToolStripMenuItem.Text = "Cutoff Table"
        '
        'SystemParameterToolStripMenuItem
        '
        Me.SystemParameterToolStripMenuItem.Name = "SystemParameterToolStripMenuItem"
        Me.SystemParameterToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.SystemParameterToolStripMenuItem.Tag = "FormSystemParameter"
        Me.SystemParameterToolStripMenuItem.Text = "System Parameter"
        '
        'EmployeeTableToolStripMenuItem
        '
        Me.EmployeeTableToolStripMenuItem.Name = "EmployeeTableToolStripMenuItem"
        Me.EmployeeTableToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.EmployeeTableToolStripMenuItem.Tag = "FormEmployee"
        Me.EmployeeTableToolStripMenuItem.Text = "Employee Table"
        '
        'KAMTableToolStripMenuItem
        '
        Me.KAMTableToolStripMenuItem.Name = "KAMTableToolStripMenuItem"
        Me.KAMTableToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.KAMTableToolStripMenuItem.Tag = "FormKAM"
        Me.KAMTableToolStripMenuItem.Text = "KAM Table"
        Me.KAMTableToolStripMenuItem.Visible = False
        '
        'MasterToolStripMenuItem
        '
        Me.MasterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MasterItemToolStripMenuItem, Me.MasterProductTypeToolStripMenuItem, Me.MasterFamilyToolStripMenuItem, Me.MasterBrandToolStripMenuItem, Me.MasterProductToolStripMenuItem, Me.MasterDescriptionToolStripMenuItem})
        Me.MasterToolStripMenuItem.Name = "MasterToolStripMenuItem"
        Me.MasterToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.MasterToolStripMenuItem.Text = "Master"
        '
        'MasterItemToolStripMenuItem
        '
        Me.MasterItemToolStripMenuItem.Name = "MasterItemToolStripMenuItem"
        Me.MasterItemToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
        Me.MasterItemToolStripMenuItem.Tag = "FormMasterItem"
        Me.MasterItemToolStripMenuItem.Text = "Master Item"
        '
        'MasterProductTypeToolStripMenuItem
        '
        Me.MasterProductTypeToolStripMenuItem.Name = "MasterProductTypeToolStripMenuItem"
        Me.MasterProductTypeToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
        Me.MasterProductTypeToolStripMenuItem.Tag = "FormMasterProductType"
        Me.MasterProductTypeToolStripMenuItem.Text = "Master Product Type"
        '
        'MasterFamilyToolStripMenuItem
        '
        Me.MasterFamilyToolStripMenuItem.Name = "MasterFamilyToolStripMenuItem"
        Me.MasterFamilyToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
        Me.MasterFamilyToolStripMenuItem.Tag = "FormMasterFamily"
        Me.MasterFamilyToolStripMenuItem.Text = "Master Family"
        '
        'MasterBrandToolStripMenuItem
        '
        Me.MasterBrandToolStripMenuItem.Name = "MasterBrandToolStripMenuItem"
        Me.MasterBrandToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
        Me.MasterBrandToolStripMenuItem.Tag = ""
        Me.MasterBrandToolStripMenuItem.Text = "Master Brand"
        '
        'MasterProductToolStripMenuItem
        '
        Me.MasterProductToolStripMenuItem.Name = "MasterProductToolStripMenuItem"
        Me.MasterProductToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
        Me.MasterProductToolStripMenuItem.Tag = "FormMasterProduct"
        Me.MasterProductToolStripMenuItem.Text = "Master Product"
        '
        'MasterDescriptionToolStripMenuItem
        '
        Me.MasterDescriptionToolStripMenuItem.Name = "MasterDescriptionToolStripMenuItem"
        Me.MasterDescriptionToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
        Me.MasterDescriptionToolStripMenuItem.Tag = "FormMasterDescription"
        Me.MasterDescriptionToolStripMenuItem.Text = "Master Description"
        '
        'PackageItemToolStripMenuItem
        '
        Me.PackageItemToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PackageItemToolStripMenuItem1, Me.ItemMustBeSoldToolStripMenuItem})
        Me.PackageItemToolStripMenuItem.Name = "PackageItemToolStripMenuItem"
        Me.PackageItemToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.PackageItemToolStripMenuItem.Text = "Package Item"
        '
        'PackageItemToolStripMenuItem1
        '
        Me.PackageItemToolStripMenuItem1.Name = "PackageItemToolStripMenuItem1"
        Me.PackageItemToolStripMenuItem1.Size = New System.Drawing.Size(217, 22)
        Me.PackageItemToolStripMenuItem1.Tag = "FormPackageItem3"
        Me.PackageItemToolStripMenuItem1.Text = "Package Item"
        '
        'ItemMustBeSoldToolStripMenuItem
        '
        Me.ItemMustBeSoldToolStripMenuItem.Name = "ItemMustBeSoldToolStripMenuItem"
        Me.ItemMustBeSoldToolStripMenuItem.Size = New System.Drawing.Size(217, 22)
        Me.ItemMustBeSoldToolStripMenuItem.Tag = "FormItemMustBeSold"
        Me.ItemMustBeSoldToolStripMenuItem.Text = "Item can be sold separately"
        '
        'SellingGroupToolStripMenuItem
        '
        Me.SellingGroupToolStripMenuItem.Name = "SellingGroupToolStripMenuItem"
        Me.SellingGroupToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.SellingGroupToolStripMenuItem.Tag = "FormSellingGroup"
        Me.SellingGroupToolStripMenuItem.Text = "Selling Group"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'OrderToolStripMenuItem
        '
        Me.OrderToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OrderItemToolStripMenuItem, Me.ChequePaymentToolStripMenuItem1})
        Me.OrderToolStripMenuItem.Name = "OrderToolStripMenuItem"
        Me.OrderToolStripMenuItem.Size = New System.Drawing.Size(86, 20)
        Me.OrderToolStripMenuItem.Text = "Transactions"
        '
        'OrderItemToolStripMenuItem
        '
        Me.OrderItemToolStripMenuItem.Name = "OrderItemToolStripMenuItem"
        Me.OrderItemToolStripMenuItem.Size = New System.Drawing.Size(280, 22)
        Me.OrderItemToolStripMenuItem.Tag = "FormOrder"
        Me.OrderItemToolStripMenuItem.Text = "Order Item"
        '
        'ChequePaymentToolStripMenuItem1
        '
        Me.ChequePaymentToolStripMenuItem1.Name = "ChequePaymentToolStripMenuItem1"
        Me.ChequePaymentToolStripMenuItem1.Size = New System.Drawing.Size(280, 22)
        Me.ChequePaymentToolStripMenuItem1.Tag = "FormChequePayment"
        Me.ChequePaymentToolStripMenuItem1.Text = "Cheque Payment && Order Cancellation"
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(410, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Label1"
        '
        'FormAdmin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(643, 108)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "FormAdmin"
        Me.Text = "e-Staff Purchase Program - Admin"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ActionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CutOffLogisticsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportAvailableQtyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ChequeReminder As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SystemParameterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CutOffCalendarToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CheckDepositReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents KAMTableToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OrderToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OrderItemToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents GenerateIFinanceOrdersToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EmployeeTableToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportPriceListIFinanceToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SendOrderCollectionEmailToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ChequePaymentToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OrderPreparationToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportPriceListToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SetCurrentCutoffToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SendEmailOrderScheduleToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TransactionReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterItemToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterProductTypeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterFamilyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterBrandToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterProductToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterDescriptionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PackageItemToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PackageItemToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ItemMustBeSoldToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SellingGroupToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
