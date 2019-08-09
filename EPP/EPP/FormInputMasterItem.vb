Public Class FormInputMasterItem
    Private myrow As DataRowView
    Private bs As New BindingSource

    Private ProductTypeBS As New BindingSource
    Private Familybs As New BindingSource
    Private Brandbs As New BindingSource
    Private Descriptionbs As New BindingSource
    Private ProductBS As New BindingSource

    'Public Sub New(ByRef bs As BindingSource)
    Private ds As DataSet
    Public Sub New(ByRef DS As DataSet, ByRef bs As BindingSource)
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        'bs.DataSource = DS.Tables(0)
        myrow = bs.Current
        Me.bs = bs
        ProductTypeBS.DataSource = DS.Tables(1)
        Familybs.DataSource = DS.Tables(2)
        Brandbs.DataSource = DS.Tables(3)
        Descriptionbs.DataSource = DS.Tables(4)
        ProductBS.DataSource = DS.Tables(6)


        ComboBox1.DisplayMember = "producttypename"
        ComboBox1.ValueMember = "producttypeid"
        ComboBox1.DataSource = ProductTypeBS

        ComboBox2.DisplayMember = "familyname"
        ComboBox2.ValueMember = "familyid"
        ComboBox2.DataSource = Familybs

        ComboBox3.DisplayMember = "brandname"
        ComboBox3.ValueMember = "brandid"
        ComboBox3.DataSource = Brandbs

        ComboBox4.DisplayMember = "descriptionname"
        ComboBox4.ValueMember = "descriptionid"
        ComboBox4.DataSource = Descriptionbs

        ComboBox5.DisplayMember = "productname"
        ComboBox5.ValueMember = "productid"
        ComboBox5.DataSource = ProductBS

        TextBox1.DataBindings.Add("text", Me.bs, "refno")
        TextBox2.DataBindings.Add("text", Me.bs, "retailprice")
        TextBox3.DataBindings.Add("text", Me.bs, "staffprice")
        TextBox4.DataBindings.Add("text", Me.bs, "promotionprice")

        DateTimePicker1.DataBindings.Add("text", Me.bs, "promotionstartdate")
        DateTimePicker2.DataBindings.Add("text", Me.bs, "promotionenddate")

        ComboBox1.DataBindings.Add("selectedvalue", Me.bs, "producttypeid")
        ComboBox2.DataBindings.Add("selectedvalue", Me.bs, "familyid")
        ComboBox3.DataBindings.Add("selectedvalue", Me.bs, "brandid")
        ComboBox4.DataBindings.Add("selectedvalue", Me.bs, "descriptionid")
        ComboBox5.DataBindings.Add("selectedvalue", Me.bs, "productid")

        myrow = bs.Current
        Me.bs = bs
        Me.ds = DS



    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Validate()

        'Check Value combobox, if selecteditem isnull then ask for creating new value
        If Not validatecombobx() Then
            Me.DialogResult = Windows.Forms.DialogResult.None
            Exit Sub
        Else
            Me.DialogResult = Windows.Forms.DialogResult.OK
        End If
        myrow.Item("producttypename") = ComboBox1.Text
        myrow.Item("familyname") = ComboBox2.Text
        myrow.Item("brandname") = ComboBox3.Text
        myrow.Item("descriptionname") = ComboBox4.Text
        myrow.Item("productname") = ComboBox5.Text

        'Textbox4 (PromotionPrice) can be blank value , RetailPrice & StaffPrice Cannot Be 0 
        If TextBox4.Text <> "" Then
            If TextBox4.Text = 0 Then
                myrow.Item("promotionprice") = DBNull.Value
            End If
        End If
        
        If Not DateTimePicker1.Checked Then
            myrow.Item("promotionstartdate") = DBNull.Value
        End If
        If Not DateTimePicker2.Checked Then
            myrow.Item("promotionenddate") = DBNull.Value
        End If
        bs.EndEdit()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        bs.CancelEdit()
    End Sub

    Private Sub FormInputCutoff_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        DateTimePicker1.DataBindings.Clear()
        DateTimePicker2.DataBindings.Clear()

    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged

    End Sub

    Private Function validatecombobx() As Boolean
        ErrorProvider1.SetError(ComboBox1, "")
        ErrorProvider1.SetError(ComboBox2, "")
        ErrorProvider1.SetError(ComboBox3, "")
        ErrorProvider1.SetError(ComboBox4, "")
        ErrorProvider1.SetError(ComboBox5, "")


        If IsNothing(ComboBox1.SelectedValue) Then
            Dim producttypename As String = ComboBox1.Text
            If MessageBox.Show(ComboBox1.Text & " is not in the list. Do you want to create one?", "Question", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                Dim drv As DataRowView = ProductTypeBS.AddNew
                Dim dr As DataRow = drv.Row
                dr.Item("producttypename") = producttypename
                ds.Tables(1).Rows.Add(dr)
                ComboBox1.SelectedValue = dr.Item("producttypeid")
            Else
                ErrorProvider1.SetError(ComboBox1, "Value not in the list")
                Return False
            End If
        End If

        If IsNothing(ComboBox2.SelectedValue) Then
            Dim familyname As String = ComboBox2.Text
            If MessageBox.Show(ComboBox2.Text & " is not in the Family list. Do you want to create one?", "Question", MessageBoxButtons.YesNo) = DialogResult.Yes Then

                Dim drv As DataRowView = Familybs.AddNew
                Dim dr As DataRow = drv.Row


                dr.Item("familyname") = familyname
                ds.Tables(2).Rows.Add(dr)
                ComboBox2.SelectedValue = dr.Item("familyid")

                Dim myform As New FormInputFamily(Familybs)
                If Not myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    ds.Tables(2).Rows.Remove(dr)
                    Dim mydr As DataRow = myrow.Row

                    ComboBox2.SelectedValue = mydr.Item("familyid") 'mydr("familyid", DataRowVersion.Original).ToString
                    

                End If
            Else
                ErrorProvider1.SetError(ComboBox2, "Value not in the list")
                Return False
            End If
        End If

        If IsNothing(ComboBox5.SelectedValue) Then
            Dim productname As String = ComboBox5.Text
            If MessageBox.Show(ComboBox5.Text & " is not in the Product Name list. Do you want to create one?", "Question", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                Dim drv As DataRowView = ProductBS.AddNew
                Dim dr As DataRow = drv.Row
                dr.Item("productname") = productname
                ds.Tables(6).Rows.Add(dr)
                ComboBox5.SelectedValue = dr.Item("productid")
            Else
                ErrorProvider1.SetError(ComboBox5, "Value not in the list")
                Return False
            End If
        End If

        If IsNothing(ComboBox3.SelectedValue) Then
            Dim brandname As String = ComboBox3.Text
            If MessageBox.Show(ComboBox3.Text & " is not in the Brand list. Do you want to create one?", "Question", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                Dim drv As DataRowView = Brandbs.AddNew
                Dim dr As DataRow = drv.Row
                dr.Item("brandname") = brandname
                ds.Tables(3).Rows.Add(dr)
                ComboBox3.SelectedValue = dr.Item("brandid")

                Dim myform As New FormInputBrand(Brandbs)
                If Not myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    ds.Tables(3).Rows.Remove(dr)
                    Dim mydr As DataRow = myrow.Row
                    ComboBox3.SelectedValue = myrow.Item("brandid") 'mydr("brandid", DataRowVersion.Original).ToString
                End If

            Else
                ErrorProvider1.SetError(ComboBox3, "Value not in the list")
                Return False
            End If
        End If

        If IsNothing(ComboBox4.SelectedValue) Then
            Dim descriptionname As String = ComboBox4.Text
            If MessageBox.Show(ComboBox4.Text & " is not in the Description list. Do you want to create one?", "Question", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                Dim drv As DataRowView = Descriptionbs.AddNew
                Dim dr As DataRow = drv.Row
                dr.Item("descriptionname") = descriptionname
                ds.Tables(4).Rows.Add(dr)
                ComboBox4.SelectedValue = dr.Item("descriptionid")
            Else
                ErrorProvider1.SetError(ComboBox4, "Value not in the list")
                Return False
            End If
        End If



        Return True
    End Function

End Class