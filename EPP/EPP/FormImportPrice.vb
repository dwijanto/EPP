Imports System.Threading
Imports System.ComponentModel
Imports EPP.PublicClass
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports EPP.SharedClass

Public Class FormImportPrice
    Dim myCount As Integer = 0
    Dim listcount As Integer = 0


    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Delegate Sub ProcessReport(ByVal osheet As Excel.Worksheet)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)

    Dim FileName As String = String.Empty
    Dim Status As Boolean = False
    Dim ReadFileStatus As Boolean = False
    Dim Dataset1 As DataSet
    Dim sb As StringBuilder

    Dim Source As String
    Dim FolderBrowserDialog1 As New FolderBrowserDialog
    Dim mySelectedPath As String
    Dim startdate As Date
    Dim deletedata As Boolean


    Dim brandseq As Long = 0
    Dim productseq As Long = 0
    Dim descriptionseq As Long = 0
    Dim productypeseq As Long = 0
    Dim itemseq As Long = 0

    Dim brandsb As New StringBuilder
    Dim familysb As New StringBuilder
    Dim productsb As New StringBuilder
    Dim descriptionsb As New StringBuilder
    Dim producttypesb As New StringBuilder
    Dim itemsb As New StringBuilder
    Dim itempricesb As New StringBuilder



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ToolStripStatusLabel1.Text = ""
        ToolStripStatusLabel2.Text = ""
        If Not myThread.IsAlive Then
            'With FolderBrowserDialog1
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"

            If (OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                mySelectedPath = OpenFileDialog1.FileName

                Try
                    myThread = New System.Threading.Thread(myThreadDelegate)
                    myThread.SetApartmentState(ApartmentState.MTA)
                    myThread.Start()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
            'End With
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 2
                    ToolStripStatusLabel1.Text = message
                Case 3
                    ToolStripStatusLabel2.Text = message
                Case 4
                    ToolStripStatusLabel1.Text = message
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ToolStripProgressBar1.Minimum = 1
                    ToolStripProgressBar1.Value = myvalue(0)
                    ToolStripProgressBar1.Maximum = myvalue(1)
            End Select
        End If
    End Sub

    Sub DoWork()
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As String = String.Empty
        Dim sw As New Stopwatch
        sw.Start()

        ProgressReport(2, "Read Folder..")
        ReadFileStatus = ImportExcelFile(mySelectedPath, errSB)
        If ReadFileStatus Then
            sw.Stop()
            ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))           
        Else
            If errSB.Length > 100 Then
                'Using mystream As New StreamWriter(Application.StartupPath & "\error.txt")
                Using mystream As New StreamWriter(Path.GetDirectoryName(mySelectedPath) & "\error.txt")
                    mystream.WriteLine(errSB.ToString)
                End Using
                Process.Start(Path.GetDirectoryName(mySelectedPath) & "\error.txt")
                ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done with Error.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            Else
                ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done with Error.{3}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString, errSB.ToString))
            End If
        End If
        ProgressReport(5, "Set to continuous mode again")
        sw.Stop()
    End Sub

    Private Function validWorksheet(ByVal oSheet As Excel.Worksheet) As Boolean
        Dim myreturn As Boolean = False
        Dim mylist = {"Kitchen Electric", "Non-Food Electric ", "Cookware"}
        For Each lst In mylist
            If lst.ToString = oSheet.Name Then
                myreturn = True
                Exit For
            End If
        Next
        Return myreturn
    End Function

    Private Function ImportExcelFile(ByVal FileName As String, ByRef errSB As String) As Boolean
        Dim result As Boolean = False
        Dim StopWatch As New Stopwatch
        StopWatch.Start()
        'Open Excel
        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr

        Dim brandid As Long = 0
        Dim productid As Long = 0
        Dim descriptionid As Long = 0
        Dim itemid As Long = 0
        Dim producttypeid As Long = 0
        Dim familyid As Long
        
        Dim myrow As Integer = 0

        brandsb = New StringBuilder
        familysb = New StringBuilder
        productsb = New StringBuilder
        descriptionsb = New StringBuilder
        producttypesb = New StringBuilder
        itemsb = New StringBuilder
        itempricesb = New StringBuilder

        Try
            'Create Object Excel 
            ProgressReport(2, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            oXl.Visible = False
            oXl.DisplayAlerts = False
            ProgressReport(2, "Opening Template...")
            oWb = oXl.Workbooks.Open(FileName)
            oXl.Visible = False

            sb = New StringBuilder

            'get Initial Data
            Dim DS As New DataSet
            FillDataset(DS, errSB)

            For Each oSheet In oWb.Worksheets
                Dim orange = oSheet.Range("A1")
                Dim irows = GetLastRow(oXl, oSheet, orange)
                ProgressReport(2, oSheet.Name)
                'productype

                Dim pkey(0) As Object
                pkey(0) = oSheet.Name
                Dim myresult1 = DS.Tables(9).Rows.Find(pkey)
                'If validWorksheet(oSheet) Then
                If Not IsNothing(myresult1) Then
                    producttypeid = myresult1.Item(1)
                    'Dim orange = oSheet.Range("A1")
                    'Dim irows = GetLastRow(oXl, oSheet, orange)
                    'ProgressReport(2, oSheet.Name)
                    ''productype

                    'Dim pkey(0) As Object
                    'pkey(0) = oSheet.Name
                    'Dim myresult1 = DS.Tables(9).Rows.Find(pkey)
                    'producttypeid = myresult1.Item(1)



                    Try


                        For myrow = 8 To irows
                            'If myrow = 228 Then
                            '    Debug.Print("debug mode")
                            'End If
                            ProgressReport(7, myrow & "," & irows)
                            Dim myresult As DataRow
                            Dim myrecord(11) As String
                            For mycol = 1 To 12
                                myrecord(mycol - 1) = ""
                                If Not IsNothing(oSheet.Cells(myrow, mycol).value) Then
                                    Dim myvalue = oSheet.Cells(myrow, mycol).value.ToString
                                    myrecord(mycol - 1) = myvalue
                                End If
                            Next

                            'Brand
                            If myrecord(2).ToString <> "" Then
                                Dim pkey0(0) As Object
                                pkey0(0) = myrecord(2).ToString.ToUpper
                                myresult = DS.Tables(0).Rows.Find(pkey0)
                                If IsNothing(myresult) Then
                                    'show errormessage
                                    errSB = "Brand Error :: " & validstr(myrecord(2)) & " is not found."
                                    'Dim dr As DataRow = DS.Tables(0).NewRow
                                    'brandseq += 1
                                    'dr.Item(1) = brandseq
                                    'dr.Item(0) = myrecord(1)
                                    'DS.Tables(0).Rows.Add(dr)
                                    'brandid = brandseq
                                    'brandsb.Append(validstr(myrecord(2)) & vbCrLf)
                                Else
                                    brandid = myresult.Item(1)
                                End If
                            End If

                            'Description
                            If myrecord(6).ToString <> "" Then
                                Dim pkey1(0) As Object
                                'pkey1(0) = validstr(myrecord(6))
                                'pkey3(0) = String.Format("A{0}A", validstr(myrecord(1)))
                                'pkey1(0) = myrecord(6)
                                pkey1(0) = String.Format("A{0}A", myrecord(6).ToUpper)
                                myresult = DS.Tables(1).Rows.Find(pkey1)
                                If IsNothing(myresult) Then
                                    Try
                                        Dim dr As DataRow = DS.Tables(1).NewRow
                                        descriptionseq += 1
                                        dr.Item(1) = descriptionseq
                                        dr.Item(0) = myrecord(6).ToUpper
                                        DS.Tables(1).Rows.Add(dr)
                                        descriptionid = descriptionseq
                                        'descriptionsb.Append(validstr(myrecord(6)) & vbCrLf)
                                        descriptionsb.Append(myrecord(6).ToUpper & vbCrLf)
                                    Catch ex As Exception
                                        MessageBox.Show(ex.Message)
                                    End Try
                                    
                                Else
                                    descriptionid = myresult.Item(1)
                                End If
                            End If

                            'family
                            If myrecord(3).ToString <> "" Then
                                Dim pkey2(0) As Object
                                pkey2(0) = myrecord(3)
                                myresult = DS.Tables(2).Rows.Find(pkey2)
                                If IsNothing(myresult) Then
                                    errSB = "Family Error :: " & myrecord(3) & " is not found."
                                    'Dim dr As DataRow = DS.Tables(2).NewRow
                                    'dr.Item(0) = myrecord(3)
                                    'dr.Item(1) = myrecord(2)
                                    'DS.Tables(2).Rows.Add(dr)
                                    'familyid = myrecord(2)
                                    'familysb.Append(validlong(myrecord(2)) & vbTab &
                                    '                validstr(myrecord(3)) & vbCrLf)
                                Else
                                    familyid = myresult.Item(1)
                                End If
                            End If

                            'product
                            If myrecord(5).ToString <> "" Then
                                Dim pkey4(0) As Object
                                'pkey4(0) = validstr(myrecord(5))
                                pkey4(0) = String.Format("A{0}A", myrecord(5).ToUpper)
                                myresult = DS.Tables(4).Rows.Find(pkey4)
                                If IsNothing(myresult) Then
                                    Dim dr As DataRow = DS.Tables(4).NewRow
                                    productseq += 1
                                    dr.Item(1) = productseq
                                    dr.Item(0) = String.Format("A{0}A", myrecord(5).ToUpper)
                                    DS.Tables(4).Rows.Add(dr)
                                    productid = productseq
                                    productsb.Append(validstr(myrecord(5).ToUpper) & vbCrLf)
                                Else
                                    productid = myresult.Item(1)
                                End If
                            End If

                            'item
                            If myrecord(1).ToString <> "" Then
                                Dim pkey3(0) As Object
                                pkey3(0) = String.Format("A{0}A", validstr(myrecord(1).ToUpper))
                                myresult = DS.Tables(3).Rows.Find(pkey3)
                                If IsNothing(myresult) Then
                                    Dim dr As DataRow = DS.Tables(3).NewRow
                                    itemseq += 1
                                    dr.Item(1) = itemseq
                                    dr.Item(0) = String.Format("A{0}A", myrecord(1).ToUpper)
                                    DS.Tables(3).Rows.Add(dr)
                                    itemid = itemseq
                                    itemsb.Append(validstr(myrecord(1).ToUpper) & vbTab &
                                                  validlong(brandid.ToString) & vbTab &
                                                  validlong(familyid.ToString) & vbTab &
                                                  validlong(productid.ToString) & vbTab &
                                                  validlong(producttypeid.ToString) & vbTab &
                                                  validlong(descriptionid.ToString) & vbCrLf)
                                Else
                                    itemid = myresult.Item(1)
                                End If
                            End If
                            If myrecord(7).ToString <> "" Then
                                itempricesb.Append(validint(itemid.ToString) & vbTab &
                                      validreal(myrecord(7).ToString) & vbTab &
                                      validreal(myrecord(8).ToString) & vbTab &
                                      validreal(myrecord(9).ToString) & vbTab &
                                      DateFormatyyyyMMddString(myrecord(10).ToString) & vbTab &
                                      DateFormatyyyyMMddString(myrecord(11).ToString) & vbCrLf)
                            Else
                                Debug.Print("")
                            End If
                            

                        Next
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try

                End If
            Next
            If errSB.Length > 0 Then
                Return False
            End If

            Dim message As String = String.Empty

            'ProgressReport(3, "Copy to DB..")

            Dim sqlstr As String = String.Empty
            Dim errmsg As String = String.Empty
            If CheckBox1.Checked Then
                Sqlstr = "delete from shop.itemprice;"
                If Not DbAdapter1.ExecuteNonQuery(Sqlstr, message:=errmsg) Then
                    errSB = errmsg
                    Return False
                End If
            End If
            

            If brandsb.Length > 0 Then
                ProgressReport(2, "Copy To Db (Brand)")
                Sqlstr = "copy shop.brand(brandname) from stdin with null as 'Null';"
                message = DbAdapter1.copy(Sqlstr, brandsb.ToString, result)
                If message <> "" Then
                    errSB = message
                    Return False
                End If
            End If

            If familysb.Length > 0 Then
                ProgressReport(2, "Copy To Db (Family)")
                Sqlstr = "copy shop.family(familyid,familyname) from stdin with null as 'Null';"
                message = DbAdapter1.copy(Sqlstr, familysb.ToString, result)
                If message <> "" Then
                    errSB = message
                    Return False
                End If
            End If

            Dim lastrow As Integer

            If DS.Tables(8).Rows.Count <> 0 Then
                lastrow = DS.Tables(8).Rows(0).Item(0) + 1
            Else
                lastrow = 1
            End If

            If productsb.Length > 0 Then
                ProgressReport(2, "Copy To Db (Product)")
                sqlstr = "select setval('shop.product_productid_seq'," & lastrow & ",false);copy shop.product(productname) from stdin with null as 'Null';"
                message = DbAdapter1.copy(Sqlstr, productsb.ToString, result)
                If message <> "" Then
                    errSB = message
                    Return False
                End If
            End If


            If DS.Tables(6).Rows.Count <> 0 Then
                lastrow = DS.Tables(6).Rows(0).Item(0) + 1
            Else
                lastrow = 1
            End If

            If descriptionsb.Length > 0 Then
                ProgressReport(2, "Copy To Db (Description)")
                sqlstr = "select setval('shop.description_descriptionid_seq'," & lastrow & ",false);copy shop.description(descriptionname) from stdin with null as 'Null';"
                message = DbAdapter1.copy(Sqlstr, descriptionsb.ToString, result)
                If message <> "" Then
                    errSB = message
                    Return False
                End If
            End If

            If DS.Tables(7).Rows.Count <> 0 Then
                lastrow = DS.Tables(7).Rows(0).Item(0) + 1
            Else
                lastrow = 1
            End If

            If itemsb.Length > 0 Then
                ProgressReport(2, "Copy To Db (Item)")
                sqlstr = "select setval('shop.item_itemid_seq'," & lastrow & ",false);copy shop.item(refno,brandid,familyid,productid,producttypeid,descriptionid) from stdin with null as 'Null';"
                message = DbAdapter1.copy(Sqlstr, itemsb.ToString, result)
                If message <> "" Then
                    errSB = message
                    Return False
                End If
            End If

            If itempricesb.Length > 0 Then
                ProgressReport(2, "Copy To Db (ItemPrice)")
                Sqlstr = "copy shop.itemprice(itempriceid,retailprice,staffprice,promotionprice,promotionstartdate,promotionenddate) from stdin with null as 'Null';"
                message = DbAdapter1.copy(Sqlstr, itempricesb.ToString, result)
                If message <> "" Then
                    errSB = message
                    Return False
                End If
            End If
            StopWatch.Stop()
            ProgressReport(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)


            result = True
        Catch ex As Exception
            ProgressReport(3, ex.Message & " - Row Number::" & myrow & " - " & FileName)
            errSB = ex.Message & " - Row Number::" & myrow & " - " & FileName
        Finally
            'oXl.ScreenUpdating = True
            'clear excel from memory
            Try
                oXl.Quit()
                releaseComObject(oSheet)
                releaseComObject(oWb)
                releaseComObject(oXl)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception

            End Try

            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try

        End Try
        Return result
    End Function

   




    Private Function FillDataset(ByRef DS As DataSet, ByRef errmessage As String) As Boolean
        Dim myret As Boolean = False
        Dim Sqlstr As String = " select brandname,brandid  from shop.brand;" &
                               " select distinct 'A' || upper(descriptionname) || 'A' as descriptionname,descriptionid from shop.description;" &
                               " select familyname,familyid from shop.family;" &
                               " select 'A' || upper(refno) || 'A' as refno,itemid from shop.item;" &
                               " select 'A' || upper(productname) || 'A' as productname,max(productid) as productid from shop.product group by upper(productname) order by productid;" &
                               " select brandid from shop.brand order by brandid desc limit 1;" &
                               " select descriptionid from shop.description order by descriptionid desc limit 1;" &
                               " select itemid from shop.item order by itemid desc limit 1;" &
                               " select productid from shop.product order by productid desc limit 1;" &
                               " select producttypename,producttypeid from shop.producttype"

        DS.CaseSensitive = True
        If DbAdapter1.GetDataSet(Sqlstr, DS, errmessage) Then
            DS.Tables(0).TableName = "brand"
            DS.Tables(1).TableName = "description"
            DS.Tables(2).TableName = "family"
            DS.Tables(3).TableName = "item"
            DS.Tables(4).TableName = "product"
            Using mystream As New StreamWriter("debug.txt")
                For Each dr In DS.Tables(1).Rows
                    mystream.WriteLine(String.Format("{0}{1}{2}", dr.Item(0).ToString, vbTab, dr.item(1).ToString))
                Next

            End Using
           

            Dim idx0(0) As DataColumn
            idx0(0) = DS.Tables(0).Columns(0)
            DS.Tables(0).PrimaryKey = idx0

            Dim idx1(0) As DataColumn
            idx1(0) = DS.Tables(1).Columns(0)
            DS.Tables(1).PrimaryKey = idx1

            Dim idx2(0) As DataColumn
            idx2(0) = DS.Tables(2).Columns(1)
            DS.Tables(2).PrimaryKey = idx2

            'For i = 0 To DS.Tables(3).Rows.Count - 1
            '    Debug.Print(String.Format("{0}", DS.Tables(3).Rows(i).Item(0).ToString))
            'Next

            Dim idx3(0) As DataColumn
            idx3(0) = DS.Tables("item").Columns(0)
            'idx3(1) = DS.Tables(3).Columns(1)
            DS.Tables("item").PrimaryKey = idx3

            Dim idx4(0) As DataColumn
            idx4(0) = DS.Tables(4).Columns(0)
            DS.Tables(4).PrimaryKey = idx4

            Dim idx9(0) As DataColumn
            idx9(0) = DS.Tables(9).Columns(0)
            DS.Tables(9).PrimaryKey = idx9

            brandseq = 0
            If DS.Tables(5).Rows.Count > 0 Then
                brandseq = DS.Tables(5).Rows(0).Item(0)
            End If

            descriptionseq = 0
            If DS.Tables(6).Rows.Count > 0 Then
                descriptionseq = DS.Tables(6).Rows(0).Item(0)
            End If

            itemseq = 0
            If DS.Tables(7).Rows.Count > 0 Then
                itemseq = DS.Tables(7).Rows(0).Item(0)
            End If

            productseq = 0
            If DS.Tables(8).Rows.Count > 0 Then
                productseq = DS.Tables(8).Rows(0).Item(0)
            End If
        Else
            Return False
        End If
        myret = True
        Return myret
    End Function


    Private Sub FormImportPrice_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If myThread.IsAlive Then
            e.Cancel = True
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub
End Class

