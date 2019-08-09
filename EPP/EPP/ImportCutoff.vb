Imports System.Threading
Imports System.Text
Imports EPP.PublicClass

Public Class ImportCutoff
    Property mythread As System.Threading.Thread
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim OpenFiledialog1 As New OpenFileDialog
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    'Property myThread As New System.Threading.Thread(myThreadDelegate)
    Dim myFileName
    Property myobj As Object

    Public Sub New(ByVal myObj As Object)
        Me.myobj = myObj
        Me.mythread = New System.Threading.Thread(myThreadDelegate)
    End Sub
    Public Sub Run()
        If Not mythread.IsAlive Then
            'With FolderBrowserDialog1
            OpenFiledialog1.FileName = ""
            OpenFiledialog1.Filter = "TXT files (*.txt|*.txt|All files (*.*)|*.*"

            If (OpenFiledialog1.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                myFileName = OpenFiledialog1.FileName

                Try
                    mythread = New System.Threading.Thread(myThreadDelegate)
                    mythread.SetApartmentState(ApartmentState.MTA)
                    mythread.Start()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
            'End With
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Sub DoWork()
        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder
        Dim myrecord() As String
        Using objTFParser = New FileIO.TextFieldParser(OpenFiledialog1.FileName)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                ProgressReport(1, "Read Data")
                Dim sqlstr = "select cutoff::date from shop.cutoff;"
                Dim myDS As New DataSet
                Dim errmessage As String = String.Empty
                If Not DbAdapter1.GetDataSet(sqlstr, myDS, errmessage) Then
                    Exit Sub
                Else
                    myDS.Tables(0).TableName = "Cutoff"

                    Dim idx0(0) As DataColumn
                    idx0(0) = myDS.Tables(0).Columns(0)
                    myDS.Tables(0).PrimaryKey = idx0
                End If

                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        If IsDate(myrecord(0)) Then
                            Dim result As Object
                            Dim pkey(0) As Object
                            pkey(0) = CDate(myrecord(0)).Date
                            result = myDS.Tables(0).Rows.Find(pkey)
                            If IsNothing(result) Then
                                Dim dr As DataRow = myDS.Tables(0).NewRow
                                dr.Item(0) = CDate(myrecord(0))
                                myDS.Tables(0).Rows.Add(dr)

                                myInsert.Append(myrecord(0) & vbTab &
                                                myrecord(1) & vbTab &
                                                myrecord(2) & vbTab &
                                                myrecord(3) & vbCrLf)
                            End If
                        End If
                    End If
                    count += 1
                Loop
            End With
        End Using
        'update record
        If myInsert.Length > 0 Then
            ProgressReport(1, "Start Add New Records")
            Dim sqlstr As String = "copy shop.cutoff(cutoff,emailnotificationdate,chequesubmitdate,collectiondate) from stdin with null as 'Null';"

            Dim ra As Long = 0
            Dim errmessage As String = String.Empty
            Dim myret As Boolean = False
            Try
                errmessage = DbAdapter1.copy(sqlstr, myInsert.ToString, myret)
                If myret Then
                    ProgressReport(1, "Add Records Done.")
                Else
                    ProgressReport(1, errmessage)
                End If
            Catch ex As Exception
                ProgressReport(1, ex.Message)
            End Try
        End If
        ProgressReport(5, "Set Continuous Again")
        ProgressReport(8, "RefreshData")
    End Sub
    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If myobj.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            myobj.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 2
                    myobj.ToolStripStatusLabel1.Text = message
                Case 3
                    myobj.ToolStripStatusLabel2.Text = message
                Case 4
                    myobj.ToolStripStatusLabel1.Text = message
                Case 5
                    myobj.ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    myobj.ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    myobj.ToolStripProgressBar1.Minimum = 1
                    myobj.ToolStripProgressBar1.Value = myvalue(0)
                    myobj.ToolStripProgressBar1.Maximum = myvalue(1)
                Case 8
                    myobj.loaddata()
            End Select
        End If
    End Sub
End Class

