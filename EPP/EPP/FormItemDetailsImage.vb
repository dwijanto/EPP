Imports System.IO

Public Class FormItemDetailsImage
    Dim bs As BindingSource
    Dim dr As DataRow

    Dim myimage As Image
    'Dim mypath As String = "C:\images"
    'Dim zoommode As PictureBoxSizeMode
    Dim m_panStartPoint As Point
    Dim zoomMode As Boolean
    Dim deltax As Integer
    Dim deltay As Integer
    Dim _imageExt As String
    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Dim b As ToolStripButton = DirectCast(sender, ToolStripButton)
        If b.Text = "Show Product Information" Then
            b.Text = "Show Picture Only"
            SplitContainer1.Panel1Collapsed = False
        Else
            SplitContainer1.Panel1Collapsed = True
            b.Text = "Show Product Information"
        End If
    End Sub

    Public Sub NewOld(ByRef bs As BindingSource, ByVal imagepath As String, ByVal imageext As String())

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

        ' Add any initialization after the InitializeComponent() call.
        SplitContainer1.FixedPanel = FixedPanel.Panel1
        SplitContainer1.SplitterWidth = 2
        Dim ext As String() = imageext '{"jpg", "jpeg", "png", "bmp", "tif", "tiff","gif"}
        Dim myfilename As String = String.Empty
        Try

            For Each myext As String In ext
                'myfilename = mypath & "\" & dr.Item("refno") & "." & myext
                myfilename = imagepath & "\" & dr.Item("refno") & "." & myext
                If File.Exists(myfilename) Then
                    Exit For
                Else
                    myfilename = ""
                End If

            Next
            If myfilename <> "" Then
                myimage = Image.FromFile(myfilename)
                PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage 'PictureBoxSizeMode.Normal
                zoomMode = myimage.Width > PictureBox1.Width Or myimage.Height > PictureBox1.Height
                If zoomMode Then
                    PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                End If
                PictureBox1.Image = myimage
            Else
                SplitContainer1.Panel2Collapsed = True
                ToolStrip1.Visible = False
                Me.Size = New Point(450, 550)
            End If

        Catch ex As Exception

        End Try
    End Sub
    Public Sub New(ByRef bs As BindingSource, ByVal imagepath As String, ByVal imageext As String())

        ' This call is required by the designer.
        InitializeComponent()
        Me.bs = bs
        Me._imageExt = String.Join(",", imageext)
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

        ' Add any initialization after the InitializeComponent() call.
        SplitContainer1.FixedPanel = FixedPanel.Panel1
        SplitContainer1.SplitterWidth = 2
        Dim ext As String() = imageext '{"jpg", "jpeg", "png", "bmp", "tif", "tiff","gif"}
        Dim myfilename As String = String.Empty
        Try
            'imagepath = "I:\GSHK Sales\Product Photos"
            subFolder(imagepath, dr.Item("refno"))

            If IsNothing(myimage) Then
                SplitContainer1.Panel2Collapsed = True
                ToolStrip1.Visible = False
                Me.Size = New Point(450, 550)
            End If
            'For Each myext As String In ext
            '    'myfilename = mypath & "\" & dr.Item("refno") & "." & myext
            '    myfilename = imagepath & "\" & dr.Item("refno") & "." & myext
            '    If File.Exists(myfilename) Then
            '        Exit For
            '    Else
            '        myfilename = ""
            '    End If

            'Next
            'If myfilename <> "" Then
            '    myimage = Image.FromFile(myfilename)
            '    PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage 'PictureBoxSizeMode.Normal
            '    zoomMode = myimage.Width > PictureBox1.Width Or myimage.Height > PictureBox1.Height
            '    If zoomMode Then
            '        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
            '    End If
            '    PictureBox1.Image = myimage
            'Else
            '    SplitContainer1.Panel2Collapsed = True
            '    ToolStrip1.Visible = False
            '    Me.Size = New Point(450, 550)
            'End If

        Catch ex As Exception

        End Try
    End Sub
    Public Sub subFolder(ByVal path As String, ByVal refno As String)
        Dim myfilename As String() = {}
        Dim sourceDir As DirectoryInfo = New DirectoryInfo(path)
        Dim pathIndex As Integer
        pathIndex = path.LastIndexOf("\")
        If sourceDir.Exists Then
            Dim subDir As DirectoryInfo
            For Each subDir In sourceDir.GetDirectories()
                subFolder(subDir.FullName, refno)
            Next
            Dim shortfilename = refno
            If shortfilename.Length > 6 Then
                shortfilename = refno.Substring(0, 6)
            End If
            myfilename = Directory.GetFiles(path, String.Format("*{0}*", shortfilename))
            If myfilename.Length > 0 Then
                For i = 0 To myfilename.Length - 1
                    If _imageExt.IndexOf(IO.Path.GetExtension(myfilename(i)).ToString.ToLower.Substring(1)) >= 0 Then
                        If IsNothing(myimage) Then
                            myimage = Image.FromFile(myfilename(0))
                            PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage 'PictureBoxSizeMode.Normal
                            zoomMode = myimage.Width > PictureBox1.Width Or myimage.Height > PictureBox1.Height
                            If zoomMode Then
                                PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                            End If
                            PictureBox1.Image = myimage
                            Exit For
                        End If
                    End If
                    

                Next

            Else
                'SplitContainer1.Panel2Collapsed = True
                'ToolStrip1.Visible = False
            End If
        End If

    End Sub
    Private Sub DisplayImage_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If Not IsNothing(myimage) Then
            myimage.Dispose()
        End If


    End Sub

    'Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
    '    Dim tb As ToolStripButton = DirectCast(sender, ToolStripButton)
    '    If tb.Text = "Show Original Size" Then
    '        tb.Text = "Fit Image"
    '        PictureBox1.SizeMode = PictureBoxSizeMode.Normal
    '        PictureBox1.Dock = DockStyle.None
    '        PictureBox1.Width = myimage.Width
    '        PictureBox1.Height = myimage.Height
    '    Else

    '        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
    '        PictureBox1.Dock = DockStyle.Fill
    '        tb.Text = "Show Original Size"
    '    End If


    'End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub PictureBox1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseDown
        Me.Cursor = Cursors.Default

        If PictureBox1.Dock = DockStyle.None Then
            Me.Cursor = Cursors.Hand
        End If
        m_panStartPoint = New Point(e.X, e.Y)
        'Debug.Print("MouseDown " & m_panStartPoint.X & " " & m_panStartPoint.Y)
        'Debug.Print("After MyPanel1 Autoscroll " & MyPanel1.AutoScrollPosition.X & "  " & MyPanel1.AutoScrollPosition.Y)
    End Sub

    Private Sub PictureBox1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.MouseEnter
        'PictureBox1.Focus()
        If PictureBox1.Dock = DockStyle.None Then
            Me.Cursor = Cursors.SizeAll
        End If
    End Sub

    Private Sub PictureBox1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.MouseLeave
        Me.Cursor = Cursors.Default
    End Sub


    Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
        'Debug.Print("mouse_move " & e.X & " " & e.Y)
        If e.Button = Windows.Forms.MouseButtons.Left Then
            deltax = m_panStartPoint.X - e.X
            deltay = m_panStartPoint.Y - e.Y

            'SplitContainer1.Panel2.AutoScrollPosition = New Drawing.Point((deltax - SplitContainer1.Panel2.AutoScrollPosition.X), (deltay - SplitContainer1.Panel2.AutoScrollPosition.Y))
            'Debug.Print("Before MyPanel1 Autoscroll " & MyPanel1.AutoScrollPosition.X & "  " & MyPanel1.AutoScrollPosition.Y)
            'Debug.Print("Before PictureBox1 Position " & PictureBox1.Location.X & "  " & PictureBox1.Location.Y)

            MyPanel1.AutoScrollPosition = New Drawing.Point((deltax - MyPanel1.AutoScrollPosition.X), (deltay - MyPanel1.AutoScrollPosition.Y))

            'Debug.Print("After MyPanel1 Autoscroll " & MyPanel1.AutoScrollPosition.X & "  " & MyPanel1.AutoScrollPosition.Y)
            'Debug.Print("After PictureBox1 Position " & PictureBox1.Location.X & "  " & PictureBox1.Location.Y)

        End If
    End Sub

    Private Sub PictureBox1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseUp
        If PictureBox1.Dock = DockStyle.None Then
            Me.Cursor = Cursors.SizeAll
        End If
        'If MyPanel1.AutoScrollPosition.X = 0 Then
        '    PictureBox1.Location = New Point(0, MyPanel1.AutoScrollPosition.Y)
        'End If
        'If MyPanel1.AutoScrollPosition.Y = 0 Then
        '    PictureBox1.Location = New Point(MyPanel1.AutoScrollPosition.X, 0)
        'End If
        PictureBox1.Location = MyPanel1.AutoScrollPosition
        'Debug.Print("MouseUp MyPanel1 Autoscroll " & MyPanel1.AutoScrollPosition.X & "  " & MyPanel1.AutoScrollPosition.Y)
        'Debug.Print("MouseUp PictureBox1 Position " & PictureBox1.Location.X & "  " & PictureBox1.Location.Y)
        'End If
    End Sub


    Private Sub SplitContainer1_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles SplitContainer1.MouseWheel
        If PictureBox1.Dock <> DockStyle.None Then
            'Assign Picturebox1.Dock after calculate myscale
            PictureBox1.SizeMode = PictureBoxSizeMode.Normal
            Dim myscale = Math.Max(myimage.Width / PictureBox1.Width, myimage.Height / PictureBox1.Height)
            PictureBox1.Dock = DockStyle.None
            PictureBox1.Width = (myimage.Width / myscale)
            PictureBox1.Height = (myimage.Height / myscale)
            '
            'MyPanel1.AutoScrollPosition = New Drawing.Point(0, 0)
            'PictureBox1.Location = New Point(0, 0)

        End If
        PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
        Dim _scale As New SizeF(1, 1)
        Dim _scaleDelta As New SizeF(0.05, 0.05)

        If e.Delta < 0 Then
            _scale += _scaleDelta
        Else
            _scale -= _scaleDelta
        End If
        PictureBox1.Scale(_scale)
        PictureBox1.Location = MyPanel1.AutoScrollPosition
    End Sub

    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click, ToolStripMenuItem2.Click, ToolStripMenuItem3.Click, ToolStripMenuItem4.Click, ToolStripMenuItem5.Click, ToolStripMenuItem6.Click, FitToolStripMenuItem.Click
        Dim ts As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)

        Select Case ts.Text
            Case "Fit"
                PictureBox1.Dock = DockStyle.Fill
                'PictureBox1.Padding = New Padding(10)
                PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage 'PictureBoxSizeMode.Normal
                If zoomMode Then
                    PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
                End If
            Case Else
                PictureBoxScale(CDbl(ts.Tag))
        End Select
    End Sub

    Private Sub ToolStripTextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ToolStripTextBox1.KeyPress
        If Not IsNumeric(e.KeyChar) Then
            If e.KeyChar <> Chr(8) Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub ToolStripTextBox1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ToolStripTextBox1.KeyUp
        If e.KeyCode = Keys.Enter Then
            Dim ts As ToolStripTextBox = DirectCast(sender, ToolStripTextBox)

            PictureBoxScale(CDbl(ts.Text) / 100)

            ToolStripDropDownButton1.DropDown.Close()
            ts.Text = ""
        End If
    End Sub

    Private Sub PictureBoxScale(ByVal scale As Double)
        PictureBox1.SizeMode = PictureBoxSizeMode.Normal
        PictureBox1.Dock = DockStyle.None
        PictureBox1.Width = myimage.Width
        PictureBox1.Height = myimage.Height
        PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
        PictureBox1.Scale(New SizeF(scale, scale))
        'SplitContainer1.Panel2.AutoScrollPosition = New Drawing.Point(0, 0)
        MyPanel1.AutoScrollPosition = New Drawing.Point(0, 0)
        PictureBox1.Location = MyPanel1.AutoScrollPosition 'New Point(0, 0)
        'PictureBox1.Padding = New Padding(0)
    End Sub
End Class