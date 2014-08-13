Imports System.Math

Public Class Image_Display
    Inherits System.Windows.Forms.Form


    Private interval As String
    Private fullscreen As Boolean
    Private filelist As ArrayList
    Private currentindex As Integer
    Private maxindex As Integer
    Dim cnt As Integer = 0

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal imagelist As ArrayList, ByVal timeinterval As String, ByVal showfullscreen As Boolean)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        filelist = imagelist
        maxindex = filelist.Count() - 1
        currentindex = 0
        interval = timeinterval
        fullscreen = showfullscreen


    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Image_Display))
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem8, Me.MenuItem7, Me.MenuItem1, Me.MenuItem2, Me.MenuItem3, Me.MenuItem4, Me.MenuItem6, Me.MenuItem5})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 2
        Me.MenuItem1.Text = "Next Image"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 3
        Me.MenuItem2.Text = "Previous Image"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 4
        Me.MenuItem3.Text = "Pause"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 5
        Me.MenuItem4.Text = "Play"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 7
        Me.MenuItem5.Text = "Stop"
        '
        'Timer1
        '
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(32, 32)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Black
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Location = New System.Drawing.Point(2, 2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(288, 262)
        Me.Panel1.TabIndex = 1
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 1
        Me.MenuItem7.Text = "-"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 0
        Me.MenuItem8.Text = ""
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 6
        Me.MenuItem6.Text = "-"
        '
        'Image_Display
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.DimGray
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.ContextMenu = Me.ContextMenu1
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Image_Display"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Image Display"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Const WM_NCHITTEST As Integer = &H84
    Private Const HTCLIENT As Integer = &H1
    Private Const HTCAPTION As Integer = &H2

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Select Case m.Msg
            Case WM_NCHITTEST
                MyBase.WndProc(m)
                If (m.Result.ToInt32 = HTCLIENT) Then
                    m.Result = IntPtr.op_Explicit(HTCAPTION)
                End If
                Exit Sub
        End Select

        MyBase.WndProc(m)
    End Sub

    Private Sub Error_Handler(ByVal message As String)
        Try
            Dim Display_Message1 As New Display_Message(message)
            Display_Message1.ShowDialog()
        Catch ex As Exception
            MsgBox("An error occurred in Image Slide Show's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Try
            Timer1.Stop()
            PictureBox1.Image.Dispose()
            Me.Close()
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while closing the Image Display Screen. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Timer1.Stop()
            display_image(CStr(filelist.Item(currentindex)))
            
            Timer1.Start()
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while checking the applications' status. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub display_image(ByVal imagetoload As String)
        Try
            PictureBox1.Visible = False
            If Not (PictureBox1.Image Is Nothing) Then
                PictureBox1.Image.Dispose()
                PictureBox1.Image = Nothing
            End If
            PictureBox1.Image = Image.FromFile(imagetoload)
            PictureBox1.Height = PictureBox1.Image.Height
            PictureBox1.Width = PictureBox1.Image.Width
            ContextMenu1.MenuItems(0).Text = imagetoload.Split("\")(imagetoload.Split("\").Length - 1)



            If fullscreen = True Then
                Zoom_Perfect()
            End If
            reposition_image()
            PictureBox1.Visible = True

            PictureBox1.Refresh()
            currentindex = currentindex + 1
            If currentindex > maxindex Then
                currentindex = currentindex - maxindex - 1
            End If
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while displaying the specified image file. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub reposition_image()
        PictureBox1.Top = Round(((Panel1.Height - PictureBox1.Height) / 2), 0)
        PictureBox1.Left = Round(((Panel1.Width - PictureBox1.Width) / 2), 0)
    End Sub

    Private Sub Zoom_Perfect()
        Try
            If PictureBox1.Image.Height <> Panel1.Height Then
                If PictureBox1.Image.Height >= PictureBox1.Image.Width Then
                    PictureBox1.Height = Panel1.Height
                    PictureBox1.Width = PictureBox1.Image.Width * (Panel1.Height / PictureBox1.Image.Height)
                Else
                    PictureBox1.Width = Panel1.Width
                    PictureBox1.Height = PictureBox1.Image.Height * (Panel1.Width / PictureBox1.Image.Width)
                End If

            End If
            PictureBox1.Top = 0
            PictureBox1.Left = 0
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Image_Display_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.Height = Screen.FromControl(Me).GetBounds(Me).Height
            Me.Width = Screen.FromControl(Me).GetBounds(Me).Width
            Panel1.Height = Screen.FromControl(Me).GetBounds(Me).Height - 8
            Panel1.Width = Screen.FromControl(Me).GetBounds(Me).Width - 8
            Me.Top = 0
            Me.Left = 0
            Panel1.Top = Round(((Me.Height - Panel1.Height) / 2), 0)
            Panel1.Left = Round(((Me.Width - Panel1.Width) / 2), 0)

            Me.Refresh()
            display_image(CStr(filelist.Item(currentindex)))
            currentindex = currentindex + 1
            If currentindex > maxindex Then
                currentindex = 0
            End If
            Timer1.Interval = CInt(interval) * 1000
            Timer1.Start()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Try
            Timer1.Stop()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        Try
            Timer1.Start()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        display_image(CStr(filelist.Item(currentindex)))

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        currentindex = currentindex - 2
        If currentindex < 0 Then
            currentindex = currentindex + maxindex + 1
        End If
        display_image(CStr(filelist.Item(currentindex)))

    End Sub





    Protected Overrides Sub OnMouseWheel(ByVal e As System.Windows.Forms.MouseEventArgs)

        'this event is fired everytime the user moves the mouse scroll wheel
        'one click. So each notch the wheel moves, this event is fired.

        'e.Delta will return either negative 120 or positive 120(returns 
        'the number 120 on my computer at least)

        'check if delta returns a negative number, then decrease the number
        If e.Delta < 0 Then

            cnt -= 10
            'MsgBox(Screen.FromControl(Me).GetBounds(Me).Height + cnt & " grth " & PictureBox1.Height & " And " & Me.Width = Screen.FromControl(Me).GetBounds(Me).Width + cnt & " grth " & PictureBox1.Width)
            'If Screen.FromControl(Me).GetBounds(Me).Height + cnt > PictureBox1.Height And Me.Width = Screen.FromControl(Me).GetBounds(Me).Width + cnt > PictureBox1.Width Then
            If Screen.FromControl(Me).GetBounds(Me).Height + cnt > 20 And Screen.FromControl(Me).GetBounds(Me).Width + cnt > 20 Then

                Me.Height = Screen.FromControl(Me).GetBounds(Me).Height + cnt
                Me.Width = Screen.FromControl(Me).GetBounds(Me).Width + cnt
                Panel1.Height = Screen.FromControl(Me).GetBounds(Me).Height - 8 + cnt
                Panel1.Width = Screen.FromControl(Me).GetBounds(Me).Width - 8 + cnt

                Panel1.Top = Round(((Me.Height - Panel1.Height) / 2), 0)
                Panel1.Left = Round(((Me.Width - Panel1.Width) / 2), 0)

                Me.Refresh()
            Else
                cnt += 10
            End If
            'delta returns a positive number, so increase the number
        Else

            cnt += 10


            Me.Height = Screen.FromControl(Me).GetBounds(Me).Height + cnt
            Me.Width = Screen.FromControl(Me).GetBounds(Me).Width + cnt
            Panel1.Height = Screen.FromControl(Me).GetBounds(Me).Height - 8 + cnt
            Panel1.Width = Screen.FromControl(Me).GetBounds(Me).Width - 8 + cnt

            Panel1.Top = Round(((Me.Height - Panel1.Height) / 2), 0)
            Panel1.Left = Round(((Me.Width - Panel1.Width) / 2), 0)

            Me.Refresh()
        End If
        currentindex = currentindex - 1
        If currentindex < 0 Then
            currentindex = currentindex + maxindex + 1
        End If
        display_image(CStr(filelist.Item(currentindex)))


    End Sub

    'Handles key presses when the image control is in focus.
    Private Sub Navigation_Key_Handler(ByVal sender As System.Object, ByVal keyselected As System.Windows.Forms.KeyEventArgs) Handles Panel1.KeyDown, PictureBox1.KeyDown, MyBase.KeyDown
        Try


            If keyselected.KeyCode = Keys.Escape Then
                Try
                    Timer1.Stop()
                    PictureBox1.Image.Dispose()
                    Me.Close()
                Catch ex As Exception
                    Error_Handler("An """ & ex.Message & """ error occurred while closing the Image Display Screen. The program will attempt to recover shortly.")
                End Try
                keyselected.Handled = True

                Exit Sub
            End If

            If keyselected.KeyCode = Keys.Right Then
                display_image(CStr(filelist.Item(currentindex)))
                keyselected.Handled = True

                Exit Sub
            End If

            If keyselected.KeyCode = Keys.Left Then
                currentindex = currentindex - 2
                If currentindex < 0 Then
                    currentindex = currentindex + maxindex + 1
                End If
                display_image(CStr(filelist.Item(currentindex)))
                keyselected.Handled = True

                Exit Sub
            End If
            keyselected.Handled = True

        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    'Handles mouse clicks when the image control is in focus.
    Private Sub Navigation_Mouse_Handler(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseDown, PictureBox1.MouseDown, MyBase.MouseDown
        Try
            'If double left-click when image control is in focus, select next image
            If e.Clicks = 2 And e.Button = MouseButtons.Left Then
                display_image(CStr(filelist.Item(currentindex)))
            End If

            'If double right-click when image control is in focus, select previous image
            If e.Clicks = 2 And e.Button = MouseButtons.Right Then
                currentindex = currentindex - 2
                If currentindex < 0 Then
                    currentindex = currentindex + maxindex + 1
                End If
                display_image(CStr(filelist.Item(currentindex)))
            End If

        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

End Class
