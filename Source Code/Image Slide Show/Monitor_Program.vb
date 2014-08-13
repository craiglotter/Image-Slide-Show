Imports Microsoft.Win32
Imports System.IO

Public Class Monitor_Program
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker
    Public Delegate Sub WorkerhHandler(ByVal Result As String)
    Public Delegate Sub WorkerProgresshHandler(ByVal Result As String)

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        AddHandler Worker1.WorkerProgress, AddressOf WorkerProgressHandler
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Monitor_Program))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Label9 = New System.Windows.Forms.Label
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.Label4 = New System.Windows.Forms.Label
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gray
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.PictureBox5)
        Me.Panel1.Controls.Add(Me.PictureBox4)
        Me.Panel1.Controls.Add(Me.PictureBox3)
        Me.Panel1.Controls.Add(Me.PictureBox2)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(8, 31)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(306, 153)
        Me.Panel1.TabIndex = 11
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(104, 128)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(64, 16)
        Me.Label16.TabIndex = 42
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(8, 128)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(96, 16)
        Me.Label17.TabIndex = 41
        Me.Label17.Text = "Recursive Listing:"
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(104, 112)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(64, 16)
        Me.Label14.TabIndex = 40
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(8, 112)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(72, 16)
        Me.Label15.TabIndex = 39
        Me.Label15.Text = "Randomize:"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(104, 96)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(64, 16)
        Me.Label8.TabIndex = 38
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(8, 96)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(72, 16)
        Me.Label13.TabIndex = 37
        Me.Label13.Text = "Full Screen:"
        '
        'Button1
        '
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Location = New System.Drawing.Point(224, 120)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(72, 24)
        Me.Button1.TabIndex = 36
        Me.Button1.Text = "Start Show"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(104, 80)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(192, 16)
        Me.Label12.TabIndex = 35
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(104, 64)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(192, 16)
        Me.Label11.TabIndex = 34
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(104, 48)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(192, 16)
        Me.Label10.TabIndex = 33
        Me.ToolTip1.SetToolTip(Me.Label10, "d")
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.White
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Gray
        Me.Label7.Location = New System.Drawing.Point(16, 10)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 26)
        Me.Label7.TabIndex = 32
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 80)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(272, 16)
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "Slide Interval:"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 64)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(272, 16)
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "Images:"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(272, 16)
        Me.Label3.TabIndex = 29
        Me.Label3.Text = "Image Folder:"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.DarkOrange
        Me.Label1.Location = New System.Drawing.Point(224, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 16)
        Me.Label1.TabIndex = 27
        Me.Label1.Text = "Stopped"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PictureBox5
        '
        Me.PictureBox5.BackgroundImage = CType(resources.GetObject("PictureBox5.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(200, 16)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox5.TabIndex = 26
        Me.PictureBox5.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.BackgroundImage = CType(resources.GetObject("PictureBox4.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(184, 16)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox4.TabIndex = 25
        Me.PictureBox4.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackgroundImage = CType(resources.GetObject("PictureBox3.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(168, 16)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox3.TabIndex = 24
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackgroundImage = CType(resources.GetObject("PictureBox2.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(152, 16)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox2.TabIndex = 23
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(136, 16)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox1.TabIndex = 22
        Me.PictureBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Location = New System.Drawing.Point(8, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(288, 32)
        Me.Label2.TabIndex = 28
        Me.Label2.Text = "Label2"
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Timer2
        '
        Me.Timer2.Interval = 1000
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.Gray
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.ForeColor = System.Drawing.Color.White
        Me.Label9.Location = New System.Drawing.Point(136, 12)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(48, 16)
        Me.Label9.TabIndex = 34
        Me.Label9.Text = "Config"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenu = Me.ContextMenu1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Image Slide Show"
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2, Me.MenuItem3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Display Application"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.Text = "Exit Application"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.DimGray
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(8, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(120, 24)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Image Slide Show"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Monitor_Program
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Silver
        Me.ClientSize = New System.Drawing.Size(322, 192)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label4)
        Me.ForeColor = System.Drawing.Color.White
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(328, 224)
        Me.MinimumSize = New System.Drawing.Size(328, 224)
        Me.Name = "Monitor_Program"
        Me.ShowInTaskbar = False
        Me.Text = "Image Slide Show"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region






    Private current_light As Integer = 0
    Private current_colour As Integer = 0
    Private currently_working As Boolean = False

    Public processfolderinfo As String
    Public interval As String
    Public fullscreen As Boolean
    Public randomize As Boolean
    Public recursive As Boolean

    Private WithEvents cform As Config



    Private Sub Error_Handler(ByVal message As String)
        Try
            Dim Display_Message1 As New Display_Message(message)
            Display_Message1.ShowDialog()
        Catch ex As Exception
            MsgBox("An error occurred in Image Slide Show's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Public Sub Load_Registry_Values()
        Try


            Dim configflag As Boolean
            configflag = False
            Dim str As String
            Dim keyflag1 As Boolean = False
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim keys() As String = oReg.GetSubKeyNames()
            System.Array.Sort(keys)

            For Each str In keys
                If str.Equals("Software\Image Slide Show") = True Then
                    keyflag1 = True
                    Exit For
                End If
            Next str

            If keyflag1 = False Then
                oReg.CreateSubKey("Software\Image Slide Show")
            End If

            keyflag1 = False

            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\Image Slide Show", True)

            str = oKey.GetValue("processfolderinfo")
            If Not IsNothing(str) And Not (str = "") Then
                processfolderinfo = str
            Else
                configflag = True
                oKey.SetValue("processfolderinfo", "C:\")
                processfolderinfo = "C:\"
            End If

            str = oKey.GetValue("interval")
            If Not IsNothing(str) And Not (str = "") Then
                interval = str
            Else
                configflag = True
                oKey.SetValue("interval", "1")
                interval = "1"
            End If

            str = oKey.GetValue("fullscreen")
            If Not IsNothing(str) And Not (str = "") Then
                fullscreen = CBool(str)
            Else
                configflag = True
                oKey.SetValue("fullscreen", "False")
                fullscreen = False
            End If

            str = oKey.GetValue("randomize")
            If Not IsNothing(str) And Not (str = "") Then
                randomize = CBool(str)
            Else
                configflag = True
                oKey.SetValue("randomize", "False")
                randomize = False
            End If

            str = oKey.GetValue("recursive")
            If Not IsNothing(str) And Not (str = "") Then
                recursive = CBool(str)
            Else
                configflag = True
                oKey.SetValue("recursive", "False")
                recursive = False
            End If

            If configflag = True Then
                Dim cform As New Config
                Dim result As DialogResult
                result = cform.ShowDialog()
                If result = DialogResult.OK Then
                    Load_Registry_Values()
                End If
                cform.Dispose()
            End If

            oKey.Close()
            oReg.Close()
            If processfolderinfo.Length > 32 Then
                Label10.Text = processfolderinfo.Substring(0, 32) & "..."
            Else
                Label10.Text = processfolderinfo
            End If


            ToolTip1.SetToolTip(Label10, processfolderinfo)
            Label12.Text = interval & " seconds"
            Label8.Text = fullscreen.ToString
            Label14.Text = randomize.ToString
            Label16.Text = recursive.ToString

            'Timer1.Stop()
            'Timer1.Interval = interval * 60 * 1000
            'Label7.Text = "Next Check: " & Now().AddMinutes(Double.Parse(interval)).ToString
            'NotifyIcon1.Text = "Resting. Next Check: " & Now().AddMinutes(Double.Parse(interval)).ToString
            'Timer1.Start()
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while loading required registry values. The program will attempt to recover shortly.")
        End Try
    End Sub



    Private Sub mConfigUpdate() Handles cform.ConfigUpdate
        Try
            Load_Registry_Values()
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while implementing Configuration changes. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub run_green_lights()
        Try
            Label1.ForeColor = Color.LimeGreen
            Label1.Text = "Slide Show"

            current_light = current_light - 1
            If current_light < 1 Then
                current_light = 5
            End If
            current_colour = 0
            PictureBox1.Image = ImageList1.Images(1)
            PictureBox2.Image = ImageList1.Images(1)
            PictureBox3.Image = ImageList1.Images(1)
            PictureBox4.Image = ImageList1.Images(1)
            PictureBox5.Image = ImageList1.Images(1)

            Select Case current_light
                Case 0

                    PictureBox1.Image = ImageList1.Images(0)
                Case 1

                    PictureBox2.Image = ImageList1.Images(0)
                Case 2

                    PictureBox3.Image = ImageList1.Images(0)
                Case 3

                    PictureBox4.Image = ImageList1.Images(0)
                Case 4

                    PictureBox5.Image = ImageList1.Images(0)
                Case 5

                    PictureBox1.Image = ImageList1.Images(0)
            End Select

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while changing status light colour. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub run_orange_lights()
        Try
            Label1.ForeColor = Color.DarkOrange
            Label1.Text = "Stopped"

            current_light = current_light - 1
            If current_light < 1 Then
                current_light = 5
            End If
            current_colour = 1
            PictureBox1.Image = ImageList1.Images(3)
            PictureBox2.Image = ImageList1.Images(3)
            PictureBox3.Image = ImageList1.Images(3)
            PictureBox4.Image = ImageList1.Images(3)
            PictureBox5.Image = ImageList1.Images(3)
            Select Case current_light
                Case 0
                    PictureBox1.Image = ImageList1.Images(2)
                Case 1
                    PictureBox2.Image = ImageList1.Images(2)
                Case 2
                    PictureBox3.Image = ImageList1.Images(2)
                Case 3
                    PictureBox4.Image = ImageList1.Images(2)
                Case 4
                    PictureBox5.Image = ImageList1.Images(2)
                Case 5
                    PictureBox1.Image = ImageList1.Images(2)
            End Select

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while changing status light colour. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub run_lights()
        Try
            If current_colour = 1 Then
                Select Case current_light
                    Case 0
                        PictureBox5.Image = ImageList1.Images(3)
                        PictureBox1.Image = ImageList1.Images(2)
                    Case 1
                        PictureBox1.Image = ImageList1.Images(3)
                        PictureBox2.Image = ImageList1.Images(2)
                    Case 2
                        PictureBox2.Image = ImageList1.Images(3)
                        PictureBox3.Image = ImageList1.Images(2)
                    Case 3
                        PictureBox3.Image = ImageList1.Images(3)
                        PictureBox4.Image = ImageList1.Images(2)
                    Case 4
                        PictureBox4.Image = ImageList1.Images(3)
                        PictureBox5.Image = ImageList1.Images(2)
                    Case 5
                        PictureBox5.Image = ImageList1.Images(3)
                        PictureBox1.Image = ImageList1.Images(2)
                End Select
            Else
                Select Case current_light
                    Case 0
                        PictureBox5.Image = ImageList1.Images(1)
                        PictureBox1.Image = ImageList1.Images(0)
                    Case 1
                        PictureBox1.Image = ImageList1.Images(1)
                        PictureBox2.Image = ImageList1.Images(0)
                    Case 2
                        PictureBox2.Image = ImageList1.Images(1)
                        PictureBox3.Image = ImageList1.Images(0)
                    Case 3
                        PictureBox3.Image = ImageList1.Images(1)
                        PictureBox4.Image = ImageList1.Images(0)
                    Case 4
                        PictureBox4.Image = ImageList1.Images(1)
                        PictureBox5.Image = ImageList1.Images(0)
                    Case 5
                        PictureBox5.Image = ImageList1.Images(1)
                        PictureBox1.Image = ImageList1.Images(0)
                End Select

            End If

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while loading the status light strip. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try
            run_lights()
            Label7.Text = "System Time: " & vbCrLf & Format(Now(), "dd/MM/yyyy hh:mm:ss tt")
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while loading the status light strip. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub Monitor_Program_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Load_Registry_Values()
            run_orange_lights()
            Timer2.Start()
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while loading the application screen. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub monitor_program_closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            Worker1.Dispose()
            NotifyIcon1.Dispose()
            Application.Exit()
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while closing the Monitoring screen. The program will attempt to recover shortly.")
        End Try
    End Sub



    Public Sub WorkerHandler(ByVal Result As String)
        Try
            'show_application()
            Button1.Text = "Start Show"
            currently_working = False
            run_orange_lights()
            Me.Refresh()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Public Sub WorkerProgressHandler(ByVal value As Integer)
        Try
            ' Displays the returned value in the appropriate label.
            Label11.Text = value & " images located"
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub run_worker()
        run_green_lights()
        Worker1.processfolderinfo = processfolderinfo
        Worker1.interval = interval
        Worker1.fullscreen = fullscreen
        Worker1.randomize = randomize
        Worker1.recursive = recursive
        
        Worker1.ChooseThreads(1)
        currently_working = True
    End Sub

    Private Sub stop_worker()
        run_orange_lights()
        Worker1.WorkerThread.Abort()
        currently_working = False
        Me.Refresh()
    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click
        Try
            Dim cform As New Config
            Dim result As DialogResult
            result = cform.ShowDialog()
            If result = DialogResult.OK Then
                Load_Registry_Values()
            End If
            cform.Dispose()
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while loading the configuration screen. The program will attempt to recover shortly.")
        End Try
    End Sub



    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Try
            Worker1.Dispose()
            NotifyIcon1.Dispose()
            Application.Exit()
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while closing Image Slide Show. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub show_application()
        Try
            'Me.Show()
            Me.Opacity = 1

            Me.BringToFront()
            Me.Refresh()
            Me.WindowState = FormWindowState.Normal

        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while trying to display the main screen. The program will attempt to recover shortly.")
        End Try
    End Sub

    Private Sub NotifyIcon1_dblclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        show_application()
    End Sub
    Private Sub NotifyIcon1_snglclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        show_application()
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        show_application()
    End Sub

    Private Sub monitor_program_resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try
            If Me.WindowState = FormWindowState.Minimized Then
                NotifyIcon1.Visible = True
                Me.Opacity = 0
            End If
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while trying to display the main screen. The program will attempt to recover shortly.")
        End Try
    End Sub






    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If currently_working = False Then
                Button1.Text = "Stop Show"
                run_worker()

            Else
                Button1.Text = "Start Show"
                stop_worker()
            End If
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while launching the slide show. The program will attempt to recover shortly.")
        End Try
    End Sub


End Class
