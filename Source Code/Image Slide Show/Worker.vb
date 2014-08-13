Imports System.IO




Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public processfolderinfo As String
    Public interval As String
    Public fullscreen As Boolean
    Public randomize As Boolean
    Public recursive As Boolean

    Public result As String = ""

    Private filelist As ArrayList

    Public Event WorkerComplete(ByVal Result As String)
    Public Event WorkerProgress(ByVal value As Integer)


#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        filelist = New ArrayList
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As String)
        Try
            Dim Display_Message1 As New Display_Message(message)
            Display_Message1.ShowDialog()
            MsgBox(message)
        Catch ex As Exception
            MsgBox("An error occurred in Image Slide Show's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    ' Sets the thread using the AddressOf the subroutine where
                    ' the thread will start.
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerExecute)
                    ' Starts the thread.
                    WorkerThread.Start()

            End Select
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Public Sub WorkerExecute()
        Try
            filelist.Clear()
            ProcessPath(processfolderinfo)



            RaiseEvent WorkerProgress(filelist.Count)
            If filelist.Count > 0 Then
                If randomize = True Then
                    ShuffleList()
                End If
                Dim slideshow As Image_Display
                slideshow = New Image_Display(filelist, interval, fullscreen)

                slideshow.ShowDialog()
            End If
            result = "Success"
            RaiseEvent WorkerComplete(result)
        Catch ex As Exception
            result = "Failure"
            RaiseEvent WorkerComplete(result)
        End Try

        WorkerThread.Abort()
    End Sub



    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True
            myProcess.Start()
            Dim sIn As StreamWriter = myProcess.StandardInput
            sIn.AutoFlush = True

            Dim sOut As StreamReader = myProcess.StandardOutput
            Dim sErr As StreamReader = myProcess.StandardError
            sIn.Write(AppToRun & _
               System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()
            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            'MessageBox.Show("The 'dir' command window was closed at: " & myProcess.ExitTime & "." & System.Environment.NewLine & "Exit Code: " & myProcess.ExitCode)

            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()
            'MessageBox.Show(s)
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while launching DOS shell. The program will attempt to recover shortly.")
        End Try
        Return s
    End Function

    Public Sub ProcessPath(ByVal inputpath As String)
        
        Dim path As String
        path = inputpath
        If Directory.Exists(path) Then
            ProcessDirectory(path)
        Else
        End If
    End Sub


    Public Sub ProcessDirectory(ByVal targetDirectory As String)

        Dim fileEntries As String() = Directory.GetFiles(targetDirectory)


        Dim fileName As String

        For Each fileName In fileEntries
            Dim finfo As FileInfo = New FileInfo(fileName)
            'MsgBox(finfo.Extension)
            Select Case finfo.Extension.ToLower()
                Case ".gif"
                    filelist.Add(finfo.FullName)
                    '       MsgBox("added " & finfo.FullName)
                Case ".jpg"
                    filelist.Add(finfo.FullName)
                    '      MsgBox("added " & finfo.FullName)
                Case ".bmp"
                    filelist.Add(finfo.FullName)
                    '     MsgBox("added " & finfo.FullName)
                Case ".png"
                    filelist.Add(finfo.FullName)
                    '    MsgBox("added " & finfo.FullName)
            End Select
        Next fileName

        If recursive = True Then
            Dim subdirectoryEntries As String() = Directory.GetDirectories(targetDirectory)
            Dim subdirectory As String
            For Each subdirectory In subdirectoryEntries
                ProcessDirectory(subdirectory)
            Next subdirectory
        End If

    End Sub 'ProcessDirectory

    Private Sub ShuffleList()
        Try

        
        Dim runner As Integer
            Dim templist As ArrayList = New ArrayList(filelist.Count)
            For runner = 0 To filelist.Count - 1
                templist.Add("@")
            Next
            Dim fileplaced As Boolean
            Dim toplace As Integer
            toplace = 0
            Dim obj As New System.Random

            For runner = 0 To filelist.Count - 1
                fileplaced = False
                toplace = obj.Next(0, filelist.Count)

                If toplace > filelist.Count - 1 Then
                    toplace = 0
                End If
                While fileplaced = False

                    If CStr(templist.Item(toplace)) = "@" Then
                        templist.Item(toplace) = filelist.Item(runner)
                        fileplaced = True

                    End If
                    toplace = toplace + 1
                    If toplace > filelist.Count - 1 Then
                        toplace = 0
                    End If
                End While

            Next
            
            For runner = 0 To filelist.Count - 1
                filelist.Item(runner) = templist.Item(runner)
            Next
    

        Catch ex As Exception
            Error_Handler(ex.ToString)
        End Try
    End Sub


End Class
