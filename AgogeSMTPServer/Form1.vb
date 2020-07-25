Imports System.IO
Imports System.Threading
Imports System.Net
Imports System.Net.Sockets
Imports System.Text

Public Class Form1
    'Add a TEXTBOX1 to form
    'Add a TEXTBOX1 to form
    'Add a TEXTBOX1 to form
    'Add a TEXTBOX1 to form

    Private Server As Sockets.TcpListener
    Private Port As Integer = 25
    Private Thread As Thread

    Delegate Sub LogCallback(S As String)

    Sub Log(S As String, Optional AppendCRLF As Boolean = False)

        If Me.TextBox1.InvokeRequired Then
            Dim d As LogCallback = AddressOf Log
            Me.Invoke(d, S)
        Else
            Me.TextBox1.AppendText(S)

            My.Computer.FileSystem.WriteAllText("logs\LOG.txt", S, True)
        End If
    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Me.Text = String.Format("Agoge SMTP Server Listening port {0} v{1}", Port, AppVersion)

        If Not My.Computer.FileSystem.DirectoryExists("Logs") Then
            My.Computer.FileSystem.CreateDirectory("Logs")
        End If

        Thread = New Thread(New ThreadStart(AddressOf Listening))
        Thread.Start()
    End Sub





    Private FlagClosing As Boolean = False

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'If Not FlagClosing Then
        '    e.Cancel = True
        '    Me.WindowState = FormWindowState.Minimized
        'End If
    End Sub

    Private FlagNotify As Boolean = False

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()

            If Not FlagNotify Then
                NotifyIcon1.BalloonTipText = "Agoge SMTP Server is minimized to system tray"
                NotifyIcon1.ShowBalloonTip(500)
                FlagNotify = True
            End If
        End If
    End Sub

    Private Sub NotifyIcon1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        Me.Show()
        Me.WindowState = FormWindowState.Normal
    End Sub





    Async Function Listening() As Task
        Server = New Sockets.TcpListener(Net.IPAddress.Any, Port)
        Server.Start()

        Do
            Dim HandleThread As New Thread(New ParameterizedThreadStart(AddressOf HandleConnection))
            HandleThread.Start(Await Server.AcceptTcpClientAsync)
        Loop

    End Function





    Class ReadTimeoutException
        Inherits Exception
    End Class

    Class WriteException
        Inherits Exception
    End Class

    Class ConnectionLost
        Inherits Exception
    End Class





    Async Function HandleConnection(TcpClient As TcpClient) As Task
        Using Stream As NetworkStream = TcpClient.GetStream
            Using WStream As New StreamWriter(Stream, Encoding.ASCII) With {.AutoFlush = True}
                Using RStream As New StreamReader(Stream, Encoding.ASCII)

                    Dim RAW As String = String.Empty
                    Dim MAILFROM As String = String.Empty
                    Dim RCPTO As String = String.Empty
                    Dim MAILDATA As String = String.Empty

                    Dim FlagMAILDATA As Boolean = False
                    Dim LastMessage As String = String.Empty
                    Dim Message As String = String.Empty

                    Try
                        Await R(WStream, RAW, "220 raptor.agoge.com.br")

                        While True

                            Dim TRead As Task(Of String) = RStream.ReadLineAsync
                            Dim TOut As Task = Task.Delay(60 * 1000)

                            Await Task.Factory.ContinueWhenAny(New Task() {TRead, TOut},
                                Sub(t)
                                    If t.Equals(TOut) Then
                                        TcpClient.Close()
                                        Throw New ReadTimeoutException
                                    End If
                                End Sub)

                            LastMessage = Message
                            Message = Await TRead
                            RAW = String.Concat(RAW, Message, vbCrLf)



                            If FlagMAILDATA Then
                                If Message = "." And LastMessage = String.Empty Then
                                    FlagMAILDATA = False
                                Else
                                    MAILDATA = String.Concat(MAILDATA, Message)
                                End If
                            End If



                            If Not FlagMAILDATA Then
                                If Message.StartsWith("HELO", True, Nothing) Then
                                    Await R(WStream, RAW, "250 raptor.agoge.com.br")
                                ElseIf Message.StartsWith("MAIL", True, Nothing) Then
                                    Await R(WStream, RAW, "250 OK")
                                    MAILFROM = Message.Replace("<", String.Empty).Replace(">", String.Empty).Replace(":", " ").Replace(vbCrLf, String.Empty)
                                ElseIf Message.StartsWith("RCPT", True, Nothing) Then
                                    Await R(WStream, RAW, "250 OK")
                                    RCPTO = Message.Replace("<", String.Empty).Replace(">", String.Empty).Replace(":", " ").Replace(vbCrLf, String.Empty)
                                ElseIf Message.StartsWith("Date", True, Nothing) Or Message.StartsWith("From", True, Nothing) Then
                                    Await R(WStream, RAW, "250 OK")
                                ElseIf Message.StartsWith("DATA", True, Nothing) Then
                                    Await R(WStream, RAW, "354 Start mail input; end with <CRLF>.<CRLF>")
                                    FlagMAILDATA = True
                                ElseIf Message.StartsWith("From", True, Nothing) Or Message.StartsWith("Received", True, Nothing) Then
                                    Await R(WStream, RAW, "250 OK")
                                ElseIf Message = "." And LastMessage = String.Empty Then
                                    Await R(WStream, RAW, "250 OK")
                                ElseIf Message.StartsWith("QUIT", True, Nothing) Then
                                    Await R(WStream, RAW, "221 raptor.agoge.com.br closing transmission channel")
                                    Exit While
                                Else
                                    Await R(WStream, RAW, "502 Command not implemented")
                                End If
                            End If

                        End While

                        SaveMail(MAILFROM, RCPTO, RAW)

                    Catch ex As ReadTimeoutException
                        If String.IsNullOrEmpty(MAILDATA) Then
                            Log(String.Format("{0:yyyy-MM-dd hh-mm-ss} Client {1} timeout {2} {3}{4}", Now, TcpClient.Client.RemoteEndPoint.ToString, MAILFROM, RCPTO, vbCrLf))
                        Else
                            SaveMail(MAILFROM, RCPTO, RAW)
                        End If
                    Catch ex As ConnectionLost
                        Log(String.Format("{0:yyyy-MM-dd hh-mm-ss} Client {1} connection lost {2}", Now, TcpClient.Client.RemoteEndPoint.ToString, vbCrLf))
                    Catch ex As WriteException
                        Log(String.Format("{0:yyyy-MM-dd hh-mm-ss} Client {1} can't write {2}", Now, TcpClient.Client.RemoteEndPoint.ToString, vbCrLf))
                    Catch ex As Exception
                        Log(String.Format("{1}{0:yyyy-MM-dd hh-mm-ss} Exception: {2}{1}Stack trace:{1}{3}{1}{1}", Now, vbCrLf, ex.Message, ex.StackTrace))
                    Finally
                        TcpClient.Close()
                    End Try

                End Using
            End Using
        End Using
    End Function

    Sub SaveMail(MAILFROM, RCPTO, RAW)
        Dim mailId As String = String.Format("{0:yyyy-MM-dd hh-mm-ss} {1} {2}.mail", Now, MAILFROM, RCPTO)
        My.Computer.FileSystem.WriteAllText(String.Format("logs\{0}", mailId), RAW, False)
        Log(String.Concat(mailId, vbCrLf))
    End Sub

    Private AppVersion As Integer = 80

    Async Function R(WStream As StreamWriter, RAW As String, Message As String) As Task

        Try
            Await WStream.WriteLineAsync(Message)
            RAW = String.Concat(RAW, Message, vbCrLf)
        Catch ex As Exception
            Throw New WriteException
        End Try

    End Function


End Class

