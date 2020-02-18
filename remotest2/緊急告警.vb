Imports System.Timers
Imports objRemoting.Remot
Imports System.Windows.Forms

Public Class 緊急告警
    Private aTimer As System.Timers.Timer
    Private RunTimer As System.Timers.Timer
    Delegate Sub closeWindow()
    Dim clw As New closeWindow(AddressOf Delegate_closeWindow)
    Public position As String
    Public ipconfig As String
    Private TcpChannel As Integer = 8085
    Private HttpChannel As Integer = 5408
    Private CallMathFunction = "CallMathFunction"
    Public RTU As objModbus

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim fileName As String
        Dim fileNum As Integer
        Dim ListM As New List(Of String)
        NotifyIcon1.Visible = True
        NotifyIcon1.Text = "程式執行中"
        NotifyIcon1.BalloonTipText = "程式執行中"
        NotifyIcon1.ShowBalloonTip(1000)
        Application.DoEvents()

        fileName = CurDir() + "\" + "Paramet.ini"
        If System.IO.File.Exists(fileName) = True Then '檔案存在

            fileNum = FreeFile() '取得檔案編號
            FileOpen(fileNum, fileName, OpenMode.Input) '開啟讀進檔案


            Dim i As Integer
            Do Until EOF(fileNum) '把檔案內容一筆筆，分送到ListM陣列記存
                Dim x = LineInput(fileNum) '一行為一筆方式讀取
                ListM.Add(x)
                i = i + 1
            Loop
            FileClose(fileNum)

            ipconfig = ListM(0)
            position = "Room_" + ListM(1)

        End If

        TryConnect()

    End Sub
    Private Sub SetTimer(ByVal dueTime As Integer, ByVal timer As String)
        If timer = "aTimer" Then
            ' Create a timer with a two second interval.
            aTimer = New System.Timers.Timer(dueTime)
            ' Hook up the Elapsed event for the timer. 
            AddHandler aTimer.Elapsed, AddressOf aTimedEvent
            aTimer.AutoReset = True
            aTimer.Enabled = True
        ElseIf timer = "RunTimer" Then
            RunTimer = New System.Timers.Timer(dueTime)
            ' Hook up the Elapsed event for the timer. 
            AddHandler RunTimer.Elapsed, AddressOf runTimedEvent
            RunTimer.AutoReset = True
            RunTimer.Enabled = True
        End If
    End Sub

    Public Sub TryConnect()

        RTU = CType(Activator.GetObject(GetType(objModbus), "tcp://" & ipconfig & ":" & TcpChannel & "/" & CallMathFunction & ""), objModbus)

        Try

            If RTU.FunNoValueGet(position) <> Nothing Then
                connect()
                SetTimer(500, "aTimer")
            Else
                disconnect()
                SetTimer(2000, "RunTimer")
            End If

        Catch ex As Exception

            disconnect()
            SetTimer(2000, "RunTimer")

        End Try

    End Sub
    Dim popstatus As Boolean
    Dim alarmstatus As Boolean
    Dim ii As Integer
    Dim st As String
    Dim R As Integer
    Private Sub aTimedEvent(source As Object, e As ElapsedEventArgs)

        Try
            connect()

            If RTU.FunNoValueGet(position) = "#" Then

                UpdateUI("找不到位置測點，請開啟設定重設", Label2)
                Label2.BackColor = DefaultBackColor

            ElseIf RTU.FunNoValueGet(position) = "1" Then

                顯示(True)
                If alarmstatus = False Then
                    Select Case R
                        Case 1
                            Me.BackColor = Color.Tomato
                        Case 2
                            Me.BackColor = SystemColors.Control
                            R = 0
                    End Select
                    R = R + 1
                End If

                Marquee()
                UpdateUI("警報已開啟" & st + "已通知車動組", Label2)
                Label2.BackColor = Color.Tomato

            ElseIf RTU.FunNoValueGet(position) = "2" Then
                If alarmstatus = False Then
                    Select Case R
                        Case 1
                            Me.BackColor = Color.Yellow
                        Case 2
                            Me.BackColor = SystemColors.Control
                            R = 0
                    End Select
                    R = R + 1
                End If
                Marquee()
                UpdateUI("警報已確認" & st + "車動組已出發", Label2)
                Label2.BackColor = Color.Yellow

            Else
                Marquee()
                popstatus = False
                alarmstatus = False
                UpdateUI("本機" & st + "車動組" + st & "設備中", Label2)
                Label2.BackColor = DefaultBackColor

            End If
        Catch ex As Exception
            disconnect()
            aTimer.Stop()
            aTimer.Dispose()
            SetTimer(2000, "RunTimer")
            Exit Sub
        End Try


    End Sub

    Private Sub runTimedEvent(source As Object, e As ElapsedEventArgs)

        Try

            If RTU.FunNoValueGet(position) <> Nothing Then
                connect()
                RunTimer.Stop()
                RunTimer.Dispose()
                SetTimer(500, "aTimer")
            End If

        Catch ex As Exception
            disconnect()
        End Try

    End Sub
    Private Sub Marquee()
        Select Case ii
            Case 1
                st = ""
            Case 2
                st = "->"
            Case 3
                st = "-->"
            Case 4
                st = "--->"
            Case 5
                st = "---->"
            Case 6
                st = "----->"
            Case 7
                st = "<-----"
            Case 8
                st = "<----"
            Case 9
                st = "<--"
            Case 10
                st = "<-"
                ii = 0
        End Select
        ii = ii + 1
    End Sub
    Private Sub connect()
        UpdateUI("已連線", Label1)
        Label1.BackColor = Color.LightGreen
    End Sub

    Private Sub disconnect()
        UpdateUI("斷線中", Label1)
        UpdateUI("連線中斷----" & vbCrLf & "請檢查電腦網路或設定是否錯誤!!", Label2)
        Label1.BackColor = Color.Red
    End Sub

    Private Delegate Sub UpdateUICallBack(ByVal newText As String, ByVal c As Control)
    Private Sub UpdateUI(ByVal newText As String, ByVal c As Control)
        If Me.InvokeRequired() Then
            Dim cb As New UpdateUICallBack(AddressOf UpdateUI)
            Me.Invoke(cb, newText, c)
        Else
            If newText = "True" Or newText = "False" Then
                c.Enabled = newText
            Else
                c.Text = newText
            End If
        End If
    End Sub

    Public Delegate Sub Delegate3(Status As Boolean)
    Private Sub 顯示(Status As Boolean)

        If Me.InvokeRequired Then
            Dim d As New Delegate3(AddressOf 顯示)
            Me.Invoke(d, New Object() {Status})
        Else
            If Status = True And popstatus = False Then
                彈出畫面()
                popstatus = Status
            End If
        End If

    End Sub

    Private Sub 彈出畫面()
        Me.Show()
        Me.WindowState = FormWindowState.Normal
    End Sub

    Sub Delegate_closeWindow()
        If Me.InvokeRequired Then
            Dim cw As New closeWindow(AddressOf Delegate_closeWindow)
            Me.Invoke(cw)
        Else
            Me.Close()
        End If
    End Sub

    Private Sub NotifyIcon1_BalloonTipShown(sender As Object, e As System.EventArgs)
        Me.Hide()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        clw()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim f2 As New 設定
        f2.Owner = Me
        f2.ShowDialog()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Process.Start("https://www.google.com.tw/")
    End Sub

    Private Sub Me_Click(sender As Object, e As EventArgs) Handles Me.Click
        alarmstatus = True
        Me.BackColor = SystemColors.Control
    End Sub

    Private Sub 隱藏_Click(sender As Object, e As EventArgs) Handles 隱藏.Click
        Me.Hide()
    End Sub
    Private Sub 主畫面ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 主畫面ToolStripMenuItem.Click
        彈出畫面()
    End Sub

    Private Sub 離開ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 離開ToolStripMenuItem.Click
        clw()
    End Sub
    Private Sub Detec_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()
            NotifyIcon1.Visible = True
            NotifyIcon1.ShowBalloonTip(1000)
            NotifyIcon1.Text = "程式執行中"
        End If

    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        彈出畫面()
    End Sub
End Class
