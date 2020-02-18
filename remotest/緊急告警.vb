Imports System
Imports System.Timers
Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Tcp
Imports System.Runtime.Remoting.Channels.Http
Imports System.Runtime.InteropServices
Imports objRemoting.Remot
Imports System.ServiceProcess


Public Class 緊急告警
    Private aTimer As System.Timers.Timer
    Private RunTimer As System.Timers.Timer
    Dim fileName As String
    Dim fileNum As Integer
    Dim ListM As New List(Of String)
    Delegate Sub closeWindow()
    Dim clw As New closeWindow(AddressOf Delegate_closeWindow)
    Public position As String
    Public ipconfig As String
    Private TcpChannel As Integer = 8085
    Private HttpChannel As Integer = 5408
    Private CallMathFunction = "CallMathFunction"
    Public RTU As objModbus

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

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

        'Dim trd As New Threading.Thread(New Threading.ThreadStart(AddressOf CreateExcel))

        'trd.Start()

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
    Dim ii As Integer
    Dim st As String
    ' The event handler for the Timer.Elapsed event. 
    Private Sub aTimedEvent(source As Object, e As ElapsedEventArgs)

        Try
            connect()

            If RTU.FunNoValueGet(position) = "#" Then

                UpdateUI("找不到位置測點，請重新設定", Label2)
                Label2.BackColor = DefaultBackColor

            ElseIf RTU.FunNoValueGet(position) = "1" Then

                Marquee()
                UpdateUI("警報已開啟" & st + "已通知車動組", Label2)
                Label2.BackColor = Color.Tomato

            ElseIf RTU.FunNoValueGet(position) = "2" Then

                Marquee()
                UpdateUI("警報已確認" & st + "車動組已出發", Label2)
                Label2.BackColor = Color.Yellow

            Else

                Marquee()
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
        Button1.BackColor = SystemColors.Control
        Button4.BackColor = SystemColors.Control
        UpdateUI(True, Button1)
        UpdateUI(True, Button4)
    End Sub

    Private Sub disconnect()
        UpdateUI("斷線中", Label1)
        UpdateUI("連線中斷---請檢查設定是否錯誤!!", Label2)
        Label1.BackColor = Color.Red
        Button1.BackColor = Color.Red
        Button4.BackColor = Color.Red
        UpdateUI(False, Button1)
        UpdateUI(False, Button4)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim result As DialogResult = MessageBox.Show("你確定要通知車動組?", "警告", MessageBoxButtons.OKCancel)
        If result = DialogResult.Cancel Then

        ElseIf result = DialogResult.OK Then
            RTU.FunNoValueUpdata(position, 1)
            Label1.BackColor = Color.Red
            UpdateUI(False, Button1)
            MessageBox.Show("已通知機動組", "通知")
        End If

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim f2 As New 設定
        f2.Owner = Me
        f2.ShowDialog()
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
    Sub Delegate_closeWindow()
        If Me.InvokeRequired Then

            Dim cw As New closeWindow(AddressOf Delegate_closeWindow)
            Me.Invoke(cw)

        Else
            Me.Close()

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        clw()
    End Sub

    Private Sub 彈出畫面()
        Me.Show()
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub 隱藏_Click(sender As Object, e As EventArgs) Handles 隱藏.Click
        Me.Hide()
    End Sub

    Private Sub 主畫面ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 主畫面ToolStripMenuItem.Click
        彈出畫面()
    End Sub

    Private Sub 離開ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 離開ToolStripMenuItem.Click
        clw()
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        彈出畫面()
    End Sub

    Private Sub Detec_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()
            NotifyIcon1.Visible = True
            NotifyIcon1.ShowBalloonTip(1000)
            NotifyIcon1.Text = "程式執行中"
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim result As DialogResult = MessageBox.Show("你確定要關閉警報?", "警告", MessageBoxButtons.OKCancel)
        If result = DialogResult.Cancel Then

        ElseIf result = DialogResult.OK Then
            RTU.FunNoValueUpdata(position, 0)
            Label1.BackColor = Color.Red
            UpdateUI(True, Button1)
            MessageBox.Show("已關閉警報", "通知")
        End If
    End Sub
End Class
