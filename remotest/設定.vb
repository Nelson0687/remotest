Public Class 設定
    Dim fileName As String
    Dim fileNum As Integer
    Dim ListM As New List(Of String)
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles Me.Load
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

            TextBox1.Text = ListM(0)
            TextBox2.Text = ListM(1)

        End If

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ListM(0) = TextBox1.Text
        ListM(1) = TextBox2.Text
        saveFile(ListM)
        Dim Mut As New System.Threading.Mutex(True, Application.ProductName)
        Mut.ReleaseMutex()
        Application.Restart()
    End Sub
    Public Sub saveFile(ByVal ListData As List(Of String)) '存檔 Paramet.ini
        FileClose(fileNum)

        fileNum = FreeFile()
        FileOpen(fileNum, fileName, OpenMode.Output)
        For Each d In ListData
            PrintLine(fileNum, d)
        Next
        FileClose(fileNum)
    End Sub
End Class