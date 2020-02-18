Imports System
Imports System.Threading
Public Class Timer_c '每觸發一次就延長,重頭重新計時
    Private Class StateObjClass
        ' 用於保存調用TimerTask的參數
        Public SomeValue As Integer
        Public TimerReference As System.Threading.Timer
        Public TimerRuning As Boolean '計時啟用中
        Public eRespone As String
        Public objRespone() As String
        Public objObjRespone() As Object
    End Class

    Public Event e(ByVal firstTime As String) '往上層傳送e
    Public Event obj(ByVal Content() As String) '往上層傳送obj
    Public Event objObj(ByVal Content() As Object) '往上層傳送obj
    '製作一委派計時器
    Private TimerDelegate As New System.Threading.TimerCallback(AddressOf TimerTask)
    Private Timer1Delegate As New System.Threading.TimerCallback(AddressOf Timer1Task)
    Private Timer2Delegate As New System.Threading.TimerCallback(AddressOf Timer2Task)

    Private StateObj As New StateObjClass
    Public Sub startTimer(ByVal dueTime As Integer) '單計時

        If StateObj.TimerRuning = True Then  '計時器正當啟用中
            Try
                StateObj.TimerReference.Dispose() '停止上一個計時器

            Catch ex As Exception

            End Try


            StateObj = New StateObjClass '另建新計時器，此時StateObj.TimerRuning也因而變False

            'Dim u = StateObj.TimerRuning
        End If

        '無論是時器啟用中或計時器無啟用都是要執行下一動作
        StateObj.TimerRuning = True '設成計時使用中
        StateObj.eRespone = DateTime.Now.ToString("yyyy/MM/dd h:mm:ss.fff") '要回的值
        '*以下技巧:
        '採委派計時，這樣
        'Dim TimerItem As New System.Threading.Timer(TimerDelegate, StateObj,
        '                                         dueTime, 0) '開始執行計時

        'StateObj.TimerReference = TimerItem '把這一計時器放到暫存區的計時器

        StateObj.TimerReference = New System.Threading.Timer(TimerDelegate, StateObj,
                                               dueTime, 0) '開始執行計時



    End Sub

    Private Sub TimerTask(ByVal StateObj As Object)


        Dim State As StateObjClass = CType(StateObj, StateObjClass)


        '往上層傳資料
        Dim Nt = DateTime.Now.ToString("yyyy/MM/dd h:mm:ss.fff")
        RaiseEvent e(StateObj.eRespone + " ; " + Nt)
        ' 請求計時器Dispose
        StateObj.TimerReference.Dispose()
        StateObj.TimerRuning = False
        StateObj = New StateObjClass '還給記存器一個新計時器
    End Sub
    '另一組計時回應方法
    Public Sub startTimer(ByVal dueTime As Integer, ByVal Content() As String) '計時

        If StateObj.TimerRuning = True Then  '計時器正當啟用中
            Try
                StateObj.TimerReference.Dispose() '停止上一個計時器
            Catch ex As Exception

            End Try

            StateObj = New StateObjClass '另建新計時器

        End If



        StateObj.TimerRuning = True '設成計時使用中
        StateObj.objRespone = Content '要往上傳的值是傳入的物
        'Dim TimerItem As New System.Threading.Timer(Timer1Delegate, StateObj,
        '                                         dueTime, 0)
        'StateObj.TimerReference = TimerItem '把這一計時器放到暫存區的計時器

        StateObj.TimerReference = New System.Threading.Timer(Timer1Delegate, StateObj,
                                               dueTime, 0)

    End Sub

    Private Sub Timer1Task(ByVal StateObj As Object)


        Dim State As StateObjClass = CType(StateObj, StateObjClass)

        '使用互鎖class遞增計數器變量
        ' System.Threading.Interlocked.Increment(StateObjClass.SomeValue)


        '往上層傳資料
        RaiseEvent obj(StateObj.objRespone)
        ' 請求計時器Dispose
        StateObj.TimerReference.Dispose()
        StateObj.TimerRuning = False
        StateObj = New StateObjClass '還給記存器一個新計時器
    End Sub

    '另二組計時回應方法
    Public Sub startTimer(ByVal dueTime As Integer, ByVal ContentObj() As Object) '計時

        If StateObj.TimerRuning = True Then  '計時器正當啟用中
            Try
                StateObj.TimerReference.Dispose() '停止上一個計時器
            Catch ex As Exception

            End Try

            StateObj = New StateObjClass '另建新計時器

        End If



        StateObj.TimerRuning = True '設成計時使用中
        StateObj.objObjRespone = ContentObj '要往上傳的值是傳入的物
        'Dim TimerItem As New System.Threading.Timer(Timer2Delegate, StateObj,
        '                                         dueTime, 0)
        'StateObj.TimerReference = TimerItem '把這一計時器放到暫存區的計時器
        StateObj.TimerReference = New System.Threading.Timer(Timer2Delegate, StateObj,
                                               dueTime, 0)


    End Sub

    Private Sub Timer2Task(ByVal StateObj As Object)


        Dim State As StateObjClass = CType(StateObj, StateObjClass)

        '使用互鎖class遞增計數器變量
        ' System.Threading.Interlocked.Increment(StateObjClass.SomeValue)


        '往上層傳資料
        RaiseEvent objObj(StateObj.objObjRespone)
        ' 請求計時器Dispose
        StateObj.TimerReference.Dispose()
        StateObj.TimerRuning = False
        StateObj = New StateObjClass '還給記存器一個新計時器
    End Sub

    Public Sub StopTimer()

        If StateObj.TimerRuning = True Then  '計時器正當啟用中
            StateObj.TimerReference.Dispose() '停止上一個計時器
            StateObj.TimerRuning = False
            StateObj = New StateObjClass '還給記存器一個新計時器,否則下次叫用計時會錯敗
        End If
    End Sub
End Class
