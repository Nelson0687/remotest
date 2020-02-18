Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Linq
Imports System.Collections.Generic
Imports System.Data
Imports System.Threading
Imports System.Xml.Linq
'Imports Modbus

Imports System
Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Tcp
Imports System.Runtime.Remoting.Channels.Http
Imports objRemoting.Remot
Imports System.IO

'Imports WebReference.ServiceRTU


' 若要允許使用 ASP.NET AJAX 從指令碼呼叫此 Web 服務，請取消註解下一行。
<WebService(Namespace:="http://tempuri.org/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class ServiceRTU
    Inherits System.Web.Services.WebService

    Public Shared objMainRun As objRemoting.Remot.objModbus
    ' Public Shared objMainRun As WebReference.ServiceRTU



    Public Sub Start()
        Dim t_Start As New Timer(AddressOf Start_Call)
        t_Start.Change(2000, 1000)
    End Sub

    Private Sub Start_Call(ByVal state As Object)
        Dim t As Timer = CType(state, Timer)
        t.Dispose()


    End Sub

    Public Function RemoteTest(ByVal yourName As String) As String

        Return yourName & " 已經成功銜接遠端,測試成功"

    End Function

    Public Sub New()
        '* 不要在此呼叫動作

        '這是與WindowsServiceRTU所建立的 MainRunRemoting類別建立起通道,才透過objRemoting進行Modbus運作
        '   objMainRun = CType(Activator.GetObject( _
        '   GetType(objRemoting.Remot.objModbus), _
        '   "tcp://localhost:8085/CallMathFunction"),  _
        'objModbus)

        ' Start()
    End Sub

#Region "Soyal卡機命令"
    Public Sub UserToSoyalAll(ByVal DeviceS() As String, ByVal TagId As String, ByVal PersonNo As String, ByVal Name As String,
                                    ByVal Sex As String, ByVal Role As String, ByVal Department As String, ByVal _Class As String, ByVal PinCode As String,
                                     ByVal AccMode As String, ByVal Patrol As Boolean, ByVal ExpireCheck As Boolean, ByVal AntiPasBack As Boolean,
                                     ByVal WGZone As Boolean, ByVal idexZone As Integer, ByVal PswChang As Boolean, ByVal Group1 As String, ByVal Group2 As String,
                                    ByVal ValidDate As String, ByVal LevelRang As Integer, ByVal LeveWGZone As Integer, ByVal ExtAntiPasBack As Boolean) '將資料載入卡機 83H
        objMainRun.UserToSoyalAll(DeviceS, TagId, PersonNo, Name, Sex, Role, Department, _Class, PinCode, AccMode, Patrol,
                            ExpireCheck, AntiPasBack, PswChang, WGZone, idexZone, Group1, Group2, ValidDate, LevelRang, LeveWGZone,
                            ExtAntiPasBack)


    End Sub
    <WebMethod()>
    Public Sub UserToSoyalAll(ByVal DeviceNo As String, ByVal TagId As String, ByVal PersonNo As String, ByVal Name As String,
                                    ByVal Sex As String, ByVal Role As String, ByVal Department As String, ByVal _Class As String, ByVal PinCode As String,
                                     ByVal AccMode As String, ByVal Patrol As Boolean, ByVal ExpireCheck As Boolean, ByVal AntiPasBack As Boolean,
                                     ByVal WGZone As Boolean, ByVal idexZone As Integer, ByVal PswChang As Boolean, ByVal Group1 As String, ByVal Group2 As String,
                                    ByVal ValidDate As String, ByVal LevelRang As Integer, ByVal LeveWGZone As Integer, ByVal ExtAntiPasBack As Boolean) '將資料載入卡機 83H
        objMainRun.UserToSoyalAll(DeviceNo.Split(","), TagId, PersonNo, Name, Sex, Role, Department, _Class, PinCode, AccMode, Patrol,
                            ExpireCheck, AntiPasBack, PswChang, WGZone, idexZone, Group1, Group2, ValidDate, LevelRang, LeveWGZone,
                            ExtAntiPasBack)


    End Sub
    <WebMethod()>
    Public Sub UserToSoyal(ByVal DeviceNo As String, ByVal idexCard As Integer, ByVal TagId As String, ByVal PersonNo As String,
                                  ByVal Name As String, ByVal Sex As String, ByVal Role As String, ByVal Department As String, ByVal _Class As String,
                                  ByVal PinCode As String, ByVal AccMode As String, ByVal _Patrol As Boolean, ByVal ExpireCheck As Boolean,
                                  ByVal AntiPasBack As Boolean, ByVal _PswChang As Boolean, ByVal WGZone As Boolean, ByVal idexZone As Integer,
                                  ByVal Group1 As String, ByVal Group2 As String, ByVal ValidDate As String, ByVal _LevelRang As Integer, ByVal LeveWGZone As Integer,
                                  ByVal _ExtAntiPasBack As Boolean)
        objMainRun.UserToSoyal(DeviceNo, idexCard, TagId, PersonNo, Name, Sex, Role, Department, _Class, PinCode, AccMode, _Patrol,
                                ExpireCheck, AntiPasBack, _PswChang, WGZone, idexZone, Group1, Group2, ValidDate, _LevelRang, LeveWGZone,
                                _ExtAntiPasBack)
    End Sub
    <WebMethod()>
    Public Sub AliasToSoyal(ByVal DeviceNo As String, ByVal idexCard As Integer, ByVal Name As String)
        objMainRun.AliasToSoyal(DeviceNo, idexCard, Name)
    End Sub
    <WebMethod()>
    Public Sub TimeZoneToSoyal(ByVal DeviceNo As String)
        objMainRun.TimeZoneToSoyal(DeviceNo)
    End Sub
    <WebMethod()>
    Public Sub HolidayToSoyal(ByVal DeviceNo As String)
        objMainRun.HolidayToSoyal(DeviceNo)
    End Sub
    <WebMethod()>
    Public Sub EraseSoyalData(ByVal DeviceNo As String)
        objMainRun.EraseSoyalData(DeviceNo)
    End Sub
    <WebMethod()>
    Public Sub EraseSoyalAlias(ByVal DeviceNo As String)
        objMainRun.EraseSoyalAlias(DeviceNo)
    End Sub
    <WebMethod()>
    Public Sub EraseSoyalTimeZone(ByVal DeviceNo As String)
        objMainRun.EraseSoyalTimeZone(DeviceNo)
    End Sub
    '整批Soyal資料庫中的UserCard資料表往卡機送,具有多個DeviceNo同時輸入,DeviceS中間用","區開可多設備輸入
    Public Sub SetSoyalCard(ByVal DeviceS() As String) '以卡號為順序，有利於同一執行緒可分散插入字串
        objMainRun.SetSoyalCard(DeviceS)

    End Sub

    '整批Soyal資料庫中的TimeZone資料表往卡機送,具有多個DeviceNo同時輸入,DeviceS中間用","區開可多設備輸入
    Public Sub SetSoyalTimeZone(ByVal DeviceS() As String)
        objMainRun.SetSoyalTimeZone(DeviceS)

    End Sub
    ''
    Public Sub DeleteSoyalCard(ByVal DeviceS() As String, ByVal UID() As String) '整批刪卡號
        objMainRun.DeleteSoyalCard(DeviceS, UID)
    End Sub


    '整批Soyal資料庫中的UserCard資料表往卡機送,具有多個DeviceNo同時輸入
    Public Sub UserSqlToSoyal(ByVal DeviceS() As String) '以卡號為順序，有利於同一執行緒可分散插入字串
        objMainRun.UserSqlToSoyal(DeviceS)
    End Sub
  
    '卡片及別名增修1筆到各卡機(同時改變SQL及卡機,但是一筆筆傳),適用於單頁面各別增修卡片資料時
    Public Sub UserToSqlSoyal(ByVal DeviceS() As String, ByVal TagId As String, ByVal PersonNo As String, ByVal Name As String, _
                              ByVal Sex As String, ByVal Role As String, ByVal Department As String, ByVal _Class As String, ByVal PinCode As String, _
                               ByVal AccMode As String, ByVal Patrol As Boolean, ByVal ExpireCheck As Boolean, ByVal AntiPasBack As Boolean, _
                               ByVal WGZone As Boolean, ByVal idexZone As Integer, ByVal PswChang As Boolean, ByVal Group1 As String, ByVal Group2 As String, _
                              ByVal ValidDate As String, ByVal LevelRang As Integer, ByVal LeveWGZone As Integer, ByVal ExtAntiPasBack As Boolean) '將資料載入卡機 83H

        objMainRun.UserToSqlSoyal(DeviceS, TagId, PersonNo, Name, Sex, Role, Department, _Class, PinCode, AccMode, Patrol, _
                        ExpireCheck, AntiPasBack, PswChang, WGZone, idexZone, Group1, Group2, ValidDate, LevelRang, LeveWGZone, _
                        ExtAntiPasBack)

    End Sub


#End Region


    <WebMethod()> _
    Public Function TimeServer() As String
        Dim h = Right("00" & Now.Hour, 2)
        Dim m = Right("00" & Now.Minute, 2)
        Dim s = Right("00" & Now.Second, 2)

        Return h & ":" & m & ":" & s
    End Function

#Region "單向送控制資料"
    <WebMethod()> _
    Public Sub FunNoValueUpdata(ByVal FunNo As String, ByVal Value As String) '物件名稱及值就可以變更資料到PLC

        objMainRun.FunNoValueUpdata(FunNo, Value)

    End Sub
    '物件名稱及值,且用連續數目變更資料到PLC
    <WebMethod()> _
    Public Sub FunNoValueUpdataMultiple(ByVal FunNo As String, ByVal Value As String, ByVal Quantity As Integer)

        objMainRun.FunNoValueUpdataMultiple(FunNo, Value, Quantity)

    End Sub
    <WebMethod()> _
    Public Sub Swtich_ON(ByVal FunNo As String) '送一個開的動作

        objMainRun.FunNoValueUpdata(FunNo, "1")

    End Sub
    <WebMethod()> _
    Public Sub Swtich_OFF(ByVal FunNo As String) '送一個關的動作

        objMainRun.FunNoValueUpdata(FunNo, "0")

    End Sub
    <WebMethod()> _
    Public Sub No_Button(ByVal FunNo As String)
        objMainRun.No_Button(FunNo)
    End Sub
    <WebMethod()> _
    Public Sub Event_Ctrl_obj(ByVal EvenName As String, ByVal EdgeUp As Boolean)

        objMainRun.Event_Ctrl_obj(EvenName, EdgeUp)
    End Sub

    <WebMethod()> _
    Public Sub NC_Button(ByVal FunNo As String)
        objMainRun.NC_Button(FunNo)
    End Sub
#End Region

#Region "雙向送控制資料"
    '傳送查詢命令
    <WebMethod()> _
    Public Sub FunNoInquiryToInCmd(ByVal FunNo As String, ByVal Items As Integer) '傳命令碼給PLC做立即之查詢
        objMainRun.FunNoInquiryToInCmd(FunNo, Items)

    End Sub

    <WebMethod()> _
    Public Function FunNoValueGetPLC(ByVal FunNo As String) As String '取得物件目前的值(傳命令碼給PLC之後所得的值)
        Return objMainRun.FunNoValueGetPLC(FunNo)

    End Function

    <WebMethod()> _
    Public Function Turn_Button(ByVal FunNo As String) As Boolean '一次開再一次關
        Try
            Return objMainRun.Turn_Button(FunNo)
        Catch ex As Exception

        End Try


    End Function
#End Region

#Region "查詢即時內容資料"
    <WebMethod()> _
    Public Function objChgArray() As String() '把objDisply內容轉成陣列方式
        Dim D = objMainRun.objChgArray
        Return D
    End Function
    <WebMethod()> _
    Public Function objArray() As String() '把objDisply內容轉成陣列方式
        If IsNothing(objMainRun.objArray) Then
            Return Nothing
        Else
            Return objMainRun.objArray()
        End If

    End Function

    <WebMethod()> _
    Public Function UserSeventScreen(ByVal UserName As String) As String() '取得使用者目前螢幕狀態
        '    UserScreen.Add(UserName, {PathFile, Detailed, 日期時間, Count})
        Return objMainRun.UserSeventScreen(UserName)

    End Function
    <WebMethod()> _
    Public Function objEvent(ByVal No As Integer) As String() '查得EventS事件目前該筆的狀態值，以便了解事件觸發後進度

        Dim ev = objMainRun.objEvent(No)

        Try
            Return {ev(0), ev(1), ev(2), ev(3)} 'Crossed狀態 , EventStatus狀態, etTick值,UpTimer

        Catch ex As Exception
            Return {"False", "False", "00", "0"}
        End Try
    End Function

    <WebMethod()> _
    Public Function FunNoValueGet(ByVal FunNo As String) As String '取得物件目前的值(並無傳命令碼給PLC)
        Try
            Return objMainRun.FunNoValueGet(FunNo)
        Catch ex As Exception
            Return "#"
        End Try


    End Function

    <WebMethod()> _
    Public Function objDisply(ByVal FunNo As String) As String() '{d.功能, d.值, d.單位, d.設備, d.地點, Now, SysType, 上一SQL變動值, 上一RecRegist變動值})
        Try
            '取得物件之陣列表示的各狀態
            'd(0)取得功能,d(1)取得值,d(2)取得單位,d(3)取得設備名稱,d(4)取得該測點的放置地點......
            Dim d0 = objMainRun.objDisply(FunNo)
            Return d0
        Catch ex As Exception  '表示Windows Sservice SRCMS停止接收中
            Dim d1 = {"???", "#", "??", "", "", "", "", "", ""}
            Return d1
        End Try


    End Function






    'Private WithEvents Timer1 As New System.Timers.Timer
    'Private ObjD_m()() As String
    'Private Sub Timer1_Elapsed(sender As Object, e As System.Timers.ElapsedEventArgs) Handles Timer1.Elapsed
    '    ObjD_m = objMainRun.objDisply
    'End Sub

    <WebMethod()> _
    Public Function objDisplyChg(ByVal FunNo As String) As String() '{d.功能, d.值, d.單位, d.設備, d.地點, d.回應收碼間距}
        Try
            '取得物件之陣列表示的各狀態
            'd(0)取得功能,d(1)取得值,d(2)取得單位,d(3)取得設備名稱,d(4)取得該測點的放置地點......
            Dim d0 = objMainRun.objDisplyChg(FunNo)
            Return d0
        Catch ex As Exception
            Dim d = {"^", "^", "^", "", "", ""}
            Return d
        End Try


    End Function

    <WebMethod()> _
    Public Function DispRegist(ByVal FunNo As String) As String() '{真值, 單位, 日期時間}
        Try
            '取得物件之陣列表示的各狀態
            'd(0)取得功能,d(1)取得值,d(2)取得單位,d(3)取得設備名稱,d(4)取得該測點的放置地點......
            Dim d0 = objMainRun.DispRegist(FunNo)
            Return d0
        Catch ex As Exception
            Dim d = {"^", "^", "^"}
            Return d
        End Try


    End Function
    <WebMethod()> _
    Public Function DispMeterPower(ByVal FunNo As String) As String() '{本年累度數, 本月累度數, 本週累度數, 本日累度數, 去年總度數, 去年同月總度數, 上月總度數, 昨日總度數}
        Try

            Dim d0 = objMainRun.DispMeterPower(FunNo)
            Return d0
        Catch ex As Exception

        End Try


    End Function

    <WebMethod()> _
    Public Function Disply() As DataTable '以DataTable方式顯示物件狀態
        Try
            Return objMainRun.Disply
        Catch ex As Exception

        End Try


    End Function
    <WebMethod()> _
    Public Function DisplyChg() As DataTable '以DataTable方式顯示物件狀態
        Try
            Return objMainRun.DisplyChg
        Catch ex As Exception
            Dim f = ""
        End Try


    End Function

    <WebMethod()> _
    Public Function Dy() As DataSet '以DataTable方式顯示物件狀態
        Dim ds = New DataSet()
        Dim dt = New DataTable()
        dt.Columns.Add("Id", GetType(Integer))
        dt.Columns.Add("Name", GetType(String))
        For i = 1 To 55
            Dim row = dt.NewRow()
            row(0) = i
            row(1) = "N" + CStr(i)
            dt.Rows.Add(row)

            ' dt.Rows.Add(i, "N" + i)
        Next
        ds.Tables.Add(dt)

        Return ds

    End Function
    <WebMethod()> _
    Public Function DisplyEvent() As DataTable '顯示EventS
        Return objMainRun.DisplyEvent

    End Function
    <WebMethod()> _
    Public Function EventHodeScreen() As DataTable '事件佇留螢幕
        Return objMainRun.EventHodeScreen()
    End Function

    <WebMethod()> _
    Public Function 通信狀態() As DataTable
        Return objMainRun.通信狀態()
    End Function
#End Region

#Region "載入更新"
    <WebMethod()> _
    Public Sub Update() '更新各物件...等等 處理暫存區
        objMainRun.Update()
    End Sub
    <WebMethod()> _
    Public Sub UpdateSerialPort() '所有Device由SQL資料表重新載入，並且關閉Modubus 重新 Modbus執行建立()
        objMainRun.UpdateSerialPort()
    End Sub

    <WebMethod()> _
    Public Sub UpdateDevice() '所有Device由SQL資料表重新載入，並且關閉Modubus 重新 Modbus執行建立()
        objMainRun.UpdateDevice()
    End Sub
    <WebMethod()> _
    Public Sub UpdateFunNo() '所有Funtion由SQL資料表重新載入
        objMainRun.UpdateFunNo()
    End Sub
    <WebMethod()> _
    Public Sub UpdateTimerTemporary()  '所有UpdateTimerTemporary() 由SQL資料表重新載入
        objMainRun.UpdateTimerTemporary()
    End Sub
    <WebMethod()> _
    Public Sub UpdateTimerEvent() '所有TimerEvent由SQL資料表重新載入
        objMainRun.UpdateTimerEvent()
    End Sub
    <WebMethod()> _
    Public Sub UpdateEvevts() '所有Evevts由SQL資料表重新載入
        objMainRun.UpdateEvevts()
    End Sub
    <WebMethod()> _
    Public Sub UpdateEvevtPLC_Music_() '所有EventToPLC , EventMisic ,EventEmail,EventMsgTxt,EventIR由SQL資料表重新載入
        objMainRun.UpdateEvevtPLC_Music_()
    End Sub
    <WebMethod()> _
    Public Sub UpdateDoorCard() '更新門卡
        objMainRun.UpdateDoorCard()
    End Sub

    <WebMethod()> _
    Public Sub UpdateScanCmd() '所有ScanCmd由SQL資料表重新載入到每一個Modus掃瞄區
        objMainRun.UpdateScanCmd()
    End Sub
    <WebMethod()> _
    Public Sub UpdataScanCmdPort(ByVal CnnPort As String) '更新掃瞄命令
        objMainRun.UpdateScanCmdPort(CnnPort)
    End Sub


    <WebMethod()> _
    Public Sub UdateTxtMsgCh()
        objMainRun.UdateTxtMsgCh()
    End Sub

#End Region



#Region "行程排定"
    '插入新的排程 , 本功能會把過期的資料庫清空
    <WebMethod()> _
    Public Sub InsertScheduleTemporary(ByVal year As Integer, ByVal month As Integer, _
                                          ByVal day As Integer, ByVal StartHour As Integer, _
                                   ByVal StartMinute As Integer, ByVal StartSecond As Integer, _
                                   ByVal EndHour As Integer, ByVal EndMinute As Integer, _
                                   ByVal EndSecond As Integer, ByVal CtrlObj As String, _
                                   ByVal StartValue As String, ByVal EndValue As String, _
                                   ByVal Keep As Integer) '臨時排程


        objMainRun.InsertScheduleTemporary(year, month, day, StartHour, StartMinute, StartSecond, EndHour, EndMinute, EndSecond, CtrlObj, StartValue, EndValue, Keep)

    End Sub

    '以物件查尋排程
    <WebMethod()> _
    Public Function SearchScheduleTemporary(ByVal CtrlObj As String) As Object()
        Return objMainRun.SearchScheduleTemporary(CtrlObj)
    End Function
    '刪除排程
    <WebMethod()> _
    Public Sub DeleteScheduleTemporary(ByVal SnNO As String)
        objMainRun.DeleteScheduleTemporary(SnNO)
    End Sub

#End Region

#Region "IR紅外線"
    <WebMethod()> _
    Public Sub irTransmit(ByVal irNo As String) '紅外線物件名稱發射一個命令
        objMainRun.irTransmit(irNo)

    End Sub
    <WebMethod()> _
    Public Sub irStudy(ByVal irNo As String) '紅外線學習
        objMainRun.irStudy(irNo)
    End Sub
    <WebMethod()> _
    Public Sub irSetAddress(ByVal DeviceNo As String, ByVal NewSite As Integer) '紅外線更改地址
        objMainRun.irSetAddress(DeviceNo, NewSite)
    End Sub
    <WebMethod()> _
    Public Sub irSetAllAddress(ByVal ComPort As String, ByVal NewSite As Integer)  '紅外線同埠全改相同地址
        objMainRun.irSetAllAddress(ComPort, NewSite)
    End Sub
    <WebMethod()> _
    Public Sub irSendCmd(ByVal ComPort As String, ByVal cmd As String) '向通信埠具IR,以字串方式下命令

        objMainRun.irSendCmd(ComPort, cmd)

    End Sub
    <WebMethod()> _
    Public Function irResponeStatus() As String '取得IR最後回應內容
        Return objMainRun.irResponeStatus
    End Function

#End Region

#Region "投影機及會議系統"

    '投影機插入
    <WebMethod()> _
    Public Sub PorjtToInCmd(ByVal FunName As String, ByVal Value As String)
        objMainRun.PorjtToInCmd(FunName, Value)
    End Sub
    '會議系統插入
    <WebMethod()> _
    Public Sub ConfToInCmd(ByVal FunName As String, ByVal Value As String)
        objMainRun.ConfToInCmd(FunName, Value)
    End Sub
#End Region



#Region "音樂狀態"
    <WebMethod()> _
    Public Function MusicGet() As List(Of Object()) '取得音樂

        '    Return objMainRun.MusicGet
        Return Nothing
    End Function
    <WebMethod()> _
    Public Sub MusicClear() '清除所有音樂
        objMainRun.MusicClear()

    End Sub
    <WebMethod()> _
    Public Sub MusicExtUse(ByVal Use As Boolean) '外部MusicDetec專案是否啟用通知本服務應用程式
        objMainRun.MusicExtUse(Use)
    End Sub
    <WebMethod()> _
    Public Sub MusicRemotAt(ByVal index As Integer) '清除第幾筆
        objMainRun.MusicRemotAt(index)
    End Sub
    <WebMethod()> _
    Public Sub SetMusicPopupTime(ByVal Lenght As Integer) '當音樂2筆以上時，設定跳轉音樂時間
        objMainRun.SetMusicPopupTime(Lenght)
    End Sub
#End Region


#Region "Text文字狀態"
    <WebMethod()> _
    Public Function TextGet() As String '取得Text
        Return objMainRun.TextGet
    End Function
#End Region

#Region "IP攝影機"
    '攝影機狀態改變插入
    <WebMethod()> _
    Public Sub CameraToInCmd(ByVal FunNo As String, ByVal Value As String)
        objMainRun.CameraToInCmd(FunNo, Value)
    End Sub

    '攝影機狀態值
    <WebMethod()> _
    Public Function CameraStatusList() As String()()
      
        '    Dim Cmlist = objMainRun.CameraStatusList.ToArray()

        'Return Cmlist
        Return Nothing
    End Function
#End Region


#Region "網頁格式"
    '新增及修改網頁格式物件
    <WebMethod()> _
    Public Sub addUpPageStyle(ByVal PageName As String, ByVal ElementId As String, ByVal FunNo As String, ByVal offsetTop As Integer, _
                                 ByVal offsetLeft As Integer, ByVal offsetWidth As Integer, ByVal offsetHeight As Integer _
                                 , ByVal BgImg As String)

        objMainRun.addUpPageStyle(PageName, ElementId, FunNo, offsetTop, offsetLeft, offsetWidth, offsetHeight, BgImg)

    End Sub


    '查詢網頁格式
    <WebMethod()> _
    Public Function QuPageStyle(ByVal PageName As String, ByVal ElementId As String) As String()
        Return objMainRun.QuPageStyle(PageName, ElementId)
    End Function
    '刪除網頁格式
    <WebMethod()> _
    Public Sub DelPageStyle(ByVal PageName As String, ByVal ElementId As String)
        objMainRun.DelPageStyle(PageName, ElementId)
    End Sub
#End Region

#Region "註冊權限查詢"
    <WebMethod()> _
    Public Function AuthQuantity() As String
        Return objMainRun.AuthQuantity
    End Function
    <WebMethod()> _
    Public Function AddFunQut() As String
        Return objMainRun.AddFunQut
    End Function
    <WebMethod()> _
    Public Function AuthorityNo() As String
        Return objMainRun.AuthorityNo
    End Function
    <WebMethod()> _
    Public Function AuthStatus() As String
        Return objMainRun.AuthStatus
    End Function
#End Region

#Region "註冊權限修改"
    '**'向遠端端註冊機重新啟動,以改變新的權密
    <WebMethod()> _
    Public Sub Regist_Act()
        objMainRun.Regist_Act()
    End Sub

    ''**向遠端端註冊機做新產品註冊
    <WebMethod()> _
    Public Sub SetRegit(ByVal RegistNo As String, ByVal AuthNo As String, ByVal UserName As String _
                                 , ByVal UserEmail As String, ByVal UserMovTel As String _
                                 , ByVal UserAddress As String)

        objMainRun.SetRegit(RegistNo, AuthNo, UserName, UserEmail, UserMovTel, UserAddress)
    End Sub
    ''**向遠端端註冊機取得產品的權密及規格
    <WebMethod()> _
    Public Function GetRegistNoMsg(ByVal RegistNo As String) As String() '取得產品註冊內容規格
        Return objMainRun.GetRegistNoMsg(RegistNo)
    End Function

    <WebMethod()> _
    Public Sub AllRegistOpen(ByVal Password As String)
        objMainRun.AllRegistOpen(Password)
    End Sub

    '**向使用者端取得資料
    <WebMethod()> _
    Public Function 註冊碼取得() As String()
        Dim xEle As XElement = XElement.Load("C:\SRCMS\SRCMS\UserKey.xml")

        'LINQ查詢
        Dim UserKey = From ex In xEle.<KeyRegist>
        Dim x0, x1 As String
        For Each x As XElement In UserKey
            x0 = x.<RegistNo>.Value
            x1 = x.<AuthNo>.Value

        Next
        Return {x0, x1}
    End Function
    '**向使用者端更改權限密碼
    <WebMethod()> _
    Public Sub AuthNoUpdate(ByVal AuthNo As String)
        Dim xEle As XElement = XElement.Load("C:\SRCMS\SRCMS\UserKey.xml")
        Dim UserKey = From ex In xEle.<KeyRegist>
        For Each x As XElement In UserKey
            '    x.Attribute("TP_NAME").SetValue("Change")
            x.<AuthNo>.Value = AuthNo
        Next
        xEle.Save("C:\SRCMS\SRCMS\UserKey.xml")
    End Sub
    '**向使用者端更改註冊及權限密碼
    <WebMethod()> _
    Public Sub RegistNo_AuthNoUpdate(ByVal RegistNo As String, ByVal AuthNo As String)
        Dim xEle As XElement = XElement.Load("C:\SRCMS\SRCMS\UserKey.xml")
        Dim UserKey = From ex In xEle.<KeyRegist>
        For Each x As XElement In UserKey
            '    x.Attribute("TP_NAME").SetValue("Change")
            x.<RegistNo>.Value = RegistNo
            x.<AuthNo>.Value = AuthNo
        Next
        xEle.Save("C:\SRCMS\SRCMS\UserKey.xml")
    End Sub

#End Region

#Region "具記錄來源者是誰來觸動,具插入功能兼記錄到SQL PLC_Record EventRec" '用於Web登入者最有效
    <WebMethod()> _
    Public Sub FunNoValueUpdataRecord(ByVal SourceUser As String, ByVal SourceValue As String, _
                                           ByVal FunNo As String, ByVal Value As String)
        objMainRun.FunNoValueUpdataRecord(SourceUser, SourceValue, FunNo, Value)
    End Sub

    <WebMethod()> _
    Public Shared Sub Turn_Button_Record(ByVal SourceUser As String, ByVal SourceValue As String, ByVal FunNo As String) '一次開再一次關
        objMainRun.Turn_Button_Record(SourceUser, SourceValue, FunNo)
    End Sub
    <WebMethod()> _
    Public Sub Swtich_ON_Record(ByVal SourceUser As String, ByVal SourceValue As String, ByVal FunNo As String) '送一個開的動作
        objMainRun.Swtich_ON_Record(SourceUser, SourceValue, FunNo)
    End Sub
    <WebMethod()> _
    Public Sub Swtich_OFF_Record(ByVal SourceUser As String, ByVal SourceValue As String, ByVal FunNo As String) '送一個開的動作
        objMainRun.Swtich_OFF_Record(SourceUser, SourceValue, FunNo)
    End Sub
    <WebMethod()> _
    Public Sub No_Button_Record(ByVal SourceUser As String, ByVal SourceValue As String, ByVal FunNo As String)  '送給物件一個Plus,先1後0
        objMainRun.No_Button_Record(SourceUser, SourceValue, FunNo)

    End Sub
    <WebMethod()> _
    Public Sub NC_Button_Record(ByVal SourceUser As String, ByVal SourceValue As String, ByVal FunNo As String) '送給物件一個Plus,先0後1
        objMainRun.NC_Button_Record(SourceUser, SourceValue, FunNo)
    End Sub
#End Region
#Region "取得歷史記錄"
    <WebMethod()> _
    Public Function GetFunRec(ByVal startDate As String, ByVal endDate As String, ByVal FunNo As String) As Object()

        Return objMainRun.GetFunRec(startDate, endDate, FunNo)
    End Function
    <WebMethod()> _
    Public Function GetFunRecstartEnd(ByVal startDate As String, ByVal endDate As String, ByVal FunNo As String) As Object()

        Return objMainRun.GetFunRecstartEnd(startDate, endDate, FunNo)
    End Function
    <WebMethod()> _
    Public Function GetEventRec(ByVal startDate As String, ByVal endDate As String, ByVal SourceFunNo As String, ByVal EventNo As String, ByVal TargetFunNo As String) As Object()

        Return objMainRun.GetEventRec(startDate, endDate, SourceFunNo, EventNo, TargetFunNo)
    End Function
#End Region


End Class
