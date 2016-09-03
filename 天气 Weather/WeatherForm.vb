Imports System.Drawing.Drawing2D

Public Class WeatherForm
    'http://www.weather.com.cn//weather1d//101180101.shtml  内有逐小时预报

    Private Const CityID As String = "101180101" '郑州市的城市ID号，其他城市的ID号，自行百度
    Dim WebPage As String '网页所有源代码
    Dim EnvironmentBody As EnvironmentModel '环境参数模型的实例
    Dim WeatherBody(5) As WeatherModel '天气参数模型的实例
    Dim BasicWeatherBody As BasicWeatherModel '基本即时天气实例
    Dim IndexBody(10) As IndexModel '指数模型的实例
    Dim DataX() As Integer = {0, 50, 100, 150, 200, 250}
    Public HightTemps(5) As Integer '最高温数组
    Public LowTemps(5) As Integer '最低温数组
    Dim TempCurves As Bitmap = New Bitmap(250, 50)

    Private Sub WeatherForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WebPage = getWeatherPage("http://wthrcdn.etouch.cn/WeatherApi?citykey=" & CityID) '获取网页源代码
        Debug.Print(WebPage)
        If WebPage = "Can Then Not Get the web page." Then MsgBox(WebPage) : Exit Sub         '如果获取失败，退出
        '显示即时基本数据
        Debug.Print("获取数据成功！正在打印..." & vbCrLf)

        ReadBasicWeather(WebPage, BasicWeatherBody)
        Me.Text = BasicWeatherBody.City & "-天气"
        ShowBasicWeather(BasicWeatherBody)
        ReadYesterday(WebPage, WeatherBody(0))              '读取昨天天气
        ReadWeather(WebPage, WeatherBody)                    '读取未来天气
        ReadEnvironment(WebPage, EnvironmentBody)       '读取环境参数
        ReadIndex(WebPage, IndexBody)                             '读取各项指数

        Debug.Print(vbCrLf & "——输出近期天气：——")
        For Index As Integer = 0 To 5
            ShowWeather(WeatherBody(Index))                     '输出天气
            Debug.Print(vbNullString)
        Next

        Debug.Print("——输出环境参数：——")
        ShowEnvironment(EnvironmentBody)                      '输出环境
        Debug.Print(vbCrLf & "——输出各项指数：——")
        For Index As Integer = 0 To 10
            ShowIndex(IndexBody(Index))
            Debug.Print(vbNullString)
        Next
    End Sub
End Class