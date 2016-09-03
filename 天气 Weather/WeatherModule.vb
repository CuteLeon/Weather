Module WeatherModule
    Public Structure BasicWeatherModel
        Dim City As String
        Dim UpdateTime As String
        Dim Temperature As String
        Dim WindPower As String
        Dim Humidity As String
        Dim WindDirection As String
        Dim SunRise As String
        Dim SunSet As String
    End Structure

    Public Structure IndexModel '指数结构模型
        Dim IndexName As String
        Dim IndexValue As String
        Dim IndexDetail As String
    End Structure
    Public Structure EnvironmentModel '环境参数结构模型
        Dim AQI As Integer
        Dim PM25 As Integer
        Dim Suggest As String
        Dim Quality As String
        Dim MajorPollutants As String
        Dim O3 As Integer
        Dim CO As Integer
        Dim PM10 As Integer
        Dim SO2 As Integer
        Dim NO2 As Integer
        Dim EnvirTime As String
    End Structure
    Public Structure WeatherModel '天气参数结构模型
        Dim WeatherDate As String
        Dim HighTemp As Integer
        Dim LowTemp As Integer
        Dim DayType As String
        Dim DayWindDirection As String
        Dim DayWindPower As String
        Dim NightType As String
        Dim NightWindDirection As String
        Dim NightWindPower As String
    End Structure


    Public Sub ReadBasicWeather(ByVal WebPage As String, ByRef BasicWeatherBody As BasicWeatherModel) '读取即时基本天气
        BasicWeatherBody.City = getValue(WebPage, "city")
        BasicWeatherBody.UpdateTime = getValue(WebPage, "updatetime")
        BasicWeatherBody.Temperature = getValue(WebPage, "wendu")
        BasicWeatherBody.WindPower = getValue(WebPage, "fengli")
        BasicWeatherBody.Humidity = getValue(WebPage, "shidu")
        BasicWeatherBody.WindDirection = getValue(WebPage, "fengxiang")
        BasicWeatherBody.SunRise = getValue(WebPage, "sunrise_1")
        BasicWeatherBody.SunSet = getValue(WebPage, "sunset_1")
    End Sub

    Public Sub ReadEnvironment(ByVal WebPage As String, ByRef EnvironmentBody As EnvironmentModel) '读取环境参数
        Dim EnvironmentBlock As String '从网页截取的环境部分代码块
        EnvironmentBlock = getValue(WebPage, "environment")
        EnvironmentBody.AQI = Val(getValue(EnvironmentBlock, "aqi"))
        EnvironmentBody.PM25 = Val(getValue(EnvironmentBlock, "pm25"))
        EnvironmentBody.Suggest = getValue(EnvironmentBlock, "suggest")
        EnvironmentBody.Quality = getValue(EnvironmentBlock, "quality")
        EnvironmentBody.MajorPollutants = getValue(EnvironmentBlock, "MajorPollutants")
        EnvironmentBody.O3 = Val(getValue(EnvironmentBlock, "o3"))
        EnvironmentBody.CO = Val(getValue(EnvironmentBlock, "co"))
        EnvironmentBody.PM10 = Val(getValue(EnvironmentBlock, "pm10"))
        EnvironmentBody.SO2 = Val(getValue(EnvironmentBlock, "so2"))
        EnvironmentBody.NO2 = Val(getValue(EnvironmentBlock, "no2"))
        EnvironmentBody.EnvirTime = getValue(EnvironmentBlock, "time")
    End Sub
    Public Sub ReadIndex(ByVal WebPage As String, ByRef IndexBody() As IndexModel) '读取各项指数
        Dim IndexsBlock As String '从网页截取的指数部分代码块
        Dim IndexBlock(10) As String
        IndexsBlock = getValue(WebPage, "zhishus")
        For Index As Int16 = LBound(IndexBody) To UBound(IndexBody)
            IndexBlock(Index) = getValue(IndexsBlock, "zhishu")
            IndexBody(Index).IndexName = getValue(IndexBlock(Index), "name")
            IndexBody(Index).IndexValue = getValue(IndexBlock(Index), "value")
            IndexBody(Index).IndexDetail = getValue(IndexBlock(Index), "detail")
            IndexsBlock = Strings.Right(IndexsBlock, IndexsBlock.Length - IndexBlock(Index).Length - 17) '17是Len("<zhishu></zhishu>")
        Next
    End Sub
    Public Sub ReadYesterday(ByVal WebPage As String, ByRef WeatherBody As WeatherModel) '读取昨天天气
        Dim YesterdayBlock As String '从网页截取的昨天天气部分代码块
        Dim DayBlock As String, NightBlock As String, Temperature As String
        YesterdayBlock = getValue(WebPage, "yesterday")
        DayBlock = getValue(YesterdayBlock, "day_1")
        NightBlock = getValue(YesterdayBlock, "night_1")

        WeatherBody.WeatherDate = getValue(YesterdayBlock, "date_1")
        Temperature = getValue(YesterdayBlock, "high_1")
        WeatherBody.HighTemp = Val(Strings.Mid(Temperature, 4, Temperature.Length - 4))
        WeatherForm.HightTemps(0) = WeatherBody.HighTemp
        Temperature = getValue(YesterdayBlock, "low_1")
        WeatherBody.LowTemp = Val(Strings.Mid(Temperature, 4, Temperature.Length - 4))
        WeatherForm.LowTemps(0) = WeatherBody.LowTemp
        WeatherBody.DayType = getValue(DayBlock, "type_1")
        WeatherBody.DayWindDirection = getValue(DayBlock, "fx_1")
        WeatherBody.DayWindPower = getValue(DayBlock, "fl_1")
        WeatherBody.NightType = getValue(NightBlock, "type_1")
        WeatherBody.NightWindDirection = getValue(NightBlock, "fx_1")
        WeatherBody.NightWindPower = getValue(NightBlock, "fl_1")
    End Sub
    Public Sub ReadWeather(ByVal WebPage As String, ByRef WeatherBody() As WeatherModel) '读取未来天气
        Dim ForecastBlock As String '从网页截取的未来天气部分代码块
        Dim WeatherBlock(4) As String
        ForecastBlock = getValue(WebPage, "forecast")
        Dim DayBlock As String, NightBlock As String, Temperature As String
        For Index As Int16 = LBound(WeatherBody) + 1 To UBound(WeatherBody)
            MsgBox(Index)
            WeatherBlock(Index - 1) = getValue(ForecastBlock, "weather")
            DayBlock = getValue(WeatherBlock(Index - 1), "day")
            NightBlock = getValue(WeatherBlock(Index - 1), "night")

            WeatherBody(Index).WeatherDate = getValue(WeatherBlock(Index - 1), "date")
            Temperature = getValue(WeatherBlock(Index - 1), "high")
            WeatherBody(Index).HighTemp = Val(Strings.Mid(Temperature, 4, Temperature.Length - 4))
            WeatherForm.HightTemps(Index) = WeatherBody(Index).HighTemp
            Temperature = getValue(WeatherBlock(Index - 1), "low")
            WeatherBody(Index).LowTemp = Val(Strings.Mid(Temperature, 4, Temperature.Length - 4))
            WeatherForm.LowTemps(Index) = WeatherBody(Index).LowTemp - 10
            WeatherBody(Index).DayType = getValue(DayBlock, "type")
            WeatherBody(Index).DayWindDirection = getValue(DayBlock, "fengxiang")
            WeatherBody(Index).DayWindPower = getValue(DayBlock, "fengli")
            WeatherBody(Index).NightType = getValue(NightBlock, "type")
            WeatherBody(Index).NightWindDirection = getValue(NightBlock, "fengxiang")
            WeatherBody(Index).NightWindPower = getValue(NightBlock, "fengli")
            ForecastBlock = Strings.Right(ForecastBlock, ForecastBlock.Length - WeatherBlock(Index - 1).Length - 19) '19是Len("<weather></weather>")
        Next
    End Sub

    Public Sub ShowBasicWeather(ByVal BasicWeatherBody As BasicWeatherModel) '输出基本天气信息
        With BasicWeatherBody
            Debug.Print("城市：" & .City & vbCrLf & "更新：" & .UpdateTime & vbCrLf & "温度：" & .Temperature & vbCrLf & "风力：" & .WindPower & vbCrLf & "湿度：" & .Humidity & vbCrLf & "风向：" & .WindDirection & vbCrLf & "日出：" & .SunRise & vbCrLf & "日落：" & .SunSet)
        End With
    End Sub
    Public Sub ShowIndex(ByVal IndexBody As IndexModel) '输出指数
        Debug.Print(IndexBody.IndexName & ":" & IndexBody.IndexValue & vbCrLf & IndexBody.IndexDetail, 0, IndexBody.IndexName)
    End Sub
    Public Sub ShowWeather(ByVal WeatherBody As WeatherModel) '输出天气
        With WeatherBody
            Debug.Print("日期：" & .WeatherDate & vbCrLf & "高温：" & .HighTemp & vbCrLf & "低温：" & .LowTemp & vbCrLf &
                        "白天：" & .DayType & vbCrLf & "风向：" & .DayWindDirection & vbCrLf & "风力：" & .DayWindPower & vbCrLf &
                        "夜间：" & .NightType & vbCrLf & "风向：" & .NightWindDirection & vbCrLf & "风力：" & .NightWindPower)
        End With
    End Sub
    Public Sub ShowEnvironment(ByVal EnvironmentBody As EnvironmentModel) '输出环境
        With EnvironmentBody
            Debug.Print("AQI：" & .AQI & vbCrLf & "PM25：" & .PM25 & vbCrLf & "建议：" & .Suggest & vbCrLf & "质量：" & .Quality & vbCrLf &
                   "主要污染物：" & .MajorPollutants & vbCrLf & "O3：" & .O3 & vbCrLf & "CO：" & .CO & vbCrLf & "PM10：" & .PM10 & vbCrLf &
                   "SO2：" & .SO2 & vbCrLf & "NO2：" & .NO2 & vbCrLf & "更新时间：" & .EnvirTime)
        End With
    End Sub

    Public Function getWeatherPage(ByVal UrlLink As String) As String '读取网页源代码
        Try
            Dim XmlHTTP As Object
            Dim WebContent As Object
            Dim XMLPage As String
            XmlHTTP = CreateObject("Microsoft.XMLHttp")
            XmlHTTP.Open("POST", UrlLink, False)
            XmlHTTP.Send()
            WebContent = XmlHTTP.ResponseText
            XMLPage = WebContent.ToString
            XmlHTTP = Nothing
            Return XMLPage
        Catch ex As Exception
            Return ("Can not get the web page.")
        End Try
    End Function

    Public Function getValue(ByVal WebPage As String, ByVal LabelName As String) As String '从源代码截取标签的数据
        Dim StartP As Integer, EndP As Integer
        Try
            StartP = InStr(WebPage, "<" & LabelName & ">")
            EndP = InStr(StartP + 1, WebPage, "</" & LabelName & ">")
            Return (IIf((StartP > 0 And EndP > 0), Mid(WebPage, StartP + Len(LabelName) + 2, EndP - StartP - Len(LabelName) - 2), ""))
        Catch ex As Exception
            Return "Can not find this label named """ & LabelName & """"
        End Try
    End Function

End Module
