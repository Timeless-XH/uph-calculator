Attribute VB_Name = "UPHCalculator"
'==================================================
' UPH 流水线计算器 - Excel VBA 版本
' 功能：模拟多工站流水线生产，计算UPH和瓶颈工站
'==================================================

Option Explicit

' 工站数据结构
Private Type Station
    Name As String
    ProcessTime As Double
    BufferCapacity As Integer
End Type

' 产品时间线事件
Private Type TimeEvent
    StationIndex As Integer
    StartTime As Double
    EndTime As Double
End Type

' 配置数据结构
Private Type Config
    Name As String
    ProductCount As Integer
    Stations() As Station
    StationCount As Integer
End Type

' 全局变量
Private CurrentConfig As Config
Private Configs As Collection
Private CurrentConfigName As String

'==================================================
' 主入口：创建UPH计算器界面
'==================================================
Public Sub CreateUPHCalculator()
    Dim ws As Worksheet
    
    ' 删除已存在的工作表
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("UPH计算器").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' 创建新工作表
    Set ws = Sheets.Add
    ws.Name = "UPH计算器"
    
    ' 初始化配置
    InitializeConfig
    
    ' 绘制界面
    DrawInterface ws
    
    ' 填充配置下拉框
    FillConfigCombo ws
    
    ' 加载默认配置
    LoadConfigToSheet ws, "默认配置"
End Sub

'==================================================
' 初始化默认配置
'==================================================
Private Sub InitializeConfig()
    Set Configs = New Collection
    
    ' 创建默认配置
    Dim defaultCfg As Config
    defaultCfg.Name = "默认配置"
    defaultCfg.ProductCount = 13
    defaultCfg.StationCount = 3
    
    ReDim defaultCfg.Stations(1 To 3)
    defaultCfg.Stations(1).Name = "焊接"
    defaultCfg.Stations(1).ProcessTime = 6
    defaultCfg.Stations(1).BufferCapacity = 2
    
    defaultCfg.Stations(2).Name = "切割"
    defaultCfg.Stations(2).ProcessTime = 20
    defaultCfg.Stations(2).BufferCapacity = 2
    
    defaultCfg.Stations(3).Name = "清洗"
    defaultCfg.Stations(3).ProcessTime = 8
    defaultCfg.Stations(3).BufferCapacity = 2
    
    Configs.Add defaultCfg, "默认配置"
    CurrentConfigName = "默认配置"
    CurrentConfig = defaultCfg
End Sub

'==================================================
' 绘制界面
'==================================================
Private Sub DrawInterface(ws As Worksheet)
    With ws
        ' 设置列宽
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 12
        .Columns("H:I").ColumnWidth = 12
        
        ' 标题
        .Range("B2:F2").Merge
        With .Range("B2")
            .Value = "UPH 流水线计算器"
            .Font.Size = 18
            .Font.Bold = True
            .Font.Color = RGB(88, 166, 255)
        End With
        
        ' 配置管理区域
        .Range("B4").Value = "当前配置："
        .Range("B4").Font.Bold = True
        
        ' 配置下拉框位置 (C4)
        With .Range("C4:D4")
            .Merge
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(48, 54, 61)
        End With
        
        ' 按钮
        Dim btn As Button
        Set btn = .Buttons.Add(.Range("E4").Left, .Range("E4").Top, 50, 20)
        btn.Caption = "保存"
        btn.OnAction = "SaveCurrentConfig"
        
        Set btn = .Buttons.Add(.Range("F4").Left, .Range("F4").Top, 60, 20)
        btn.Caption = "另存为"
        btn.OnAction = "SaveAsNewConfig"
        
        Set btn = .Buttons.Add(.Range("G4").Left, .Range("G4").Top, 50, 20)
        btn.Caption = "删除"
        btn.OnAction = "DeleteCurrentConfig"
        
        ' 基本参数区域
        .Range("B6:D6").Merge
        With .Range("B6")
            .Value = "基本参数"
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = RGB(139, 148, 158)
            .Interior.Color = RGB(22, 27, 34)
        End With
        
        .Range("B7").Value = "产品总数："
        .Range("C7").Value = 13
        With .Range("C7")
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .NumberFormat = "0"
        End With
        
        ' 工站配置区域
        .Range("B9:G9").Merge
        With .Range("B9")
            .Value = "工站配置"
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = RGB(139, 148, 158)
            .Interior.Color = RGB(22, 27, 34)
        End With
        
        ' 工站表头
        .Range("B10").Value = "序号"
        .Range("C10").Value = "工站名称"
        .Range("D10").Value = "处理时间(分)"
        .Range("E10").Value = "缓冲容量"
        With .Range("B10:E10")
            .Font.Bold = True
            .Interior.Color = RGB(13, 17, 23)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(48, 54, 61)
        End With
        
        ' 预留10行工站输入
        Dim i As Integer
        For i = 1 To 10
            .Range("B" & (10 + i)).Value = i
            .Range("B" & (10 + i)).HorizontalAlignment = xlCenter
            With .Range("C" & (10 + i) & ":E" & (10 + i))
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(48, 54, 61)
            End With
        Next i
        
        ' 添加/删除工站按钮
        Set btn = .Buttons.Add(.Range("B21:E21").Left, .Range("B21").Top, .Range("B21:E21").Width, 22)
        btn.Caption = "+ 添加工站"
        btn.OnAction = "AddStation"
        
        ' 计算按钮
        Set btn = .Buttons.Add(.Range("B23:C23").Left, .Range("B23").Top, .Range("B23:C23").Width, 28)
        btn.Caption = "计算"
        btn.OnAction = "RunSimulation"
        With btn
            .Font.Size = 12
            .Font.Bold = True
        End With
        
        Set btn = .Buttons.Add(.Range("D23:E23").Left, .Range("D23").Top, .Range("D23:E23").Width, 28)
        btn.Caption = "重置"
        btn.OnAction = "ResetResults"
        
        ' 结果区域
        .Range("G6:I6").Merge
        With .Range("G6")
            .Value = "计算结果"
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = RGB(139, 148, 158)
            .Interior.Color = RGB(22, 27, 34)
        End With
        
        ' 结果指标
        .Range("G7").Value = "实际UPH："
        .Range("H7").Value = "-"
        .Range("H7").Font.Size = 16
        .Range("H7").Font.Bold = True
        .Range("H7").Font.Color = RGB(88, 166, 255)
        
        .Range("G8").Value = "总耗时："
        .Range("H8").Value = "-"
        .Range("H8").Font.Size = 16
        .Range("H8").Font.Bold = True
        .Range("H8").Font.Color = RGB(63, 185, 80)
        
        .Range("G9").Value = "理论UPH："
        .Range("H9").Value = "-"
        .Range("H9").Font.Size = 14
        
        .Range("G10").Value = "瓶颈工站："
        .Range("H10").Value = "-"
        .Range("H10").Font.Size = 14
        .Range("H10").Font.Color = RGB(240, 136, 62)
        
        ' 工站利用率区域
        .Range("G12:I12").Merge
        With .Range("G12")
            .Value = "工站利用率"
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = RGB(139, 148, 158)
            .Interior.Color = RGB(22, 27, 34)
        End With
        
        ' 时间线区域
        .Range("G24:I24").Merge
        With .Range("G24")
            .Value = "时间线事件"
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = RGB(139, 148, 158)
            .Interior.Color = RGB(22, 27, 34)
        End With
        
        ' 时间线表头
        .Range("G25").Value = "时间"
        .Range("H25").Value = "事件"
        With .Range("G25:I25")
            .Font.Bold = True
            .Interior.Color = RGB(13, 17, 23)
            .Borders.LineStyle = xlContinuous
        End With
        
        ' 设置整体样式
        With .Cells
            .Interior.Color = RGB(13, 17, 23)
            .Font.Color = RGB(230, 237, 243)
        End With
        
        ' 隐藏网格线
        .Activate
        ActiveWindow.DisplayGridlines = False
    End With
End Sub

'==================================================
' 填充配置下拉框
'==================================================
Private Sub FillConfigCombo(ws As Worksheet)
    ' 在Excel中用数据验证模拟下拉框
    Dim configList As String
    Dim i As Integer
    
    For i = 1 To Configs.Count
        configList = configList & Configs(i).Name
        If i < Configs.Count Then configList = configList & ","
    Next i
    
    With ws.Range("C4").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=configList
    End With
    
    ws.Range("C4").Value = CurrentConfigName
End Sub

'==================================================
' 加载配置到工作表
'==================================================
Private Sub LoadConfigToSheet(ws As Worksheet, configName As String)
    Dim cfg As Config
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    Set cfg = Configs(configName)
    CurrentConfigName = configName
    CurrentConfig = cfg
    
    ' 清空工站区域
    ws.Range("C11:E20").ClearContents
    
    ' 填充产品数
    ws.Range("C7").Value = cfg.ProductCount
    
    ' 填充工站数据
    For i = 1 To cfg.StationCount
        ws.Range("C" & (10 + i)).Value = cfg.Stations(i).Name
        ws.Range("D" & (10 + i)).Value = cfg.Stations(i).ProcessTime
        ws.Range("E" & (10 + i)).Value = cfg.Stations(i).BufferCapacity
    Next i
    
    ws.Range("C4").Value = configName
    Exit Sub
    
ErrHandler:
    MsgBox "配置加载失败：" & Err.Description
End Sub

'==================================================
' 从工作表读取配置
'==================================================
Private Function ReadConfigFromSheet(ws As Worksheet) As Config
    Dim cfg As Config
    Dim i As Integer, count As Integer
    
    cfg.Name = CurrentConfigName
    cfg.ProductCount = ws.Range("C7").Value
    
    ' 统计有效工站数
    count = 0
    For i = 1 To 10
        If Trim(ws.Range("C" & (10 + i)).Value) <> "" Then
            count = count + 1
        End If
    Next i
    
    If count = 0 Then count = 1
    cfg.StationCount = count
    ReDim cfg.Stations(1 To count)
    
    For i = 1 To count
        cfg.Stations(i).Name = ws.Range("C" & (10 + i)).Value
        If cfg.Stations(i).Name = "" Then cfg.Stations(i).Name = "工站" & i
        cfg.Stations(i).ProcessTime = Val(ws.Range("D" & (10 + i)).Value)
        If cfg.Stations(i).ProcessTime <= 0 Then cfg.Stations(i).ProcessTime = 1
        cfg.Stations(i).BufferCapacity = Val(ws.Range("E" & (10 + i)).Value)
        If cfg.Stations(i).BufferCapacity <= 0 Then cfg.Stations(i).BufferCapacity = 1
    Next i
    
    ReadConfigFromSheet = cfg
End Function

'==================================================
' 保存当前配置
'==================================================
Public Sub SaveCurrentConfig()
    Dim ws As Worksheet
    Set ws = Sheets("UPH计算器")
    
    Dim cfg As Config
    cfg = ReadConfigFromSheet(ws)
    cfg.Name = CurrentConfigName
    
    ' 更新集合中的配置
    Configs.Remove CurrentConfigName
    Configs.Add cfg, cfg.Name
    CurrentConfig = cfg
    
    MsgBox "配置已保存！", vbInformation
End Sub

'==================================================
' 另存为新配置
'==================================================
Public Sub SaveAsNewConfig()
    Dim newName As String
    newName = InputBox("请输入新配置名称：", "另存为", CurrentConfigName & " 副本")
    
    If newName = "" Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = Sheets("UPH计算器")
    
    Dim cfg As Config
    cfg = ReadConfigFromSheet(ws)
    cfg.Name = newName
    
    Configs.Add cfg, newName
    CurrentConfigName = newName
    CurrentConfig = cfg
    
    FillConfigCombo ws
    ws.Range("C4").Value = newName
    
    MsgBox "配置已保存为：" & newName, vbInformation
End Sub

'==================================================
' 删除当前配置
'==================================================
Public Sub DeleteCurrentConfig()
    If Configs.Count <= 1 Then
        MsgBox "至少保留一个配置！", vbExclamation
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = Sheets("UPH计算器")
    
    Configs.Remove CurrentConfigName
    
    ' 切换到第一个配置
    Dim firstCfg As Config
    Set firstCfg = Configs(1)
    CurrentConfigName = firstCfg.Name
    
    FillConfigCombo ws
    LoadConfigToSheet ws, CurrentConfigName
    
    MsgBox "配置已删除", vbInformation
End Sub

'==================================================
' 添加工站
'==================================================
Public Sub AddStation()
    Dim ws As Worksheet
    Set ws = Sheets("UPH计算器")
    
    ' 找到第一个空行
    Dim i As Integer
    For i = 1 To 10
        If Trim(ws.Range("C" & (10 + i)).Value) = "" Then
            ws.Range("C" & (10 + i)).Value = "工站" & i
            ws.Range("D" & (10 + i)).Value = 10
            ws.Range("E" & (10 + i)).Value = 2
            Exit For
        End If
    Next i
    
    If i > 10 Then
        MsgBox "最多支持10个工站！", vbExclamation
    End If
End Sub

'==================================================
' 运行模拟计算
'==================================================
Public Sub RunSimulation()
    Dim ws As Worksheet
    Set ws = Sheets("UPH计算器")
    
    Dim cfg As Config
    cfg = ReadConfigFromSheet(ws)
    
    ' 运行模拟
    Dim result As SimulationResult
    result = Simulate(cfg)
    
    ' 显示结果
    DisplayResults ws, result, cfg
End Sub

'==================================================
' 模拟结果结构
'==================================================
Private Type SimulationResult
    TotalTime As Double
    ActualUPH As Double
    TheoryUPH As Double
    BottleneckIndex As Integer
    Utilization() As Double
    Completed As Integer
    Timeline() As TimeEvent
    TimelineCount As Integer
End Type

'==================================================
' 核心模拟算法 - 阻塞反压
'==================================================
Private Function Simulate(cfg As Config) As SimulationResult
    Dim nProducts As Integer, nStations As Integer
    nProducts = cfg.ProductCount
    nStations = cfg.StationCount
    
    Dim result As SimulationResult
    
    If nProducts = 0 Or nStations = 0 Then
        Simulate = result
        Exit Function
    End If
    
    ' 产品时间线
    ReDim productTimeline(1 To nProducts) As Collection
    Dim p As Integer, s As Integer
    For p = 1 To nProducts
        Set productTimeline(p) = New Collection
    Next p
    
    ' 缓冲区
    ReDim buffer(1 To nStations) As Collection
    ReDim bufferCap(1 To nStations) As Integer
    bufferCap(1) = 9999  ' 第一个工站输入无限制
    For s = 2 To nStations
        bufferCap(s) = cfg.Stations(s).BufferCapacity
    Next s
    
    For s = 1 To nStations
        Set buffer(s) = New Collection
    Next s
    
    ' 初始化：所有产品在第一个工站缓冲区
    For p = 1 To nProducts
        buffer(1).Add p
    Next p
    
    ' 工站状态
    ReDim machineProduct(1 To nStations) As Integer
    ReDim machineFinish(1 To nStations) As Double
    ReDim stationBlocked(1 To nStations) As Boolean
    ReDim blockedQueue(1 To nStations) As Collection
    
    For s = 1 To nStations
        machineProduct(s) = 0
        machineFinish(s) = 0
        stationBlocked(s) = False
        Set blockedQueue(s) = New Collection
    Next s
    
    ' 事件队列
    Dim events As Collection
    Set events = New Collection
    
    Dim currentTime As Double
    currentTime = 0
    
    ' 尝试启动工站处理
    Dim maxIter As Long, iter As Long
    maxIter = nProducts * nStations * 100 + 10000
    
    ' 初始启动
    For s = 1 To nStations
        TryStartStation s, currentTime, buffer, machineProduct, machineFinish, _
            productTimeline, events, stationBlocked, cfg
    Next s
    
    ' 事件循环
    Do While events.Count > 0 And iter < maxIter
        iter = iter + 1
        
        ' 找最早事件
        Dim minTime As Double, minIdx As Integer
        minTime = 1E+30
        minIdx = 1
        Dim i As Integer
        For i = 1 To events.Count
            If events(i) < minTime Then
                minTime = events(i)
                minIdx = i
            End If
        Next i
        
        currentTime = minTime
        
        ' 获取事件信息
        Dim evStation As Integer, evProduct As Integer
        ParseEvent events(minIdx), evStation, evProduct
        
        events.Remove minIdx
        
        ' 处理完成事件
        ' 工站完成
        machineProduct(evStation) = 0
        machineFinish(evStation) = 0
        
        ' 产品进入下一工站或完成
        If evStation < nStations Then
            Dim nextS As Integer
            nextS = evStation + 1
            
            If buffer(nextS).Count < bufferCap(nextS) Then
                ' 缓冲区未满
                buffer(nextS).Add evProduct
                TryStartStation nextS, currentTime, buffer, machineProduct, machineFinish, _
                    productTimeline, events, stationBlocked, cfg
            Else
                ' 缓冲区满，阻塞
                blockedQueue(evStation).Add evProduct
                stationBlocked(evStation) = True
            End If
        Else
            ' 最后工站完成，检查上游
            For s = 1 To nStations - 1
                CheckUnblock s, currentTime, buffer, bufferCap, blockedQueue, _
                    stationBlocked, machineProduct, machineFinish, productTimeline, _
                    events, cfg
            Next s
        End If
        
        ' 当前工站尝试处理下一个
        TryStartStation evStation, currentTime, buffer, machineProduct, machineFinish, _
            productTimeline, events, stationBlocked, cfg
        
        ' 检查上游阻塞
        For s = 1 To evStation - 1
            CheckUnblock s, currentTime, buffer, bufferCap, blockedQueue, _
                stationBlocked, machineProduct, machineFinish, productTimeline, _
                events, cfg
        Next s
    Loop
    
    ' 计算总时间
    Dim totalTime As Double
    totalTime = 0
    Dim completed As Integer
    completed = 0
    
    For p = 1 To nProducts
        If productTimeline(p).Count > 0 Then
            Dim lastEvent As TimeEvent
            lastEvent = productTimeline(p)(productTimeline(p).Count)
            If lastEvent.EndTime > totalTime Then
                totalTime = lastEvent.EndTime
            End If
            If productTimeline(p).Count = nStations Then
                completed = completed + 1
            End If
        End If
    Next p
    
    ' 计算利用率
    ReDim busyTime(1 To nStations) As Double
    ReDim utilization(1 To nStations) As Double
    
    For s = 1 To nStations
        busyTime(s) = 0
    Next s
    
    For p = 1 To nProducts
        Dim j As Integer
        For j = 1 To productTimeline(p).Count
            Dim te As TimeEvent
            te = productTimeline(p)(j)
            busyTime(te.StationIndex) = busyTime(te.StationIndex) + (te.EndTime - te.StartTime)
        Next j
    Next p
    
    For s = 1 To nStations
        If totalTime > 0 Then
            utilization(s) = busyTime(s) / totalTime
        Else
            utilization(s) = 0
        End If
    Next s
    
    ' 找瓶颈工站
    Dim btlIdx As Integer
    btlIdx = 1
    Dim maxU As Double
    maxU = 0
    For s = 1 To nStations
        If utilization(s) > maxU Then
            maxU = utilization(s)
            btlIdx = s
        End If
    Next s
    
    ' 理论UPH
    Dim slowest As Double
    slowest = 0
    For s = 1 To nStations
        If cfg.Stations(s).ProcessTime > slowest Then
            slowest = cfg.Stations(s).ProcessTime
        End If
    Next s
    
    Dim theoryUPH As Double, actualUPH As Double
    If slowest > 0 Then theoryUPH = 60 / slowest
    If totalTime > 0 Then actualUPH = nProducts / (totalTime / 60)
    
    ' 收集时间线事件
    Dim allEvents As Collection
    Set allEvents = New Collection
    
    For p = 1 To nProducts
        For j = 1 To productTimeline(p).Count
            Dim te2 As TimeEvent
            te2 = productTimeline(p)(j)
            allEvents.Add te2
        Next j
    Next p
    
    ' 排序并存储
    ReDim result.Timeline(1 To allEvents.Count)
    result.TimelineCount = allEvents.Count
    
    ' 简单冒泡排序
    For i = 1 To allEvents.Count - 1
        For j = i + 1 To allEvents.Count
            If allEvents(j).StartTime < allEvents(i).StartTime Then
                Dim temp As TimeEvent
                temp = allEvents(i)
                allEvents.Remove i
                allEvents.Add temp, Before:=j
                ' 重新获取
                Dim temp2 As TimeEvent
                temp2 = allEvents(j)
                allEvents.Remove j
                allEvents.Add temp2, Before:=i
            End If
        Next j
    Next j
    
    For i = 1 To allEvents.Count
        result.Timeline(i) = allEvents(i)
    Next i
    
    result.TotalTime = totalTime
    result.ActualUPH = actualUPH
    result.TheoryUPH = theoryUPH
    result.BottleneckIndex = btlIdx
    result.Utilization = utilization
    result.Completed = completed
    
    Simulate = result
End Function

' 尝试启动工站
Private Sub TryStartStation(s As Integer, currentTime As Double, buffer() As Collection, _
    machineProduct() As Integer, machineFinish() As Double, _
    productTimeline() As Collection, events As Collection, _
    stationBlocked() As Boolean, cfg As Config)
    
    If machineProduct(s) > 0 Then Exit Sub  ' 正在忙
    If stationBlocked(s) Then Exit Sub  ' 被阻塞
    If buffer(s).Count = 0 Then Exit Sub  ' 没有产品
    
    Dim p As Integer
    p = buffer(s)(1)
    buffer(s).Remove 1
    
    Dim startTime As Double, endTime As Double
    startTime = currentTime
    endTime = startTime + cfg.Stations(s).ProcessTime
    
    machineProduct(s) = p
    machineFinish(s) = endTime
    
    Dim te As TimeEvent
    te.StationIndex = s
    te.StartTime = startTime
    te.EndTime = endTime
    productTimeline(p).Add te
    
    ' 添加完成事件
    events.Add CreateEvent(endTime, s, p)
End Sub

' 检查解除阻塞
Private Sub CheckUnblock(s As Integer, currentTime As Double, buffer() As Collection, _
    bufferCap() As Integer, blockedQueue() As Collection, _
    stationBlocked() As Boolean, machineProduct() As Integer, _
    machineFinish() As Double, productTimeline() As Collection, _
    events As Collection, cfg As Config)
    
    Dim nStations As Integer
    nStations = cfg.StationCount
    
    If s >= nStations Then Exit Sub
    If Not stationBlocked(s) Then Exit Sub
    
    Dim nextS As Integer
    nextS = s + 1
    
    If buffer(nextS).Count < bufferCap(nextS) Then
        If blockedQueue(s).Count > 0 Then
            Dim p As Integer
            p = blockedQueue(s)(1)
            blockedQueue(s).Remove 1
            buffer(nextS).Add p
            TryStartStation nextS, currentTime, buffer, machineProduct, machineFinish, _
                productTimeline, events, stationBlocked, cfg
        End If
        
        If blockedQueue(s).Count = 0 Then
            stationBlocked(s) = False
            TryStartStation s, currentTime, buffer, machineProduct, machineFinish, _
                productTimeline, events, stationBlocked, cfg
        End If
    End If
End Sub

' 创建事件字符串
Private Function CreateEvent(time As Double, station As Integer, product As Integer) As String
    CreateEvent = time & "|" & station & "|" & product
End Function

' 解析事件字符串
Private Sub ParseEvent(ev As String, ByRef station As Integer, ByRef product As Integer)
    Dim parts() As String
    parts = Split(ev, "|")
    station = Val(parts(1))
    product = Val(parts(2))
End Sub

'==================================================
' 显示结果
'==================================================
Private Sub DisplayResults(ws As Worksheet, result As SimulationResult, cfg As Config)
    ' 清空之前的结果
    ws.Range("G13:I22").ClearContents
    ws.Range("G26:I50").ClearContents
    
    ' 主要指标
    ws.Range("H7").Value = Format(result.ActualUPH, "0.00")
    ws.Range("H8").Value = Format(result.TotalTime, "0") & " 分钟"
    ws.Range("H9").Value = Format(result.TheoryUPH, "0.00")
    ws.Range("H10").Value = cfg.Stations(result.BottleneckIndex).Name
    
    ' 利用率
    Dim s As Integer
    For s = 1 To cfg.StationCount
        ws.Range("G" & (12 + s)).Value = cfg.Stations(s).Name
        ws.Range("H" & (12 + s)).Value = Format(result.Utilization(s) * 100, "0") & "%"
        
        ' 瓶颈标记
        If s = result.BottleneckIndex Then
            ws.Range("G" & (12 + s) & ":H" & (12 + s)).Font.Color = RGB(240, 136, 62)
            ws.Range("I" & (12 + s)).Value = "← 瓶颈"
            ws.Range("I" & (12 + s)).Font.Color = RGB(240, 136, 62)
        Else
            ws.Range("G" & (12 + s) & ":H" & (12 + s)).Font.Color = RGB(230, 237, 243)
            ws.Range("I" & (12 + s)).Value = ""
        End If
    Next s
    
    ' 时间线事件（最多显示25条）
    Dim count As Integer
    count = result.TimelineCount
    If count > 25 Then count = 25
    
    Dim i As Integer
    For i = 1 To count
        Dim te As TimeEvent
        te = result.Timeline(i)
        ws.Range("G" & (25 + i)).Value = te.StartTime & "-" & te.EndTime & "m"
        ws.Range("H" & (25 + i)).Value = "产品" & te.StationIndex & " -> " & cfg.Stations(te.StationIndex).Name
    Next i
    
    MsgBox "计算完成！" & vbCrLf & _
           "完成产品: " & result.Completed & "/" & cfg.ProductCount & vbCrLf & _
           "总耗时: " & Format(result.TotalTime, "0") & " 分钟" & vbCrLf & _
           "UPH: " & Format(result.ActualUPH, "0.00"), vbInformation
End Sub

'==================================================
' 重置结果
'==================================================
Public Sub ResetResults()
    Dim ws As Worksheet
    Set ws = Sheets("UPH计算器")
    
    ws.Range("H7:H10").Value = "-"
    ws.Range("G13:I22").ClearContents
    ws.Range("G26:I50").ClearContents
End Sub
