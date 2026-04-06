Attribute VB_Name = "UPHCalculator"
'==================================================
' UPH 流水线计算器 - Excel VBA 版本
' 功能：模拟多工站流水线生产，计算UPH和瓶颈工站
'==================================================

Option Explicit

' 全局变量 - 使用简单类型避免Type定义问题
Private gProductCount As Integer
Private gStationCount As Integer
Private gStationNames() As String
Private gStationTimes() As Double
Private gStationBuffers() As Integer

Private gConfigs As Collection
Private gCurrentConfigName As String

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
    Set gConfigs = New Collection
    
    ' 创建默认配置
    Dim defaultCfg As String
    defaultCfg = "默认配置|13|焊接|6|2|切割|20|2|清洗|8|2"
    
    gConfigs.Add defaultCfg, "默认配置"
    gCurrentConfigName = "默认配置"
    
    ' 解析到全局变量
    ParseConfig defaultCfg
End Sub

' 解析配置字符串
Private Sub ParseConfig(cfgStr As String)
    Dim parts() As String
    parts = Split(cfgStr, "|")
    
    Dim idx As Integer
    idx = 0
    
    ' 跳过名称
    idx = 1
    gProductCount = Val(parts(idx))
    idx = idx + 1
    
    ' 计算工站数
    gStationCount = (UBound(parts) - 2) / 3
    
    ReDim gStationNames(1 To gStationCount)
    ReDim gStationTimes(1 To gStationCount)
    ReDim gStationBuffers(1 To gStationCount)
    
    Dim i As Integer
    For i = 1 To gStationCount
        gStationNames(i) = parts(idx)
        gStationTimes(i) = Val(parts(idx + 1))
        gStationBuffers(i) = Val(parts(idx + 2))
        idx = idx + 3
    Next i
End Sub

' 生成配置字符串
Private Function BuildConfig(name As String) As String
    Dim result As String
    result = name & "|" & gProductCount
    
    Dim i As Integer
    For i = 1 To gStationCount
        result = result & "|" & gStationNames(i) & "|" & gStationTimes(i) & "|" & gStationBuffers(i)
    Next i
    
    BuildConfig = result
End Function

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
    Dim configList As String
    Dim i As Integer
    
    For i = 1 To gConfigs.Count
        Dim cfgStr As String
        cfgStr = gConfigs(i)
        Dim parts() As String
        parts = Split(cfgStr, "|")
        configList = configList & parts(0)
        If i < gConfigs.Count Then configList = configList & ","
    Next i
    
    With ws.Range("C4").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=configList
    End With
    
    ws.Range("C4").Value = gCurrentConfigName
End Sub

'==================================================
' 加载配置到工作表
'==================================================
Private Sub LoadConfigToSheet(ws As Worksheet, configName As String)
    On Error GoTo ErrHandler
    
    Dim cfgStr As String
    cfgStr = gConfigs(configName)
    
    ParseConfig cfgStr
    gCurrentConfigName = configName
    
    ' 清空工站区域
    ws.Range("C11:E20").ClearContents
    
    ' 填充产品数
    ws.Range("C7").Value = gProductCount
    
    ' 填充工站数据
    Dim i As Integer
    For i = 1 To gStationCount
        ws.Range("C" & (10 + i)).Value = gStationNames(i)
        ws.Range("D" & (10 + i)).Value = gStationTimes(i)
        ws.Range("E" & (10 + i)).Value = gStationBuffers(i)
    Next i
    
    ws.Range("C4").Value = configName
    Exit Sub
    
ErrHandler:
    MsgBox "配置加载失败：" & Err.Description
End Sub

'==================================================
' 从工作表读取配置
'==================================================
Private Sub ReadConfigFromSheet(ws As Worksheet)
    gProductCount = ws.Range("C7").Value
    
    ' 统计有效工站数
    Dim count As Integer
    count = 0
    Dim i As Integer
    For i = 1 To 10
        If Trim(ws.Range("C" & (10 + i)).Value) <> "" Then
            count = count + 1
        End If
    Next i
    
    If count = 0 Then count = 1
    gStationCount = count
    
    ReDim gStationNames(1 To count)
    ReDim gStationTimes(1 To count)
    ReDim gStationBuffers(1 To count)
    
    For i = 1 To count
        gStationNames(i) = ws.Range("C" & (10 + i)).Value
        If gStationNames(i) = "" Then gStationNames(i) = "工站" & i
        gStationTimes(i) = Val(ws.Range("D" & (10 + i)).Value)
        If gStationTimes(i) <= 0 Then gStationTimes(i) = 1
        gStationBuffers(i) = Val(ws.Range("E" & (10 + i)).Value)
        If gStationBuffers(i) <= 0 Then gStationBuffers(i) = 1
    Next i
End Sub

'==================================================
' 保存当前配置
'==================================================
Public Sub SaveCurrentConfig()
    Dim ws As Worksheet
    Set ws = Sheets("UPH计算器")
    
    ReadConfigFromSheet ws
    
    Dim cfgStr As String
    cfgStr = BuildConfig(gCurrentConfigName)
    
    ' 更新集合中的配置
    gConfigs.Remove gCurrentConfigName
    gConfigs.Add cfgStr, gCurrentConfigName
    
    MsgBox "配置已保存！", vbInformation
End Sub

'==================================================
' 另存为新配置
'==================================================
Public Sub SaveAsNewConfig()
    Dim newName As String
    newName = InputBox("请输入新配置名称：", "另存为", gCurrentConfigName & " 副本")
    
    If newName = "" Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = Sheets("UPH计算器")
    
    ReadConfigFromSheet ws
    
    Dim cfgStr As String
    cfgStr = BuildConfig(newName)
    
    gConfigs.Add cfgStr, newName
    gCurrentConfigName = newName
    
    FillConfigCombo ws
    ws.Range("C4").Value = newName
    
    MsgBox "配置已保存为：" & newName, vbInformation
End Sub

'==================================================
' 删除当前配置
'==================================================
Public Sub DeleteCurrentConfig()
    If gConfigs.Count <= 1 Then
        MsgBox "至少保留一个配置！", vbExclamation
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = Sheets("UPH计算器")
    
    gConfigs.Remove gCurrentConfigName
    
    ' 切换到第一个配置
    Dim firstCfg As String
    firstCfg = gConfigs(1)
    Dim parts() As String
    parts = Split(firstCfg, "|")
    gCurrentConfigName = parts(0)
    
    FillConfigCombo ws
    LoadConfigToSheet ws, gCurrentConfigName
    
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
    
    ReadConfigFromSheet ws
    
    ' 运行模拟
    Dim totalTime As Double
    Dim actualUPH As Double
    Dim theoryUPH As Double
    Dim btlIdx As Integer
    Dim completed As Integer
    
    Dim util() As Double
    Dim timeline() As String
    Dim timelineCount As Integer
    
    Simulate totalTime, actualUPH, theoryUPH, btlIdx, completed, util, timeline, timelineCount
    
    ' 显示结果
    DisplayResults ws, totalTime, actualUPH, theoryUPH, btlIdx, completed, util, timeline, timelineCount
End Sub

'==================================================
' 核心模拟算法 - 阻塞反压
'==================================================
Private Sub Simulate(ByRef totalTime As Double, ByRef actualUPH As Double, _
    ByRef theoryUPH As Double, ByRef btlIdx As Integer, ByRef completed As Integer, _
    ByRef util() As Double, ByRef timeline() As String, ByRef timelineCount As Integer)
    
    Dim nProducts As Integer
    Dim nStations As Integer
    nProducts = gProductCount
    nStations = gStationCount
    
    If nProducts = 0 Or nStations = 0 Then
        totalTime = 0
        actualUPH = 0
        theoryUPH = 0
        btlIdx = 1
        completed = 0
        Exit Sub
    End If
    
    ' 产品时间线 - 使用字符串存储: "station|start|end"
    Dim productTimeline() As Collection
    ReDim productTimeline(1 To nProducts)
    Dim p As Integer, s As Integer
    For p = 1 To nProducts
        Set productTimeline(p) = New Collection
    Next p
    
    ' 缓冲区
    Dim buffer() As Collection
    ReDim buffer(1 To nStations)
    
    Dim bufferCap() As Integer
    ReDim bufferCap(1 To nStations)
    bufferCap(1) = 9999  ' 第一个工站输入无限制
    For s = 2 To nStations
        bufferCap(s) = gStationBuffers(s)
    Next s
    
    For s = 1 To nStations
        Set buffer(s) = New Collection
    Next s
    
    ' 初始化：所有产品在第一个工站缓冲区
    For p = 1 To nProducts
        buffer(1).Add p
    Next p
    
    ' 工站状态
    Dim machineProduct() As Integer
    Dim machineFinish() As Double
    Dim stationBlocked() As Boolean
    Dim blockedQueue() As Collection
    
    ReDim machineProduct(1 To nStations)
    ReDim machineFinish(1 To nStations)
    ReDim stationBlocked(1 To nStations)
    ReDim blockedQueue(1 To nStations)
    
    For s = 1 To nStations
        machineProduct(s) = 0
        machineFinish(s) = 0
        stationBlocked(s) = False
        Set blockedQueue(s) = New Collection
    Next s
    
    ' 事件队列 - 存储字符串 "time|station|product"
    Dim events As Collection
    Set events = New Collection
    
    Dim currentTime As Double
    currentTime = 0
    
    Dim maxIter As Long
    maxIter = nProducts * nStations * 100 + 10000
    
    ' 初始启动
    For s = 1 To nStations
        TryStartStation s, currentTime, buffer, machineProduct, machineFinish, _
            productTimeline, events, stationBlocked
    Next s
    
    ' 事件循环
    Dim iter As Long
    iter = 0
    
    Do While events.Count > 0 And iter < maxIter
        iter = iter + 1
        
        ' 找最早事件
        Dim minTime As Double
        Dim minIdx As Integer
        minTime = 1E+30
        minIdx = 1
        
        Dim i As Integer
        For i = 1 To events.Count
            Dim evParts() As String
            evParts = Split(events(i), "|")
            If Val(evParts(0)) < minTime Then
                minTime = Val(evParts(0))
                minIdx = i
            End If
        Next i
        
        currentTime = minTime
        
        ' 获取事件信息
        Dim evStr As String
        evStr = events(minIdx)
        Dim evParts2() As String
        evParts2 = Split(evStr, "|")
        Dim evStation As Integer
        Dim evProduct As Integer
        evStation = Val(evParts2(1))
        evProduct = Val(evParts2(2))
        
        events.Remove minIdx
        
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
                    productTimeline, events, stationBlocked
            Else
                ' 缓冲区满，阻塞
                blockedQueue(evStation).Add evProduct
                stationBlocked(evStation) = True
            End If
        Else
            ' 最后工站完成，检查上游
            For s = 1 To nStations - 1
                CheckUnblock s, currentTime, buffer, bufferCap, blockedQueue, _
                    stationBlocked, machineProduct, machineFinish, productTimeline, events
            Next s
        End If
        
        ' 当前工站尝试处理下一个
        TryStartStation evStation, currentTime, buffer, machineProduct, machineFinish, _
            productTimeline, events, stationBlocked
        
        ' 检查上游阻塞
        For s = 1 To evStation - 1
            CheckUnblock s, currentTime, buffer, bufferCap, blockedQueue, _
                stationBlocked, machineProduct, machineFinish, productTimeline, events
        Next s
    Loop
    
    ' 计算总时间
    totalTime = 0
    completed = 0
    
    For p = 1 To nProducts
        If productTimeline(p).Count > 0 Then
            Dim lastEv As String
            lastEv = productTimeline(p)(productTimeline(p).Count)
            Dim lastParts() As String
            lastParts = Split(lastEv, "|")
            If Val(lastParts(2)) > totalTime Then
                totalTime = Val(lastParts(2))
            End If
            If productTimeline(p).Count = nStations Then
                completed = completed + 1
            End If
        End If
    Next p
    
    ' 计算利用率
    Dim busyTime() As Double
    ReDim busyTime(1 To nStations)
    ReDim util(1 To nStations)
    
    For s = 1 To nStations
        busyTime(s) = 0
    Next s
    
    For p = 1 To nProducts
        Dim j As Integer
        For j = 1 To productTimeline(p).Count
            Dim teStr As String
            teStr = productTimeline(p)(j)
            Dim teParts() As String
            teParts = Split(teStr, "|")
            Dim teStation As Integer
            teStation = Val(teParts(0))
            busyTime(teStation) = busyTime(teStation) + (Val(teParts(2)) - Val(teParts(1)))
        Next j
    Next p
    
    For s = 1 To nStations
        If totalTime > 0 Then
            util(s) = busyTime(s) / totalTime
        Else
            util(s) = 0
        End If
    Next s
    
    ' 找瓶颈工站
    btlIdx = 1
    Dim maxU As Double
    maxU = 0
    For s = 1 To nStations
        If util(s) > maxU Then
            maxU = util(s)
            btlIdx = s
        End If
    Next s
    
    ' 理论UPH
    Dim slowest As Double
    slowest = 0
    For s = 1 To nStations
        If gStationTimes(s) > slowest Then
            slowest = gStationTimes(s)
        End If
    Next s
    
    If slowest > 0 Then theoryUPH = 60 / slowest
    If totalTime > 0 Then actualUPH = nProducts / (totalTime / 60)
    
    ' 收集时间线事件
    Dim allEvents As Collection
    Set allEvents = New Collection
    
    For p = 1 To nProducts
        For j = 1 To productTimeline(p).Count
            Dim teStr2 As String
            teStr2 = productTimeline(p)(j)
            allEvents.Add teStr2 & "|" & p
        Next j
    Next p
    
    ' 排序并存储
    timelineCount = allEvents.Count
    ReDim timeline(1 To timelineCount)
    
    ' 简单排序（按开始时间）
    For i = 1 To timelineCount
        timeline(i) = allEvents(i)
    Next i
    
    ' 冒泡排序
    For i = 1 To timelineCount - 1
        For j = i + 1 To timelineCount
            Dim tiParts() As String
            Dim tjParts() As String
            tiParts = Split(timeline(i), "|")
            tjParts = Split(timeline(j), "|")
            If Val(tjParts(1)) < Val(tiParts(1)) Then
                Dim temp As String
                temp = timeline(i)
                timeline(i) = timeline(j)
                timeline(j) = temp
            End If
        Next j
    Next j
End Sub

' 尝试启动工站
Private Sub TryStartStation(s As Integer, currentTime As Double, buffer() As Collection, _
    machineProduct() As Integer, machineFinish() As Double, _
    productTimeline() As Collection, events As Collection, _
    stationBlocked() As Boolean)
    
    If machineProduct(s) > 0 Then Exit Sub  ' 正在忙
    If stationBlocked(s) Then Exit Sub  ' 被阻塞
    If buffer(s).Count = 0 Then Exit Sub  ' 没有产品
    
    Dim p As Integer
    p = buffer(s)(1)
    buffer(s).Remove 1
    
    Dim startTime As Double
    Dim endTime As Double
    startTime = currentTime
    endTime = startTime + gStationTimes(s)
    
    machineProduct(s) = p
    machineFinish(s) = endTime
    
    ' 记录时间线: "station|start|end"
    productTimeline(p).Add s & "|" & startTime & "|" & endTime
    
    ' 添加完成事件: "time|station|product"
    events.Add endTime & "|" & s & "|" & p
End Sub

' 检查解除阻塞
Private Sub CheckUnblock(s As Integer, currentTime As Double, buffer() As Collection, _
    bufferCap() As Integer, blockedQueue() As Collection, _
    stationBlocked() As Boolean, machineProduct() As Integer, _
    machineFinish() As Double, productTimeline() As Collection, _
    events As Collection)
    
    If s >= gStationCount Then Exit Sub
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
                productTimeline, events, stationBlocked
        End If
        
        If blockedQueue(s).Count = 0 Then
            stationBlocked(s) = False
            TryStartStation s, currentTime, buffer, machineProduct, machineFinish, _
                productTimeline, events, stationBlocked
        End If
    End If
End Sub

'==================================================
' 显示结果
'==================================================
Private Sub DisplayResults(ws As Worksheet, totalTime As Double, actualUPH As Double, _
    theoryUPH As Double, btlIdx As Integer, completed As Integer, _
    util() As Double, timeline() As String, timelineCount As Integer)
    
    ' 清空之前的结果
    ws.Range("G13:I22").ClearContents
    ws.Range("G26:I50").ClearContents
    
    ' 主要指标
    ws.Range("H7").Value = Format(actualUPH, "0.00")
    ws.Range("H8").Value = Format(totalTime, "0") & " 分钟"
    ws.Range("H9").Value = Format(theoryUPH, "0.00")
    ws.Range("H10").Value = gStationNames(btlIdx)
    
    ' 利用率
    Dim s As Integer
    For s = 1 To gStationCount
        ws.Range("G" & (12 + s)).Value = gStationNames(s)
        ws.Range("H" & (12 + s)).Value = Format(util(s) * 100, "0") & "%"
        
        ' 瓶颈标记
        If s = btlIdx Then
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
    count = timelineCount
    If count > 25 Then count = 25
    
    Dim i As Integer
    For i = 1 To count
        Dim parts() As String
        parts = Split(timeline(i), "|")
        Dim teStation As Integer
        Dim teStart As Double
        Dim teEnd As Double
        Dim teProduct As Integer
        teStation = Val(parts(0))
        teStart = Val(parts(1))
        teEnd = Val(parts(2))
        teProduct = Val(parts(3))
        
        ws.Range("G" & (25 + i)).Value = teStart & "-" & teEnd & "m"
        ws.Range("H" & (25 + i)).Value = "产品" & teProduct & " -> " & gStationNames(teStation)
    Next i
    
    MsgBox "计算完成！" & vbCrLf & _
           "完成产品: " & completed & "/" & gProductCount & vbCrLf & _
           "总耗时: " & Format(totalTime, "0") & " 分钟" & vbCrLf & _
           "UPH: " & Format(actualUPH, "0.00"), vbInformation
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
