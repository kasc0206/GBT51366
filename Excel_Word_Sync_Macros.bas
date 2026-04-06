
' ========================================
' 公共建筑碳排放计算报告 - Excel到Word数据同步宏
' 基于GB/T 51366-2019和JS/T 303-2026
' ========================================

Option Explicit

' 定义Word应用程序对象
Dim wdApp As Object
Dim wdDoc As Object

' ========================================
' 主同步函数：将Excel数据更新到Word文档
' ========================================
Sub UpdateWordReport()
    On Error GoTo ErrorHandler
    
    Dim wsBasic As Worksheet
    Dim wsOperation As Worksheet
    Dim wsMaterials As Worksheet
    Dim wsConstruction As Worksheet
    Dim wsSummary As Worksheet
    Dim wordTemplatePath As String
    Dim wordReportPath As String
    
    ' 设置工作表
    Set wsBasic = ThisWorkbook.Sheets("项目基本信息")
    Set wsOperation = ThisWorkbook.Sheets("运行阶段碳排放")
    Set wsMaterials = ThisWorkbook.Sheets("建材生产及运输")
    Set wsConstruction = ThisWorkbook.Sheets("建造及拆除阶段")
    Set wsSummary = ThisWorkbook.Sheets("碳排放汇总")
    
    ' 获取Word模板路径（与Excel同目录）
    wordTemplatePath = ThisWorkbook.Path & "\公共建筑碳排放计算报告模板.docx"
    wordReportPath = ThisWorkbook.Path & "\公共建筑碳排放计算报告_生成版.docx"
    
    ' 检查Word模板是否存在
    If Dir(wordTemplatePath) = "" Then
        MsgBox "Word模板文件不存在：" & wordTemplatePath, vbExclamation
        Exit Sub
    End If
    
    ' 创建或获取Word应用程序
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo ErrorHandler
    
    wdApp.Visible = True
    
    ' 打开Word模板
    Set wdDoc = wdApp.Documents.Open(wordTemplatePath)
    
    ' 更新Word文档中的数据
    Call UpdateBasicInfo(wsBasic)
    Call UpdateEnergyData(wsOperation)
    Call UpdateMaterialData(wsMaterials)
    Call UpdateConstructionData(wsConstruction)
    Call UpdateSummaryData(wsSummary)
    
    ' 另存为报告文件
    wdDoc.SaveAs2 wordReportPath
    
    MsgBox "Word报告已成功更新！" & vbCrLf & "保存路径：" & wordReportPath, vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "更新过程中发生错误：" & Err.Description & vbCrLf & "错误代码：" & Err.Number, vbCritical
    If Not wdDoc Is Nothing Then
        wdDoc.Close SaveChanges:=False
    End If
    If Not wdApp Is Nothing Then
        wdApp.Quit
    End If
End Sub

' ========================================
' 更新项目基本信息
' ========================================
Sub UpdateBasicInfo(ws As Worksheet)
    Dim projectName As String
    Dim buildingAddress As String
    Dim buildingType As String
    Dim buildingArea As Double
    Dim calculationYear As String
    
    ' 从Excel读取数据
    projectName = ws.Range("D4").Value
    buildingAddress = ws.Range("D5").Value
    buildingType = ws.Range("D6").Value
    buildingArea = ws.Range("D7").Value
    calculationYear = ws.Range("D14").Value
    
    ' 在Word中查找并替换占位符
    Call FindAndReplace("[项目名称]", projectName)
    Call FindAndReplace("[建筑地址]", buildingAddress)
    Call FindAndReplace("[建筑类型]", buildingType)
    Call FindAndReplace("[建筑面积]", Format(buildingArea, "#,##0"))
    Call FindAndReplace("[计算年度]", calculationYear)
End Sub

' ========================================
' 更新能源消耗数据
' ========================================
Sub UpdateEnergyData(ws As Worksheet)
    Dim electricity As Double
    Dim naturalGas As Double
    Dim heating As Double
    Dim totalEmission As Double
    
    ' 从Excel读取数据
    electricity = ws.Range("B5").Value
    naturalGas = ws.Range("B6").Value
    heating = ws.Range("B7").Value
    totalEmission = ws.Range("G13").Value
    
    ' 更新Word
    Call FindAndReplace("[年用电量]", Format(electricity, "#,##0"))
    Call FindAndReplace("[年用气量]", Format(naturalGas, "#,##0"))
    Call FindAndReplace("[年用热量]", Format(heating, "#,##0"))
    Call FindAndReplace("[运行阶段总碳排放]", Format(totalEmission, "#,##0.00"))
End Sub

' ========================================
' 更新建材数据
' ========================================
Sub UpdateMaterialData(ws As Worksheet)
    Dim productionEmission As Double
    Dim transportEmission As Double
    Dim totalMaterialEmission As Double
    
    productionEmission = ws.Range("H19").Value
    transportEmission = ws.Range("G29").Value
    totalMaterialEmission = productionEmission + transportEmission
    
    Call FindAndReplace("[建材生产阶段碳排放]", Format(productionEmission, "#,##0.00"))
    Call FindAndReplace("[建材运输阶段碳排放]", Format(transportEmission, "#,##0.00"))
    Call FindAndReplace("[建材总碳排放]", Format(totalMaterialEmission, "#,##0.00"))
End Sub

' ========================================
' 更新建造及拆除数据
' ========================================
Sub UpdateConstructionData(ws As Worksheet)
    Dim constructionEmission As Double
    Dim demolitionEmission As Double
    
    constructionEmission = ws.Range("H18").Value
    demolitionEmission = ws.Range("H24").Value
    
    Call FindAndReplace("[建造阶段碳排放]", Format(constructionEmission, "#,##0.00"))
    Call FindAndReplace("[拆除阶段碳排放]", Format(demolitionEmission, "#,##0.00"))
End Sub

' ========================================
' 更新汇总数据
' ========================================
Sub UpdateSummaryData(ws As Worksheet)
    Dim totalLifecycleEmission As Double
    Dim emissionIntensity As Double
    Dim buildingArea As Double
    
    totalLifecycleEmission = ws.Range("B24").Value
    buildingArea = ws.Range("B6").Value ' 从项目基本信息获取
    
    If buildingArea > 0 Then
        emissionIntensity = totalLifecycleEmission * 1000 / buildingArea
    Else
        emissionIntensity = 0
    End If
    
    Call FindAndReplace("[全生命周期总碳排放]", Format(totalLifecycleEmission, "#,##0.00"))
    Call FindAndReplace("[单位面积碳排放]", Format(emissionIntensity, "#,##0.00"))
End Sub

' ========================================
' 查找并替换Word文档中的文本
' ========================================
Sub FindAndReplace(findText As String, replaceText As String)
    On Error Resume Next
    
    If wdDoc Is Nothing Then Exit Sub
    
    With wdApp.Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = 1 ' wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=2 ' wdReplaceAll
    End With
    
    On Error GoTo 0
End Sub

' ========================================
' 自动更新：当Excel数据变化时自动更新Word
' ========================================
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' 可选：启用此功能可实现实时同步
    ' 注意：这可能会影响性能，建议手动触发
    
    ' If Sh.Name = "碳排放汇总" Then
    '     Call UpdateWordReport
    ' End If
End Sub

' ========================================
' 导出报告为PDF
' ========================================
Sub ExportToPDF()
    On Error GoTo ErrorHandler
    
    Dim wordTemplatePath As String
    Dim pdfPath As String
    
    wordTemplatePath = ThisWorkbook.Path & "\公共建筑碳排放计算报告模板.docx"
    pdfPath = ThisWorkbook.Path & "\公共建筑碳排放计算报告.pdf"
    
    ' 先更新Word
    Call UpdateWordReport
    
    ' 导出为PDF
    If Not wdDoc Is Nothing Then
        wdDoc.ExportAsFixedFormat OutputFileName:=pdfPath, _
            ExportFormat:=17 ' wdExportFormatPDF
        MsgBox "PDF报告已导出：" & pdfPath, vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "导出PDF时发生错误：" & Err.Description, vbCritical
End Sub

' ========================================
' 数据验证：检查所有必填字段
' ========================================
Sub ValidateData()
    Dim wsBasic As Worksheet
    Dim wsOperation As Worksheet
    Dim missingFields As String
    Dim isValid As Boolean
    
    Set wsBasic = ThisWorkbook.Sheets("项目基本信息")
    Set wsOperation = ThisWorkbook.Sheets("运行阶段碳排放")
    
    missingFields = ""
    isValid = True
    
    ' 检查必填字段
    If wsBasic.Range("D4").Value = "" Then
        missingFields = missingFields & "项目名称" & vbCrLf
        isValid = False
    End If
    
    If wsBasic.Range("D7").Value = "" Then
        missingFields = missingFields & "建筑面积" & vbCrLf
        isValid = False
    End If
    
    If wsOperation.Range("B5").Value = "" Then
        missingFields = missingFields & "年用电量" & vbCrLf
        isValid = False
    End If
    
    If isValid Then
        MsgBox "数据验证通过！所有必填字段已填写。", vbInformation
    Else
        MsgBox "以下必填字段为空：" & vbCrLf & vbCrLf & missingFields, vbExclamation
    End If
End Sub

' ========================================
' 创建快捷按钮
' ========================================
Sub CreateToolbarButton()
    Dim cmdBar As Object
    Dim cmdButton As Object
    
    On Error Resume Next
    Application.CommandBars("碳排放报告").Delete
    On Error GoTo 0
    
    Set cmdBar = Application.CommandBars.Add(Name:="碳排放报告", Position:=1) ' msoBarTop
    cmdBar.Visible = True
    
    Set cmdButton = cmdBar.Controls.Add(Type:=1) ' msoControlButton
    With cmdButton
        .Caption = "更新Word报告"
        .OnAction = "UpdateWordReport"
        .Style = 3 ' msoButtonIconAndCaption
        .FaceId = 18 ' 保存图标
    End With
    
    Set cmdButton = cmdBar.Controls.Add(Type:=1)
    With cmdButton
        .Caption = "数据验证"
        .OnAction = "ValidateData"
        .Style = 3
        .FaceId = 108 ' 检查图标
    End With
    
    Set cmdButton = cmdBar.Controls.Add(Type:=1)
    With cmdButton
        .Caption = "导出PDF"
        .OnAction = "ExportToPDF"
        .Style = 3
        .FaceId = 4 ' 打印图标
    End With
End Sub

' ========================================
' 使用帮助
' ========================================
Sub ShowHelp()
    Dim helpMsg As String
    
    helpMsg = "【公共建筑碳排放计算报告系统】" & vbCrLf & vbCrLf
    helpMsg = helpMsg & "使用说明：" & vbCrLf
    helpMsg = helpMsg & "1. 在Excel各工作表中填写基础数据（黄色单元格）" & vbCrLf
    helpMsg = helpMsg & "2. 绿色单元格为自动计算结果，无需手动填写" & vbCrLf
    helpMsg = helpMsg & "3. 点击"更新Word报告"按钮生成Word文档" & vbCrLf
    helpMsg = helpMsg & "4. Word文档中的数据与Excel保持同步" & vbCrLf & vbCrLf
    helpMsg = helpMsg & "快捷工具栏：" & vbCrLf
    helpMsg = helpMsg & "- 更新Word报告：生成最新Word文档" & vbCrLf
    helpMsg = helpMsg & "- 数据验证：检查必填字段是否填写完整" & vbCrLf
    helpMsg = helpMsg & "- 导出PDF：将Word报告导出为PDF格式" & vbCrLf & vbCrLf
    helpMsg = helpMsg & "技术支持：基于GB/T 51366-2019和JS/T 303-2026"
    
    MsgBox helpMsg, vbInformation, "帮助"
End Sub

' ========================================
' 工作簿打开时自动创建工具栏
' ========================================
Private Sub Workbook_Open()
    Call CreateToolbarButton
    MsgBox "碳排放计算工具已加载！" & vbCrLf & "请使用顶部工具栏中的按钮操作。", vbInformation
End Sub
