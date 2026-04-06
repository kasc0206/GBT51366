#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公共建筑碳排放计算报告 - VBA宏代码生成器
生成Excel-Word数据同步宏代码
"""

vba_code = '''
' ========================================
' 公共建筑碳排放计算报告 - Excel到Word数据同步宏
' 基于GB/T 51366-2019和JS/T 303-2026
' ========================================

Option Explicit

Dim wdApp As Object
Dim wdDoc As Object

Sub UpdateWordReport()
    On Error GoTo ErrorHandler
    
    Dim wsBasic As Worksheet, wsOperation As Worksheet
    Dim wsMaterials As Worksheet, wsConstruction As Worksheet
    Dim wsSummary As Worksheet
    Dim wordTemplatePath As String, wordReportPath As String
    
    Set wsBasic = ThisWorkbook.Sheets("项目基本信息")
    Set wsOperation = ThisWorkbook.Sheets("运行阶段-能源")
    Set wsMaterials = ThisWorkbook.Sheets("建材生产运输")
    Set wsConstruction = ThisWorkbook.Sheets("建造拆除阶段")
    Set wsSummary = ThisWorkbook.Sheets("碳排放汇总")
    
    wordTemplatePath = ThisWorkbook.Path & "\\公共建筑碳排放计算报告.docx"
    wordReportPath = ThisWorkbook.Path & "\\公共建筑碳排放计算报告_生成版.docx"
    
    If Dir(wordTemplatePath) = "" Then
        MsgBox "Word模板文件不存在：" & wordTemplatePath, vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo ErrorHandler
    
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Open(wordTemplatePath)
    
    Call UpdateBasicInfo(wsBasic)
    Call UpdateEnergyData(wsOperation)
    Call UpdateMaterialData(wsMaterials)
    Call UpdateConstructionData(wsConstruction)
    Call UpdateSummaryData(wsSummary)
    
    wdDoc.SaveAs2 wordReportPath
    MsgBox "Word报告已成功更新！" & vbCrLf & "保存路径：" & wordReportPath, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "更新过程中发生错误：" & Err.Description, vbCritical
End Sub

Sub UpdateBasicInfo(ws As Worksheet)
    Call FindAndReplace("[项目名称]", ws.Range("D4").Value)
    Call FindAndReplace("[建筑面积]", ws.Range("D7").Value)
    Call FindAndReplace("[计算年度]", ws.Range("D14").Value)
End Sub

Sub UpdateEnergyData(ws As Worksheet)
    Call FindAndReplace("[年用电量]", ws.Range("B5").Value)
    Call FindAndReplace("[年用气量]", ws.Range("B6").Value)
    Call FindAndReplace("[年用热量]", ws.Range("B7").Value)
End Sub

Sub UpdateMaterialData(ws As Worksheet)
    Call FindAndReplace("[建材生产阶段碳排放]", ws.Range("H67").Value)
    Call FindAndReplace("[建材运输阶段碳排放]", ws.Range("G84").Value)
End Sub

Sub UpdateConstructionData(ws As Worksheet)
    Call FindAndReplace("[建造阶段碳排放]", ws.Range("I44").Value)
    Call FindAndReplace("[拆除阶段碳排放]", ws.Range("H67").Value)
End Sub

Sub UpdateSummaryData(ws As Worksheet)
    Call FindAndReplace("[全生命周期总碳排放]", ws.Range("B24").Value)
    Call FindAndReplace("[单位面积碳排放]", ws.Range("C24").Value)
End Sub

Sub FindAndReplace(findText As String, replaceText As String)
    On Error Resume Next
    If wdDoc Is Nothing Then Exit Sub
    With wdApp.Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchCase = False
        .Execute Replace:=2
    End With
    On Error GoTo 0
End Sub

Sub ValidateData()
    Dim wsBasic As Worksheet
    Set wsBasic = ThisWorkbook.Sheets("项目基本信息")
    Dim missingFields As String, isValid As Boolean
    missingFields = "": isValid = True
    
    If wsBasic.Range("D4").Value = "" Then missingFields = missingFields & "项目名称" & vbCrLf: isValid = False
    If wsBasic.Range("D7").Value = "" Then missingFields = missingFields & "建筑面积" & vbCrLf: isValid = False
    
    If isValid Then
        MsgBox "数据验证通过！", vbInformation
    Else
        MsgBox "以下必填字段为空：" & vbCrLf & missingFields, vbExclamation
    End If
End Sub

Sub CreateToolbarButton()
    Dim cmdBar As Object, cmdButton As Object
    On Error Resume Next
    Application.CommandBars("碳排放报告").Delete
    On Error GoTo 0
    
    Set cmdBar = Application.CommandBars.Add(Name:="碳排放报告", Position:=1)
    cmdBar.Visible = True
    
    Set cmdButton = cmdBar.Controls.Add(Type:=1)
    With cmdButton
        .Caption = "更新Word报告"
        .OnAction = "UpdateWordReport"
        .Style = 3
        .FaceId = 18
    End With
    
    Set cmdButton = cmdBar.Controls.Add(Type:=1)
    With cmdButton
        .Caption = "数据验证"
        .OnAction = "ValidateData"
        .Style = 3
        .FaceId = 108
    End With
End Sub

Private Sub Workbook_Open()
    Call CreateToolbarButton
    MsgBox "碳排放计算工具已加载！", vbInformation
End Sub
'''

# 保存VBA代码
vba_path = __file__.replace("create_vba_macros.py", "Excel_Word_Sync_Macros.bas")
with open(vba_path, 'w', encoding='utf-8') as f:
    f.write(vba_code)

print(f"✅ VBA宏代码已保存至: {vba_path}")
print("\n使用说明：")
print("1. 打开Excel文件")
print("2. 按Alt+F11打开VBA编辑器")
print("3. 文件 → 导入文件 → 选择Excel_Word_Sync_Macros.bas")
print("4. 保存Excel为.xlsm格式（启用宏的工作簿）")
print("5. 重新打开Excel即可看到快捷工具栏")
