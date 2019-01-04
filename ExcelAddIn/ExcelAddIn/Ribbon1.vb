Imports System.Diagnostics
Imports System.IO
Imports Microsoft.Office.Tools.Ribbon
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Data
Imports System.Windows.Forms

Public Class Ribbon1
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

    End Sub
    Public Function hideSortExcel()

        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '获取excel行数
        Dim a = getLineNumber()
        '填充仪表标题颜色
        fillColor()

        '调取填充表格函数
        fillTable()

        Dim xl = Globals.ThisAddIn.Application
        xl.ActiveWindow.SplitColumn = 1
        xl.ActiveWindow.SplitRow = 0
        xl.Application.ActiveWindow.FreezePanes = True

        xlSheet.Cells(1, 75).value = ""
        xlSheet.Cells(1, 50).value = ""
        xlSheet.Cells(1, 125).value = ""



        For i = 7 To a
            xlSheet.Cells(i, 50).Interior.ColorIndex = 0  '设置单元格背景颜色
            xlSheet.Cells(i, 125).Interior.ColorIndex = 0  '设置单元格背景颜色
            xlSheet.Cells(i, 75).Interior.ColorIndex = 0  '设置单元格背景颜色
            xlSheet.Cells(i, 4).WrapText = False
            Dim lenfour = Len(xlSheet.Cells(i, 4).Value)
            If lenfour <> 4 Then
                xlSheet.Cells(i, 4).Interior.ColorIndex = 27  '设置单元格背景颜色
            End If

            xlSheet.Rows(i).RowHeight = 20        '设置行高
            If (CStr(xlSheet.Cells(i, 48).value) = "1" And CStr(xlSheet.Cells(i, 53).value) = "1") Or ((CStr(xlSheet.Cells(i, 48).value) = "" Or CStr(xlSheet.Cells(i, 48).value) = Chr(10)) And (CStr(xlSheet.Cells(i, 53).value) = "2")) Or ((CStr(xlSheet.Cells(i, 53).value) = "" Or CStr(xlSheet.Cells(i, 53).value) = Chr(10)) And (CStr(xlSheet.Cells(i, 48).value) = "2")) Then
                xlSheet.Columns(50).Hidden = False
                xlSheet.Cells(1, 50).value = "Please input 2OO2/1OO2."
                xlSheet.Cells(1, 50).Interior.ColorIndex = 27  '设置单元格背景颜色   
                xlSheet.Cells(i, 50).Interior.ColorIndex = 27  '设置单元格背景颜色
            End If
            If CStr(xlSheet.Cells(i, 123).value) = "1" Then
                xlSheet.Columns(125).Hidden = False
                xlSheet.Cells(1, 125).value = "Please input direction CLOSE/OPEN."
                xlSheet.Cells(1, 125).Interior.ColorIndex = 27  '设置单元格背景颜色
                xlSheet.Cells(i, 125).Interior.ColorIndex = 27  '设置单元格背景颜色
            End If
            If CStr(xlSheet.Cells(i, 72).value) <> "" And CStr(xlSheet.Cells(i, 72).value) <> Chr(10) Then
                xlSheet.Columns(75).Hidden = False
                xlSheet.Cells(1, 75).value = "Include PPS or VG must input PPS function."
                xlSheet.Cells(1, 75).Interior.ColorIndex = 27  '设置单元格背景颜色
                xlSheet.Cells(i, 75).Interior.ColorIndex = 27  '设置单元格背景颜色
            ElseIf InStr(CStr(xlSheet.Cells(i, 27).value), "VG") Then
                If InStr(LCase(CStr(xlSheet.Cells(i, 26).value)), "delete") Then
                Else
                    xlSheet.Columns(75).Hidden = False
                    xlSheet.Cells(1, 75).value = "Include PPS or VG must input PPS function."
                    xlSheet.Cells(1, 75).Interior.ColorIndex = 27  '设置单元格背景颜色
                    xlSheet.Cells(i, 75).Interior.ColorIndex = 27  '设置单元格背景颜色
                End If
            End If
        Next

        For i = 1 To 197
            xlSheet.Columns(i).AutoFit  '设置自适应列宽
        Next
        xlSheet.Columns(13).ColumnWidth = 20        '设置列宽
        xlSheet.Cells(1, 13).WrapText = True

        '根据条件隐藏前15列空列      
        For i = 1 To 15
            Dim sum = 0
            For j = 7 To CInt(a)
                If (CStr(xlSheet.Cells(j, i).value) = "") Or (CStr(xlSheet.Cells(j, i).value) = Chr(10)) Then
                    sum = sum + 1
                Else
                    Exit For
                End If
            Next
            If sum = a - 6 Then xlSheet.Columns(i).Hidden = True
        Next
        '根据条件隐藏后面的空列
        For i = 16 To 190 Step 5
            Dim sum = 0
            Dim sum1 = 0
            Dim sum2 = 0

            For j = 7 To CInt(a)
                If InStr(LCase(CStr(xlSheet.Cells(j, 196).Value)), "comment") Then
                Else
                    xlSheet.Cells(j, 196).value = "Comment:"
                End If
                xlSheet.Cells(j, 196).Font.Color = 255
                If (CStr(xlSheet.Cells(j, i).value) = "") Or (CStr(xlSheet.Cells(j, i).value) = Chr(10)) Then
                    sum = sum + 1
                Else
                    Exit For
                End If
            Next
            For j = 7 To CInt(a) 'Gauge Block如果写了NoPressureGauge，也会隐藏列
                If (CStr(xlSheet.Cells(j, i + 1).value) = "") Or (CStr(xlSheet.Cells(j, i + 1).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(j, i + 1).value)), "nopressuregauge") Then
                    sum = sum + 1
                End If
            Next
            For j = 7 To CInt(a)
                If (CStr(xlSheet.Cells(j, i + 2).value) = "") Or (CStr(xlSheet.Cells(j, i + 2).value) = Chr(10)) Then
                    sum = sum + 1
                End If
            Next
            For j = 7 To CInt(a)
                If (CStr(xlSheet.Cells(j, i + 3).value) = "") Or (CStr(xlSheet.Cells(j, i + 3).value) = Chr(10)) Then
                    sum1 = sum1 + 1
                End If
            Next
            For j = 7 To CInt(a)
                If (CStr(xlSheet.Cells(j, i + 4).value) = "") Or (CStr(xlSheet.Cells(j, i + 4).value) = Chr(10)) Then
                    sum2 = sum2 + 1
                End If
            Next
            If sum = (3 * a - 18) Then
                xlSheet.Columns(i).Hidden = True
                xlSheet.Columns(i + 1).Hidden = True
                xlSheet.Columns(i + 2).Hidden = True
            End If
            If sum1 = a - 6 Then
                xlSheet.Columns(i + 3).Hidden = True
            End If
            If sum2 = a - 6 Then
                xlSheet.Columns(i + 4).Hidden = True
            End If
        Next
        For i = 191 To 196
            Dim sum = 0
            For j = 7 To CInt(a)
                If (CStr(xlSheet.Cells(j, i).value) = "") Or (CStr(xlSheet.Cells(j, i).value) = Chr(10)) Then
                    sum = sum + 1
                Else
                    Exit For
                End If
            Next
            If sum = a - 6 Then xlSheet.Columns(i).Hidden = True
        Next
        If InStr(CStr(xlSheet.Cells(1, 50).value), "Please input 2OO2/1OO2.") Then
            xlSheet.Columns(50).Hidden = False
        End If
        If InStr(CStr(xlSheet.Cells(1, 125).value), "Please input direction CLOSE/OPEN.") Then
            xlSheet.Columns(125).Hidden = False
        End If
        If InStr(CStr(xlSheet.Cells(1, 75).value), "Include PPS or VG must input PPS function.") Then
            xlSheet.Columns(75).Hidden = False
        End If

    End Function
    Public Function fillColor()
        '获取excel行数
        Dim a = getLineNumber()
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        For i = 17 To 190 Step 5
            If i = 37 Or i = 42 Or i = 57 Or i = 82 Or i = 87 Or i = 92 Or i = 127 Or i = 142 Or i = 147 Or i = 152 Or i = 157 Or i = 162 Or i = 167 Or i = 172 Or i = 177 Or i = 182 Or i = 187 Then
                xlSheet.Cells(5, i).Interior.ColorIndex = 36  '设置单元格背景颜色
            Else
                xlSheet.Cells(5, i).Interior.ColorIndex = 4  '设置单元格背景颜色
            End If
            For j = 7 To a
                If InStr(LCase(CStr(xlSheet.Cells(j, i + 2).Value)), "comment") Then
                Else
                    xlSheet.Cells(j, i + 2).Interior.ColorIndex = 0  '设置单元格背景颜色
                End If
            Next
        Next
    End Function
    Public Function confEleFunc(Arr1， a)

        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        For i = 7 To a
            xlSheet.Cells(i, 194).Interior.ColorIndex = 0  '设置单元格背景颜色

            '判断气缸失电方向与销售下单选择的方向是否一致
            If Arr1(3, i - 7, 0) = 0 Then
                '若dif值为0，说明程序有部分气缸未覆盖
                MsgBox("程序未完全覆盖气缸")
            ElseIf Arr1(3, i - 7, 0) = 3 Then
                '双作用气缸判断失电要求有没有
                If Arr1(2, i - 7, 1) = "3" Then
                    xlSheet.Cells(i, 194).value += "Missing electric fail action!!!"
                    xlSheet.Cells(i, 194).Interior.ColorIndex = 20  '设置单元格背景颜色
                End If
            ElseIf Arr1(2, i - 7, 1) = "3" Then
                '缺少失电方向
                xlSheet.Cells(i, 194).value += "Missing electric fail action!!!"
                xlSheet.Cells(i, 194).Interior.ColorIndex = 20  '设置单元格背景颜色
            ElseIf Arr1(3, i - 7, 0) <> Arr1(2, i - 7, 1) Then
                'dif<>Arr(i - 7, 1)，则需要写note,note列写“Please confrm the fail electric action!!!”
                xlSheet.Cells(i, 194).value += "Please confrm the electric fail action!!!"
                xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
                'dif=Arr(i - 7, 1)，则不需要写note
            End If
        Next

    End Function
    Public Function fillTable()  '191~194列自动填充文字说明

        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        xlSheet.Cells(6, 191).value = "SITWA Drawing(For reference)"
        xlSheet.Cells(6, 192).value = "Simular SITWA Drawing(Lack AFR for reference)"
        xlSheet.Cells(6, 193).value = "NO SITWA Drawing(For reference)"
        xlSheet.Cells(6, 194).value = "Note"
        xlSheet.Cells(6, 195).value = "Configuration"
        xlSheet.Cells(6, 196).value = "Add accessories"

        '生成autofilter
        xlSheet.Range("A6:GN6").AutoFilter(Field:=1)

    End Function
    Public Function getLineNumber()  '获取excel行数函数

        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        Dim a = 5
        For k = 7 To 200
            If CStr(xlSheet.Cells(k, 1).value) = "" Then
                a = k - 1
                Exit For
            End If
        Next
        Return a
    End Function
    Public Function insQuantityCheck(Arr1, a)

        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim b = a - 7
        Dim flagNum(b)

        For i = 7 To a
            flagNum(i - 7) = "0"  '值为0表示符合SITWA，值为1表示不符合SITWA
            xlSheet.Cells(i, 194).Interior.ColorIndex = 0  '设置单元格背景颜色

            'AFR判断
            If Arr1(4, i - 7, 0) = "0" Then
                flagNum(i - 7) = "1"
            ElseIf Arr1(4, i - 7, 0) <> "0" And Arr1(4, i - 7, 0) <> "1" Then
                flagNum(i - 7) = "1"
                xlSheet.Cells(i, 194).value += "AFR quantity more than 1!!!"
                xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
            End If

            'PPS判断
            If Arr1(4, i - 7, 1) <> "0" And Arr1(4, i - 7, 1) <> "1" Then
                flagNum(i - 7) = "1"
                xlSheet.Cells(i, 194).value += "PPS quantity more than 1!!!"
                xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
            End If

            'QEV判断
            If Arr1(4, i - 7, 2) <> "0" And Arr1(4, i - 7, 2) <> "1" Then
                flagNum(i - 7) = "1"
                xlSheet.Cells(i, 194).value += "QEV quantity more than 1!!!"
                xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
            End If

            'NV判断
            If Arr1(4, i - 7, 3) <> "0" And Arr1(4, i - 7, 3) <> "1" Then
                flagNum(i - 7) = "1"
                xlSheet.Cells(i, 194).value += "NV quantity more than 1!!!"
                xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
            End If

            'SCV判断
            If Arr1(4, i - 7, 4) <> "0" And Arr1(4, i - 7, 4) <> "1" And Arr1(4, i - 7, 4) <> "2" Then
                flagNum(i - 7) = "1"
                xlSheet.Cells(i, 194).value += "SCV quantity more than 2!!!"
                xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
            End If

            'SOV判断
            If Arr1(4, i - 7, 5) <> "0" And Arr1(4, i - 7, 5) <> "1" And Arr1(4, i - 7, 5) <> "2" Then
                flagNum(i - 7) = "1"
                xlSheet.Cells(i, 194).value += "SOV quantity more than 2!!!"
                xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
            End If

            'VB 判断
            If Arr1(2, i - 7, 0) = "1" Then  '单作用气缸
                If Arr1(4, i - 7, 6) <> "0" And Arr1(4, i - 7, 6) <> "1" Then
                    flagNum(i - 7) = "1"
                    xlSheet.Cells(i, 194).value += "VB quantity can't match SITWA!!!"
                    xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色               
                End If
            ElseIf Arr1(2, i - 7, 0) = "2" Then  '双作用气缸
                If Arr1(4, i - 7, 6) <> "0" And Arr1(4, i - 7, 6) <> "2" Then
                    flagNum(i - 7) = "1"
                    xlSheet.Cells(i, 194).value += "VB quantity can't match SITWA!!!"
                    xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
                End If
            End If

            'CV判断 
            If Arr1(4, i - 7, 7) = "0" Then
                If Arr1(0, i - 7, 4) = "" Then  'SITWA AOV==0
                    Arr1(1, i - 7, 10) = ""  '非SITWA  CV
                ElseIf Arr1(0, i - 7, 4) = "AOV+" And Arr1(0, i - 7, 1) = "" And Arr1(1, i - 7, 1) = "pos_" Then   'SITWA 有AOV无VB有定位器
                    Arr1(1, i - 7, 10) = ""  '非SITWA  CV
                ElseIf Arr1(0, i - 7, 4) = "" And Arr1(0, i - 7, 1) = "VB+" And Arr1(1, i - 7, 1) = "pos_" Then   'SITWA 有AOV无VB有定位器
                    Arr1(1, i - 7, 10) = ""  '非SITWA  CV
                Else
                    flagNum(i - 7) = "1"
                    xlSheet.Cells(i, 194).value += "Need Check valve!!!"
                    xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
                End If
            ElseIf Arr1(2, i - 7, 0) = "1" Then    '单作用气缸 
                If Arr1(4, i - 7, 7) = "1" Then
                    If Arr1(0, i - 7, 4) = "AOV+" Then
                        Arr1(1, i - 7, 10) = "cv1_"  '非SITWA  CV
                    Else
                        flagNum(i - 7) = "1"
                        xlSheet.Cells(i, 194).value += "Normally no need CV!!!"
                        xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
                    End If
                Else
                    flagNum(i - 7) = "1"
                    xlSheet.Cells(i, 194).value += "The CV quantity can't match SITWA!!!"
                    xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
                End If
                If Arr1(0, i - 7, 0) = "POS+" And Arr1(0, i - 7, 1) = "VB+" And Arr1(0, i - 7, 7) = "VT+" And Arr1(0, i - 7, 4) = "AOV+" Then
                    xlSheet.Cells(i, 194).value += "Better use 2 CV!!!"
                    xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
                ElseIf Arr1(0, i - 7, 0) = "VG+" And Arr1(0, i - 7, 1) = "VB+" And Arr1(0, i - 7, 7) = "VT+" And Arr1(0, i - 7, 4) = "AOV+" Then
                    xlSheet.Cells(i, 194).value += "Better use 2 CV!!!"
                    xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
                End If
            ElseIf Arr1(2, i - 7, 0) = "2" Then  '双作用气缸
                If Arr1(4, i - 7, 7) = "2" Then
                    If Arr1(0, i - 7, 4) = "AOV+" And Arr1(1, i - 7, 1) = "" And Arr1(0, i - 7, 7) = "VT+" Then  '不带定位器，带AOV,带VT的配置
                        Arr1(1, i - 7, 10) = "cv2_"  '非SITWA  CV
                    Else
                        flagNum(i - 7) = "1"
                        xlSheet.Cells(i, 194).value += "The CV quantity can't match SITWA!!!"
                        xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
                    End If
                ElseIf Arr1(4, i - 7, 7) = "1" Then
                    If Arr1(0, i - 7, 4) = "AOV+" And Arr1(1, i - 7, 1) = "" And Arr1(0, i - 7, 7) = "VT+" Then '不带定位器，带AOV,带VT的配置
                        flagNum(i - 7) = "1"
                        xlSheet.Cells(i, 194).value += "The CV quantity should be 2!!!"
                        xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
                    ElseIf Arr1(0, i - 7, 0) = "VG+" And Arr1(0, i - 7, 1) = "VB+" And Arr1(0, i - 7, 7) = "VT+" Then 'VG+VB+VT+ 
                        Arr1(1, i - 7, 10) = "cv1_" '非SITWA  CV
                    ElseIf Arr1(0, i - 7, 4) = "AOV+" Then
                        Arr1(1, i - 7, 10) = "cv1_" '非SITWA  CV
                    Else
                        flagNum(i - 7) = "1"
                        xlSheet.Cells(i, 194).value += "There is no need CV!!!"
                        xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
                    End If
                Else
                    flagNum(i - 7) = "1"
                    xlSheet.Cells(i, 194).value += "The CV quantity can't match SITWA!!!"
                    xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
                End If
            End If

            'AOV判断
            If Arr1(4, i - 7, 8) = "0" Then
                Arr1(0, i - 7, 4) = ""
            ElseIf Arr1(2, i - 7, 0) = "1" Then    '单作用气缸 
                If Arr1(4, i - 7, 8) = "1" Then
                    'AOV只有1个的情况
                    If InStr(LCase(CStr(xlSheet.Cells(i, 75).value)), "lock") And Arr1(0, i - 7, 3) = "SOV+" Then
                        flagNum(i - 7) = "1"
                        Arr1(0, i - 7, 4) = "AOV+"   'SITWA EXCEL 查找序列号
                        Arr1(1, i - 7, 4) = "aov1_"  '非SITWA
                        xlSheet.Cells(i, 194).value += "AOV need 2,please confirm function!!!"
                        xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色                   
                    End If
                ElseIf Arr1(4, i - 7, 8) = "2" Then
                    'AOV只有2个的情况
                    If InStr(LCase(CStr(xlSheet.Cells(i, 75).value)), "lock") And Arr1(0, i - 7, 3) = "SOV+" Then
                        Arr1(0, i - 7, 4) = "AOV+"   'SITWA EXCEL 查找序列号
                        Arr1(1, i - 7, 4) = "aov2_"  '非SITWA
                    Else
                        flagNum(i - 7) = "1"
                        xlSheet.Cells(i, 194).value += "AOV quantity can't match SITWA!!!"
                        xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
                    End If
                Else
                    flagNum(i - 7) = "1"
                    xlSheet.Cells(i, 194).value += "AOV quantity can't match SITWA!!!"
                    xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
                End If
            ElseIf Arr1(2, i - 7, 0) = "2" Then    '双作用气缸                
                If Arr1(4, i - 7, 8) = "2" Then
                    'AOV为2个的情况
                    If InStr(LCase(CStr(xlSheet.Cells(i, 75).value)), "lock") And Arr1(0, i - 7, 1) = "VB+" And Arr1(0, i - 7, 3) = "SOV+" Then
                        flagNum(i - 7) = "1"
                        xlSheet.Cells(i, 194).value += "AOV need 4,please confirm function!!!"
                        xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色                   
                    End If
                ElseIf Arr1(4, i - 7, 8) = "4" Then
                    'AOV为4个的情况
                    If InStr(LCase(CStr(xlSheet.Cells(i, 75).value)), "lock") And Arr1(0, i - 7, 1) = "VB+" And Arr1(0, i - 7, 3) = "SOV+" Then
                        Arr1(0, i - 7, 4) = "AOV+"   'SITWA EXCEL 查找序列号
                        Arr1(1, i - 7, 4) = "aov4_"  '非SITWA
                    Else
                        flagNum(i - 7) = "1"
                        xlSheet.Cells(i, 194).value += "AOV quantity can't match SITWA!!!"
                        xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
                    End If
                Else
                    flagNum(i - 7) = "1"
                    xlSheet.Cells(i, 194).value += "AOV quantity can't match SITWA!!!"
                    xlSheet.Cells(i, 194).Interior.ColorIndex = 27  '设置单元格背景颜色
                End If
            End If

        Next

        Return flagNum

    End Function
    Public Function getArray(ByVal a)
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '获取每条line的所需配置信息
        Dim b = a - 7
        Dim Arr1(6, b, 13) As String '定义组件数组长度  SG+VG+POS+AXIOM+AXIOM+1VB+2QEV+3SOV+4AOV+5PPS+6AFR+7VT
        '0表示sitwa，1表示非sitwa，2表示sitwa表格数值

        For i = 7 To a
            Arr1(3, i - 7, 0) = 0  'flag初始化，定义单作用气缸的失电方式，0表示初始化，1表示失电关，2表示失电开,3表示双作用气缸
            xlSheet.Cells(i, 194).value = ""  '初始化note内容
            Arr1(4, i - 7, 0) = "0"  '0表示正常， 1表示配置仪表数量有问题

            '判断气缸
            If InStr(CStr(xlSheet.Cells(i, 22).value), "B1C") Then
                '判断BC双作用气缸
                Arr1(1, i - 7, 0) = "bc_"  '非SITWA
                Arr1(2, i - 7, 0) = "2"    'SITWA
                Arr1(3, i - 7, 0) = 3      'flag=3
            ElseIf InStr(CStr(xlSheet.Cells(i, 22).value), "VPVL") Then
                If InStr(CStr(xlSheet.Cells(i, 22).value), "DA") Then
                    '判断VPVL双作用气缸
                    Arr1(1, i - 7, 0) = "bc_" '非SITWA
                    Arr1(2, i - 7, 0) = "2"   'SITWA
                    Arr1(3, i - 7, 0) = 3      'flag=3
                ElseIf InStr(CStr(xlSheet.Cells(i, 22).value), "FO") Then
                    '判断VPVL spring to open气缸
                    Arr1(1, i - 7, 0) = "bja_" '非SITWA
                    Arr1(2, i - 7, 0) = "1"    'SITWA
                    Arr1(3, i - 7, 0) = 2      'flag=2
                Else
                    '判断VPVL spring to close气缸
                    Arr1(1, i - 7, 0) = "bj_" '非SITWA
                    Arr1(2, i - 7, 0) = "1"   'SITWA
                    Arr1(3, i - 7, 0) = 1     'flag=1
                End If
            ElseIf InStr(CStr(xlSheet.Cells(i, 22).value), "B1J") Then
                '判断BJ spring to open气缸
                If InStr(CStr(xlSheet.Cells(i, 22).value), "B1JA") Or InStr(CStr(xlSheet.Cells(i, 22).value), "B1JVA") Or InStr(CStr(xlSheet.Cells(i, 22).value), "B1JKA") Then
                    Arr1(1, i - 7, 0) = "bja_" '非SITWA 
                    Arr1(2, i - 7, 0) = "1"    'SITWA
                    Arr1(3, i - 7, 0) = 2      'flag=2
                Else
                    '判断BJ spring to close气缸
                    Arr1(1, i - 7, 0) = "bj_" '非SITWA 
                    Arr1(2, i - 7, 0) = "1"   'SITWA
                    Arr1(3, i - 7, 0) = 1     'flag=1
                End If
            ElseIf InStr(CStr(xlSheet.Cells(i, 22).value), "QPX") Then
                If InStr(CStr(xlSheet.Cells(i, 22).value), "SO") Then
                    '判断QPX spring to open气缸
                    Arr1(1, i - 7, 0) = "bja_" '非SITWA 
                    Arr1(2, i - 7, 0) = "1"    'SITWA
                    Arr1(3, i - 7, 0) = 2      'flag=2
                Else
                    '判断QPX spring to CLOSE气缸
                    Arr1(1, i - 7, 0) = "bj_" '非SITWA 
                    Arr1(2, i - 7, 0) = "1"   'SITWA
                    Arr1(3, i - 7, 0) = 1     'flag=1
                End If
            End If

            '判断失电方向,大写转小写 lcase
            If InStr(LCase(CStr(xlSheet.Cells(i, 50).value)), "close") Then
                Arr1(2, i - 7, 1) = "1"  'SITWA
                Arr1(0, i - 7, 8) = "close"  'PPS失气状态
            ElseIf InStr(LCase(CStr(xlSheet.Cells(i, 50).value)), "open") Then
                Arr1(2, i - 7, 1) = "2"  'SITWA
                Arr1(0, i - 7, 8) = "open"  'PPS失气状态
            Else
                Arr1(2, i - 7, 1) = "3"  'SITWA
            End If

            'PPS失气状态
            If Arr1(2, i - 7, 0) = "2" Then  '双作用气缸
                If Arr1(0, i - 7, 5) = "" And Arr1(0, i - 7, 7) = "" Then
                    Arr1(0, i - 7, 8) = "last"
                End If
            End If
            '带PPS
            If InStr(LCase(CStr(xlSheet.Cells(i, 75).value)), "lock") Then
                Arr1(0, i - 7, 8) = "lock"
            ElseIf InStr(LCase(CStr(xlSheet.Cells(i, 75).value)), "open") Then
                Arr1(0, i - 7, 8) = "open"
            ElseIf InStr(LCase(CStr(xlSheet.Cells(i, 75).value)), "close") Then
                Arr1(0, i - 7, 8) = "close"
            End If

            '判断Controlling device
            '非SITWA 定位器和限位开关筛选
            If CStr(xlSheet.Cells(i, 27).value) = "" Or CStr(xlSheet.Cells(i, 27).value) = Chr(10) Or InStr(LCase(CStr(xlSheet.Cells(i, 26).value)), "delete") Then
                Arr1(1, i - 7, 1) = ""  '非SITWA
            Else
                Arr1(1, i - 7, 1) = "pos_"  '非SITWA
            End If
            If CStr(xlSheet.Cells(i, 32).value) = "" Or CStr(xlSheet.Cells(i, 32).value) = Chr(10) Or InStr(LCase(CStr(xlSheet.Cells(i, 31).value)), "delete") Then
                Arr1(1, i - 7, 2) = ""  '非SITWA
            Else
                Arr1(1, i - 7, 2) = "ls_"  '非SITWA
            End If
            'SITWA 定位器和限位开关筛选            
            If (CStr(xlSheet.Cells(i, 27).value) = "") Or (CStr(xlSheet.Cells(i, 27).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 26).value)), "delete") Then
                '如果pos idcode 写了delete 或者为空，则判断限位开关和电磁阀
                If (CStr(xlSheet.Cells(i, 32).value) = "") Or (CStr(xlSheet.Cells(i, 32).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 31).value)), "delete") Then
                    '如果限位开关也写了delete或者为空，则只判断电磁阀
                    If (CStr(xlSheet.Cells(i, 47).value) <> "") And (CStr(xlSheet.Cells(i, 47).value) <> Chr(10)) Then
                        Arr1(2, i - 7, 2) = "SV"  'SITWA
                        Arr1(0, i - 7, 0) = ""    'SITWA EXCEL 查找序列号
                    Else
                        Arr1(2, i - 7, 2) = "SV抛异常"  'SITWA
                        Arr1(0, i - 7, 0) = ""         'SITWA EXCEL 查找序列号
                        MsgBox("Line " & (i - 6) & " not included in SITWA.")
                    End If
                    '如果限位开关没写delete，判断限位开关和电磁阀
                Else
                    If InStr(CStr(xlSheet.Cells(i, 32).value), "AX") Then
                        Arr1(0, i - 7, 0) = "AXIOM+"  'SITWA EXCEL 查找序列号
                        Arr1(2, i - 7, 2) = "AX"      'SITWA
                    ElseIf InStr(CStr(xlSheet.Cells(i, 32).value), "AM") Then
                        Arr1(0, i - 7, 0) = "AXIOM+"  'SITWA EXCEL 查找序列号
                        Arr1(2, i - 7, 2) = "AX"      'SITWA
                    ElseIf (CStr(xlSheet.Cells(i, 47).value) <> "") And (CStr(xlSheet.Cells(i, 47).value) <> Chr(10)) Then
                        Arr1(0, i - 7, 0) = ""        'SITWA EXCEL 查找序列号
                        Arr1(2, i - 7, 2) = "SV"      'SITWA
                    Else
                        Arr1(0, i - 7, 0) = ""         'SITWA EXCEL 查找序列号
                        Arr1(2, i - 7, 2) = "SV抛异常"  'SITWA
                        MsgBox("Line " & (i - 6) & " not included in SITWA.")
                    End If
                End If
                '如果pos没写delete
            Else
                If InStr(CStr(xlSheet.Cells(i, 27).value), "SG") Then
                    Arr1(0, i - 7, 0) = "SG+"  'SITWA EXCEL 查找序列号
                    Arr1(2, i - 7, 2) = "SG"   'SITWA
                ElseIf InStr(CStr(xlSheet.Cells(i, 27).value), "VG") Then
                    Arr1(0, i - 7, 0) = "VG+"  'SITWA EXCEL 查找序列号
                    Arr1(2, i - 7, 2) = "VG"   'SITWA
                ElseIf InStr(CStr(xlSheet.Cells(i, 27).value), "ND") Then
                    Arr1(0, i - 7, 0) = "POS+"  'SITWA EXCEL 查找序列号
                    Arr1(2, i - 7, 2) = "ND"    'SITWA
                ElseIf InStr(CStr(xlSheet.Cells(i, 27).value), "NE7") Then
                    Arr1(0, i - 7, 0) = "POS+"  'SITWA EXCEL 查找序列号
                    Arr1(2, i - 7, 2) = "ND"    'NE =ND
                ElseIf InStr(CStr(xlSheet.Cells(i, 27).value), "NP7") Then
                    Arr1(0, i - 7, 0) = "POS+"  'SITWA EXCEL 查找序列号
                    Arr1(2, i - 7, 2) = "ND"    'NP =ND
                ElseIf InStr(CStr(xlSheet.Cells(i, 27).value), "6DR") Then
                    Arr1(0, i - 7, 0) = "POS+"  'SITWA EXCEL 查找序列号
                    Arr1(2, i - 7, 2) = "ND"    '6DR=ND
                ElseIf InStr(CStr(xlSheet.Cells(i, 27).value), "DVC") Then
                    Arr1(0, i - 7, 0) = "POS+"  'SITWA EXCEL 查找序列号
                    Arr1(2, i - 7, 2) = "ND"    'DVC=ND   
                ElseIf InStr(CStr(xlSheet.Cells(i, 27).value), "AVP") Then
                    Arr1(0, i - 7, 0) = "POS+"  'SITWA EXCEL 查找序列号
                    Arr1(2, i - 7, 2) = "ND"    'AVP=ND   
                ElseIf InStr(CStr(xlSheet.Cells(i, 32).value), "AX") Or InStr(CStr(xlSheet.Cells(i, 32).value), "AM") Then
                    MsgBox("AXIOM limit switch please move to correct position!")
                Else
                    Arr1(0, i - 7, 0) = "POS+"  'SITWA EXCEL 查找序列号
                    Arr1(2, i - 7, 2) = "ND"    'AVP=ND 
                End If
            End If

            '判断Config of Solenoid valve，暂时将连接方式放于AX列
            If (CStr(xlSheet.Cells(i, 47).value) = "") Or (CStr(xlSheet.Cells(i, 47).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 46).value)), "delete") Then
                '如果电磁阀为空，直接判断为空,电磁阀idcode写delete，判断为空
                Arr1(4, i - 7, 5) = "0"
                Arr1(1, i - 7, 3) = ""   '非SITWA
                Arr1(0, i - 7, 3) = ""   'SITWA EXCEL 查找序列号
                If Arr1(0, i - 7, 0) = "AXIOM+" Then
                    Arr1(2, i - 7, 3) = "1N"   'SITWA
                Else
                    Arr1(2, i - 7, 3) = "NN"  'SITWA

                End If
            ElseIf (CStr(xlSheet.Cells(i, 48).value) = "1" And CStr(xlSheet.Cells(i, 53).value) = "1") Or ((CStr(xlSheet.Cells(i, 48).value) = "" Or CStr(xlSheet.Cells(i, 48).value) = Chr(10)) And (CStr(xlSheet.Cells(i, 53).value) = "2")) Or ((CStr(xlSheet.Cells(i, 53).value) = "" Or CStr(xlSheet.Cells(i, 53).value) = Chr(10)) And (CStr(xlSheet.Cells(i, 48).value) = "2")) Then
                Arr1(4, i - 7, 5) = "2"
                If InStr(LCase(CStr(xlSheet.Cells(i, 50).value)), "1oo2") Then
                    Arr1(0, i - 7, 3) = "SOV+"   'SITWA EXCEL 查找序列号
                    Arr1(1, i - 7, 3) = "sov2s_" '非SITWA 
                    Arr1(2, i - 7, 3) = "2S"     'SITWA
                ElseIf InStr(LCase(CStr(xlSheet.Cells(i, 50).value)), "2oo2") Then
                    Arr1(0, i - 7, 3) = "SOV+"   'SITWA EXCEL 查找序列号
                    Arr1(1, i - 7, 3) = "sov2r_" '非SITWA 
                    Arr1(2, i - 7, 3) = "2R"     'SITWA
                End If
            ElseIf ((CStr(xlSheet.Cells(i, 48).value) = "" Or CStr(xlSheet.Cells(i, 48).value) = Chr(10)) And (CStr(xlSheet.Cells(i, 53).value) = "1")) Or ((CStr(xlSheet.Cells(i, 53).value) = "" Or CStr(xlSheet.Cells(i, 53).value) = Chr(10)) And (CStr(xlSheet.Cells(i, 48).value) = "1")) Then
                Arr1(4, i - 7, 5) = "1"
                Arr1(0, i - 7, 3) = "SOV+" 'SITWA EXCEL 查找序列号
                Arr1(1, i - 7, 3) = "sov_" '非SITWA    
                Arr1(2, i - 7, 3) = "1N"   'SITWA
            Else
                '空值转化为0
                If (CStr(xlSheet.Cells(i, 48).value) = "") Or (CStr(xlSheet.Cells(i, 48).value) = Chr(10)) Then
                    xlSheet.Cells(i, 48).value = 0
                End If
                If (CStr(xlSheet.Cells(i, 53).value) = "") Or (CStr(xlSheet.Cells(i, 53).value) = Chr(10)) Then
                    xlSheet.Cells(i, 53).value = 0
                End If
                Arr1(4, i - 7, 5) = Str(Int(Trim(xlSheet.Cells(i, 48).value)) + Int(Trim(xlSheet.Cells(i, 53).value)))
                Arr1(1, i - 7, 3) = "sov" + Trim(Str(Int(Trim(xlSheet.Cells(i, 48).value)) + Int(Trim(xlSheet.Cells(i, 53).value)))) + "_" '非SITWA 
            End If

            '判断Adjustable speed direction,暂定将方向放于DU列
            If (CStr(xlSheet.Cells(i, 122).value) = "") Or (CStr(xlSheet.Cells(i, 122).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 121).value)), "delete") Then
                'scv 判断空,如果SCV id code写了delete，直接判断为空
                Arr1(1, i - 7, 11) = ""   '非SITWA  
                Arr1(2, i - 7, 4) = "No"  'SITWA
                Arr1(4, i - 7, 4) = "0"
            ElseIf CStr(xlSheet.Cells(i, 123).value) = "2" Then
                Arr1(1, i - 7, 11) = "scvco_"     '非SITWA 
                Arr1(2, i - 7, 4) = "Close+Open"  'SITWA
                Arr1(4, i - 7, 4) = "2"
            ElseIf CStr(xlSheet.Cells(i, 123).value) = "1" Then
                Arr1(4, i - 7, 4) = "1"
                If LCase(CStr(xlSheet.Cells(i, 125).value)) = "close" Then
                    Arr1(1, i - 7, 11) = "scvc_" '非SITWA 
                    Arr1(2, i - 7, 4) = "Close"  'SITWA
                Else
                    Arr1(1, i - 7, 11) = "scvo_" '非SITWA 
                    Arr1(2, i - 7, 4) = "Open"   'SITWA
                End If
            Else
                Arr1(4, i - 7, 4) = xlSheet.Cells(i, 123).value
                Arr1(1, i - 7, 11) = "scv" + Trim(xlSheet.Cells(i, 123).value) + "_"  '非SITWA
            End If

            'VB 判断
            If (CStr(xlSheet.Cells(i, 97).value) = "") Or (CStr(xlSheet.Cells(i, 97).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 96).value)), "delete") Then
                'VB判断空,VB id code 写了delete,直接判断为空
                Arr1(0, i - 7, 1) = ""  'SITWA EXCEL 查找序列号
                Arr1(1, i - 7, 7) = ""  '非SITWA   
                Arr1(4, i - 7, 6) = "0"
            Else
                Arr1(0, i - 7, 1) = "VB+"  'SITWA EXCEL 查找序列号                   
                Arr1(1, i - 7, 7) = "vb" + Trim(xlSheet.Cells(i, 98).value) + "_"  '非SITWA
                Arr1(4, i - 7, 6) = xlSheet.Cells(i, 98).value
            End If

            'QEV判断
            If (CStr(xlSheet.Cells(i, 102).value) = "") Or (CStr(xlSheet.Cells(i, 102).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 101).value)), "delete") Then
                'QEV 判断空,QEV id code 写了delete,直接判断为空
                Arr1(0, i - 7, 2) = ""  'SITWA EXCEL 查找序列号
                Arr1(1, i - 7, 8) = ""  '非SITWA    
                Arr1(4, i - 7, 2) = "0"
            Else
                Arr1(0, i - 7, 2) = "QEV+"  'SITWA EXCEL 查找序列号
                Arr1(1, i - 7, 8) = "qev" + Trim(xlSheet.Cells(i, 103).value) + "_" '非SITWA 
                Arr1(4, i - 7, 2) = xlSheet.Cells(i, 103).value
            End If

            'NV判断
            If (CStr(xlSheet.Cells(i, 107).value) = "") Or (CStr(xlSheet.Cells(i, 107).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 106).value)), "delete") Then
                'NV判断空,NV id code 写了delete,直接判断为空               
                Arr1(1, i - 7, 9) = "" '非SITWA      
                Arr1(4, i - 7, 3) = "0"
            Else
                Arr1(1, i - 7, 9) = "bv" + Trim(xlSheet.Cells(i, 108).value) + "_" '非SITWA
                Arr1(4, i - 7, 3) = xlSheet.Cells(i, 108).value
            End If

            'AOV判断
            If ((CStr(xlSheet.Cells(i, 61).value) = "") Or (CStr(xlSheet.Cells(i, 61).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 61).value)), "delete")) And ((CStr(xlSheet.Cells(i, 66).value) = "") Or (CStr(xlSheet.Cells(i, 66).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 66).value)), "delete")) Then
                'AOV判断空,AOV id code 写了delete,直接判断为空               
                Arr1(0, i - 7, 4) = ""  'SITWA EXCEL 查找序列号
                Arr1(1, i - 7, 4) = ""  '非SITWA    
                Arr1(4, i - 7, 8) = "0"
            Else
                '空值转化为0
                If (CStr(xlSheet.Cells(i, 63).value) = "") Or (CStr(xlSheet.Cells(i, 63).value) = Chr(10)) Then
                    xlSheet.Cells(i, 63).value = 0
                End If
                If (CStr(xlSheet.Cells(i, 68).value) = "") Or (CStr(xlSheet.Cells(i, 68).value) = Chr(10)) Then
                    xlSheet.Cells(i, 68).value = 0
                End If
                Arr1(0, i - 7, 4) = "AOV+"   'SITWA EXCEL 查找序列号               
                Arr1(4, i - 7, 8) = Trim(Str(Int(Trim(xlSheet.Cells(i, 63).value)) + Int(Trim(xlSheet.Cells(i, 68).value))))
                Arr1(1, i - 7, 4) = "aov" + Trim(Str(Int(Trim(xlSheet.Cells(i, 63).value)) + Int(Trim(xlSheet.Cells(i, 68).value)))) + "_"  '非SITWA               
            End If

            'PPS判断
            If (CStr(xlSheet.Cells(i, 72).value) = "") Or (CStr(xlSheet.Cells(i, 72).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 71).value)), "delete") Then
                'PPS判断空,PPS id code 写了delete,直接判断为空
                Arr1(0, i - 7, 5) = ""  'SITWA EXCEL 查找序列号
                Arr1(1, i - 7, 5) = ""  '非SITWA     
                Arr1(4, i - 7, 1) = "0"
            Else
                Arr1(0, i - 7, 5) = "PPS+"  'SITWA EXCEL 查找序列号
                Arr1(1, i - 7, 5) = "pps" + Trim(xlSheet.Cells(i, 73).value) + "_"  '非SITWA
                Arr1(4, i - 7, 1) = xlSheet.Cells(i, 73).value
            End If

            'AFR判断
            If (CStr(xlSheet.Cells(i, 77).value) = "") Or (CStr(xlSheet.Cells(i, 77).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 76).value)), "delete") Then
                'AFR判断空,AFR id code 写了delete,直接判断为空
                Arr1(0, i - 7, 6) = ""  'SITWA EXCEL 查找序列号
                Arr1(1, i - 7, 6) = ""  '非SITWA   
                Arr1(4, i - 7, 0) = "0"
            Else
                Arr1(0, i - 7, 6) = "AFR+"  'SITWA EXCEL 查找序列号               
                Arr1(1, i - 7, 6) = "afr" + Trim(xlSheet.Cells(i, 78).value) + "_"  '非SITWA                
                Arr1(4, i - 7, 0) = xlSheet.Cells(i, 78).value
            End If

            'CV判断 
            If ((CStr(xlSheet.Cells(i, 111).value) = "") Or (CStr(xlSheet.Cells(i, 111).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 111).value)), "delete")) And ((CStr(xlSheet.Cells(i, 116).value) = "") Or (CStr(xlSheet.Cells(i, 116).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 116).value)), "delete")) Then
                'CV id code 写了delete,直接判断为空 
                Arr1(4, i - 7, 7) = "0"
            Else
                '空值转化为0
                If (CStr(xlSheet.Cells(i, 113).value) = "") Or (CStr(xlSheet.Cells(i, 113).value) = Chr(10)) Then
                    xlSheet.Cells(i, 113).value = 0
                End If
                If (CStr(xlSheet.Cells(i, 118).value) = "") Or (CStr(xlSheet.Cells(i, 118).value) = Chr(10)) Then
                    xlSheet.Cells(i, 118).value = 0
                End If
                Arr1(4, i - 7, 7) = Int(Trim(xlSheet.Cells(i, 113).value)) + Int(Trim(xlSheet.Cells(i, 118).value))
                Arr1(1, i - 7, 10) = "cv" + Trim(Int(Trim(xlSheet.Cells(i, 113).value)) + Int(Trim(xlSheet.Cells(i, 118).value))) + "_"  '非SITWA  CV               
            End If

            'VT判断  
            Arr1(0, i - 7, 7) = ""   'SITWA VT   初始化      
            Arr1(1, i - 7, 12) = ""  '非SITWA  VT  初始化
            If Arr1(2, i - 7, 0) = "1" Then    '单作用气缸
                Arr1(0, i - 7, 7) = ""   'SITWA VT              
                Arr1(1, i - 7, 12) = ""  '非SITWA  VT                    
            ElseIf Arr1(2, i - 7, 0) = "2" Then  '双作用气缸
                If InStr(LCase(CStr(xlSheet.Cells(i, 75).value)), "lock") Then
                    Arr1(0, i - 7, 7) = ""   'SITWA VT              
                    Arr1(1, i - 7, 12) = ""  '非SITWA  VT
                ElseIf Arr1(0, i - 7, 4) = "AOV+" And Arr1(0, i - 7, 5) = "PPS+" Then
                    Arr1(0, i - 7, 7) = "VT+"  'SITWA EXCEL 查找序列号         
                    Arr1(1, i - 7, 12) = "vt_" '非SITWA  VT
                ElseIf Arr1(0, i - 7, 0) = "VG+" And Arr1(0, i - 7, 1) = "VB+" And Arr1(4, i - 7, 7) = "1" Then
                    Arr1(0, i - 7, 7) = "VT+"  'SITWA EXCEL 查找序列号         
                    Arr1(1, i - 7, 12) = "vt_" '非SITWA  VT
                End If
            End If

            '非SITWA 结构

            Arr1(1, i - 7, 13) = Arr1(1, i - 7, 0) + Arr1(1, i - 7, 1) + Arr1(1, i - 7, 2) + Arr1(1, i - 7, 3) + Arr1(1, i - 7, 4) + Arr1(1, i - 7, 5) + Arr1(1, i - 7, 6) + Arr1(1, i - 7, 7) + Arr1(1, i - 7, 8) + Arr1(1, i - 7, 9) + Arr1(1, i - 7, 10) + Arr1(1, i - 7, 11) + Arr1(1, i - 7, 12)
            'SITWA 结构
            Arr1(2, i - 7, 5) = Arr1(0, i - 7, 0) + Arr1(0, i - 7, 1) + Arr1(0, i - 7, 3) + Arr1(0, i - 7, 2) + Arr1(0, i - 7, 4) + Arr1(0, i - 7, 5) + Arr1(0, i - 7, 6) + Arr1(0, i - 7, 7)

        Next
        Return Arr1
    End Function
    Public Function saveExcel(fileAddress)
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        Dim xlWorkbook As Excel.Workbook
        xlWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook
        '保存至 fileAddress
        xlWorkbook.SaveAs(fileAddress)
    End Function
    Public Function moveExcel(path, pathFinal)
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        Dim xlWorkbook As Excel.Workbook
        xlWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook
        xlSheet = xlWorkbook.ActiveSheet
        '习惯打开excel再move， Move file
        xlWorkbook.Close(True)
        xlWorkbook = Nothing
        'xlApp.Quit()
        GC.Collect()
        Try
            If File.Exists(path) = False Then
                Dim fs As FileStream = File.Create(path)
                fs.Close()
            End If
            If File.Exists(pathFinal) Then
                File.Delete(pathFinal)
            End If
            File.Move(path, pathFinal)
        Catch ex As Exception
        End Try
    End Function
    Public Function openFile(path)
        If System.IO.File.Exists(path) = False Then
            MsgBox("The CPQ excel doesn't exist in WTC file!")
        ElseIf System.IO.File.Exists(path) Then
            Dim xlApp As Excel.Application      '定义 Excel 程序
            Dim xlBook As Excel.Workbook      '定义 Excel 工作簿
            Dim xlSheet As Excel.Worksheet    '定义 Excel 工作表

            '3、进行Excel操作
            xlApp = CreateObject("Excel.Application") '创建EXCEL对象
            xlBook = xlApp.Workbooks.Open(path) '打开已经存在的EXCEL工件簿文件       
            xlApp.Visible = True 'Excel的可见性
            xlSheet = xlBook.Worksheets(1) '设置活动工作表 表名可用 1\2\3\4代替
        End If
    End Function
    Public Function findDrawing(Arr1, a, flagNum)
        Dim addr As String = "https://metso.sharepoint.com/sites/shanghai_coe_work_and_knowhow/SHA%20Engineered%20Products/Lists/Standard%20Instrumentation%20Functions/Attachments/"
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim Con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\TSHAIPS96-31.mdir.co\SHA_Engineering\Instrumentation\CPQExcelView\database\SITWAexcel.accdb")
        Con.Open()

        For i = 7 To a
            xlSheet.Cells(i, 195).value = Arr1(1, i - 7, 13)
            xlSheet.Cells(i, 191).value = "No" '初始化SITWA
            xlSheet.Cells(i, 192).value = "No" '初始化SITWA 不带AFR
            xlSheet.Cells(i, 193).value = "No"  '初始化非SITWA

            '查找SITWA图纸或只缺少AFR的SITWA图纸

            If Arr1(2, i - 7, 1) = "3" Then
                '缺少失电方向，直接退出搜索图纸
                xlSheet.Cells(i, 191).value = ""
                Exit For
            End If

            If flagNum(i - 7) = "1" And Arr1(0, i - 7, 6) = "AFR+" Then
                '如果不满足判断条件，直接放弃搜索SIWTWA excel
                xlSheet.Cells(i, 191).value = "No"
                xlSheet.Cells(i, 192).value = "No"
            ElseIf flagNum(i - 7) = "0" Then
                '调试
                'MsgBox("1")
                'Dim strtiaoshi As String
                'strtiaoshi = Arr1(2, i - 7, 5) & "_" & Arr1(2, i - 7, 4) & "_" & Arr1(2, i - 7, 3) & "_" & Arr1(2, i - 7, 2) _
                '& "_" & Int(Arr1(2, i - 7, 1)) & "_" & Int(Arr1(2, i - 7, 0)) & "_" & Arr1(0, i - 7, 8)
                'MsgBox(strtiaoshi)
                '有AFR搜索SITWA  excel
                '初始化OLEDB命令的语句 就是查询 什么字段从什么表 条件是ID等于你在t1中输入的内容 xlSheet.Cells(i, 192).value
                Dim sql As String = "Select Field2,Field1 from Sheet1 where Field9 = '" & Arr1(2, i - 7, 5) _
                & "' And Field21 = '" & Arr1(2, i - 7, 4) & "' And Field6 = '" & Arr1(2, i - 7, 3) _
                & "' And Field5 = '" & Arr1(2, i - 7, 2) & "' And Field4 = " & Int(Arr1(2, i - 7, 1)) _
                & " And Field3 = " & Int(Arr1(2, i - 7, 0)) & " And Field15 = '" & Arr1(0, i - 7, 8) & "'"
                Dim strData As String
                strData = String.Empty
                Dim objCommand As New OleDbCommand(sql, Con)
                Dim objReader As OleDbDataReader
                objReader = objCommand.ExecuteReader()
                While objReader.Read()
                    For intindex As Integer = 0 To objReader.FieldCount - 1
                        strData &= objReader.Item(intindex).ToString
                    Next
                End While
                'MsgBox(strData)
                If strData <> "" Then
                    Dim pdfname() As String
                    pdfname = Split(strData, "/")
                    xlSheet.Cells(i, 191).value = "=HYPERLINK(" + Chr(34) + addr _
                    + strData + ".pdf" + Chr(34) + "," + Chr(34) + pdfname(1) + Chr(34) + ")"
                End If
            ElseIf Arr1(0, i - 7, 6) = "" Then
                '查找SITWA配置只缺少AFR的图纸
                Dim strsitwa As String
                'Sitwa 号加AFR
                strsitwa = Arr1(0, i - 7, 0) + Arr1(0, i - 7, 1) + Arr1(0, i - 7, 3) + Arr1(0, i - 7, 2) + Arr1(0, i - 7, 4) + Arr1(0, i - 7, 5) + "AFR+" + Arr1(0, i - 7, 7)
                'MsgBox(strsitwa)
                Dim sql As String = "Select Field2,Field1 from Sheet1 where Field9 = '" & strsitwa _
                & "' And Field21 = '" & Arr1(2, i - 7, 4) & "' And Field6 = '" & Arr1(2, i - 7, 3) _
                & "' And Field5 = '" & Arr1(2, i - 7, 2) & "' And Field4 = " & Int(Arr1(2, i - 7, 1)) _
                & " And Field3 = " & Int(Arr1(2, i - 7, 0)) & " And Field15 = '" & Arr1(0, i - 7, 8) & "'"
                Dim strData As String
                strData = String.Empty
                Dim objCommand As New OleDbCommand(sql, Con)
                Dim objReader As OleDbDataReader
                objReader = objCommand.ExecuteReader()
                While objReader.Read()
                    For intindex As Integer = 0 To objReader.FieldCount - 1
                        strData &= objReader.Item(intindex).ToString
                    Next
                End While
                If strData <> "" Then
                    Dim pdfname() As String
                    pdfname = Split(strData, "/")
                    xlSheet.Cells(i, 192).value = "=HYPERLINK(" + Chr(34) + addr _
                    + strData + ".pdf" + Chr(34) + "," + Chr(34) + pdfname(1) + Chr(34) + ")"
                End If
            End If
            If xlSheet.Cells(i, 191).value = "No" And xlSheet.Cells(i, 192).value = "No" Then
                Dim filename As String
                filename = "\\TSHAIPS96-31.mdir.co\SHA_Engineering\Instrumentation\CPQExcelView\nonSITWAdrawing\PDF\" + Arr1(1, i - 7, 13) + ".pdf"
                'If System.IO.File.Exists(filename) Then
                If Dir(filename) <> "" Then
                    xlSheet.Cells(i, 193).value = "=HYPERLINK(" + Chr(34) + filename + Chr(34) + "," + Chr(34) + Arr1(1, i - 7, 13) + Chr(34) + ")"
                Else
                    xlSheet.Cells(i, 193).value = Arr1(1, i - 7, 13) + " doesn't exist!"
                End If
            End If
                '调试语句
                'MsgBox("FLAGnUM = ", flagNum(i - 7))
                'MsgBox(Arr1(2, i - 7, 5))
                If xlSheet.Cells(i, 191).value = "No" Then
                xlSheet.Cells(i, 191).value = ""
            End If
            If xlSheet.Cells(i, 192).value = "No" Then
                xlSheet.Cells(i, 192).value = ""
            End If
            If xlSheet.Cells(i, 193).value = "No" Then
                xlSheet.Cells(i, 193).value = ""
            End If

        Next
        Con.Close()
        'MsgBox(flagNum(0))

    End Function
    Public Function getArray1(ByVal a)
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '获取每条line的所需配置信息
        Dim Arr1(a - 7, 7, 1) As String '定义组件数组长度

        For i = 7 To a
            'SOV
            If (CStr(xlSheet.Cells(i, 46).value) = "") Or (CStr(xlSheet.Cells(i, 46).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 46).value)), "delete") Then
                Arr1(i - 7, 0, 0) = ""
                Arr1(i - 7, 0, 1) = 0
            Else
                Arr1(i - 7, 0, 0) = xlSheet.Cells(i, 46).value
                Arr1(i - 7, 0, 1) = Int(Trim(xlSheet.Cells(i, 48).value))
            End If

            If (CStr(xlSheet.Cells(i, 51).value) = "") Or (CStr(xlSheet.Cells(i, 51).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 51).value)), "delete") Then
                Arr1(i - 7, 1, 0) = ""
                Arr1(i - 7, 1, 1) = 0
            Else
                Arr1(i - 7, 1, 0) = xlSheet.Cells(i, 51).value
                Arr1(i - 7, 1, 1) = Int(Trim(xlSheet.Cells(i, 53).value))
            End If

            'AOV
            If (CStr(xlSheet.Cells(i, 61).value) = "") Or (CStr(xlSheet.Cells(i, 61).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 61).value)), "delete") Then
                Arr1(i - 7, 2, 0) = ""
                Arr1(i - 7, 2, 1) = 0
            Else
                Arr1(i - 7, 2, 0) = xlSheet.Cells(i, 61).value
                Arr1(i - 7, 2, 1) = Int(Trim(xlSheet.Cells(i, 63).value))
            End If

            If (CStr(xlSheet.Cells(i, 66).value) = "") Or (CStr(xlSheet.Cells(i, 66).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 66).value)), "delete") Then
                Arr1(i - 7, 3, 0) = ""
                Arr1(i - 7, 3, 1) = 0
            Else
                Arr1(i - 7, 3, 0) = xlSheet.Cells(i, 66).value
                Arr1(i - 7, 3, 1) = Int(Trim(xlSheet.Cells(i, 68).value))
            End If

            'VB
            If (CStr(xlSheet.Cells(i, 96).value) = "") Or (CStr(xlSheet.Cells(i, 96).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 96).value)), "delete") Then
                Arr1(i - 7, 4, 0) = ""
                Arr1(i - 7, 4, 1) = 0
            Else
                Arr1(i - 7, 4, 0) = xlSheet.Cells(i, 96).value
                Arr1(i - 7, 4, 1) = Int(Trim(xlSheet.Cells(i, 98).value))
            End If

            'AFR
            If (CStr(xlSheet.Cells(i, 76).value) = "") Or (CStr(xlSheet.Cells(i, 76).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 76).value)), "delete") Then
                Arr1(i - 7, 5, 0) = ""
                Arr1(i - 7, 5, 1) = 0
            Else
                Arr1(i - 7, 5, 0) = xlSheet.Cells(i, 76).value
                Arr1(i - 7, 5, 1) = Int(Trim(xlSheet.Cells(i, 78).value))
            End If

            'PPS
            If (CStr(xlSheet.Cells(i, 71).value) = "") Or (CStr(xlSheet.Cells(i, 71).value) = Chr(10)) Or InStr(LCase(CStr(xlSheet.Cells(i, 71).value)), "delete") Then
                Arr1(i - 7, 6, 0) = ""
                Arr1(i - 7, 6, 1) = 0
            Else
                Arr1(i - 7, 6, 0) = xlSheet.Cells(i, 71).value
                Arr1(i - 7, 6, 1) = Int(Trim(xlSheet.Cells(i, 73).value))
            End If
        Next
        'Dim str1 As String
        'str1 = Arr1(0, 0) + Arr1(0, 1) + Arr1(1, 0) + Arr1(1, 1) + Arr1(2, 0) + Arr1(2, 1) + Arr1(3, 0) + Arr1(3, 1) + Arr1(4, 0) + Arr1(4, 1) + Arr1(5, 0) + Arr1(5, 1) + Arr1(6, 0) + Arr1(6, 1)
        'MsgBox(str1)
        Return Arr1

    End Function
    Public Function codeExist(sqltxt)
        Dim Con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\CPQExcelView\Ins.accdb")
        Con.Open()
        Dim objCommand As New OleDbCommand(sqltxt, Con)
        Dim objReader As OleDbDataReader = objCommand.ExecuteReader()
        Dim strData As String
        If objReader.HasRows Then

            While objReader.Read()
                For intIndex As Integer = 0 To objReader.FieldCount - 1
                    strData &= objReader.Item(intIndex).ToString
                Next
            End While
        End If
        Return strData

    End Function
    Public Function searchdata(sqltxt, insname)
        'MsgBox("1")
        Dim Con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\TSHAIPS96-31.mdir.co\SHA_Engineering\Instrumentation\CPQExcelView\database\Ins.accdb")
        Con.Open()
        'MsgBox("2")
        Dim strData As String
        strData = String.Empty
        Dim objCommand As New OleDbCommand(sqltxt, Con)
        Dim objReader As OleDbDataReader
        objReader = objCommand.ExecuteReader()
        'MsgBox("3")
        While objReader.Read()
            For intindex As Integer = 0 To objReader.FieldCount - 1
                strData &= objReader.Item(intindex).ToString
            Next
        End While
        'MsgBox(strData)
        Dim str1 As String
        str1 = ""
        If strData <> "" Then
            str1 = strData
            'MsgBox("4")
        Else
            Dim msg = insname & " doesn't exist in database! Please add."
            MsgBox(msg)
        End If
        'MsgBox("5")
        Con.Close()
        Return str1
    End Function
    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

        hideSortExcel()

    End Sub
    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        '获得当前激活的sheet
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        For i = 2 To 195
            xlSheet.Columns(i).Hidden = False
        Next
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '获取excel行数
        Dim a = getLineNumber()

        '获取表格搜图纸判断所需的值
        Dim Arr1 = getArray(a)

        '判断气缸失电方向与销售下单选择的方向是否一致
        confEleFunc(Arr1， a)

        Dim flagNum = insQuantityCheck(Arr1, a)


        '查找图纸
        findDrawing(Arr1, a, flagNum)

        For i = 190 To 196
            xlSheet.Columns(i).AutoFit  '设置自适应列宽
        Next
        xlSheet.Columns(13).ColumnWidth = 20        '设置列宽
        xlSheet.Cells(1, 13).WrapText = True
        For i = 191 To 194
            Dim sum = 0
            For j = 7 To CInt(a)
                If (CStr(xlSheet.Cells(j, i).value) = "") Or (CStr(xlSheet.Cells(j, i).value) = Chr(10)) Then
                    sum = sum + 1
                Else
                    Exit For
                End If
            Next
            If sum = a - 6 Then
                xlSheet.Columns(i).Hidden = True
            Else
                xlSheet.Columns(i).Hidden = False
            End If
        Next

        MsgBox("Finish searching！")


    End Sub

    Private Sub Gallery1_Click(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub Gallery1_Click_1(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加positioner
        xlSheet.Columns(26).Hidden = False
        xlSheet.Columns(27).Hidden = False
        xlSheet.Columns(28).Hidden = False
    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加ls
        xlSheet.Columns(31).Hidden = False
        xlSheet.Columns(32).Hidden = False
        xlSheet.Columns(33).Hidden = False
    End Sub

    Private Sub Button7_Click(sender As Object, e As RibbonControlEventArgs) Handles Button7.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加sov1
        xlSheet.Columns(46).Hidden = False
        xlSheet.Columns(47).Hidden = False
        xlSheet.Columns(48).Hidden = False
    End Sub

    Private Sub Button8_Click(sender As Object, e As RibbonControlEventArgs) Handles Button8.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加sov2
        xlSheet.Columns(51).Hidden = False
        xlSheet.Columns(52).Hidden = False
        xlSheet.Columns(53).Hidden = False
    End Sub

    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加fixing plate
        xlSheet.Columns(56).Hidden = False
        xlSheet.Columns(57).Hidden = False
        xlSheet.Columns(58).Hidden = False
    End Sub

    Private Sub Button9_Click(sender As Object, e As RibbonControlEventArgs) Handles Button9.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加aov1
        xlSheet.Columns(61).Hidden = False
        xlSheet.Columns(62).Hidden = False
        xlSheet.Columns(63).Hidden = False
    End Sub

    Private Sub Button10_Click(sender As Object, e As RibbonControlEventArgs) Handles Button10.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加aov2
        xlSheet.Columns(66).Hidden = False
        xlSheet.Columns(67).Hidden = False
        xlSheet.Columns(68).Hidden = False
    End Sub

    Private Sub Button11_Click(sender As Object, e As RibbonControlEventArgs) Handles Button11.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加pps
        xlSheet.Columns(71).Hidden = False
        xlSheet.Columns(72).Hidden = False
        xlSheet.Columns(73).Hidden = False
    End Sub

    Private Sub Button12_Click(sender As Object, e As RibbonControlEventArgs) Handles Button12.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加afr
        xlSheet.Columns(76).Hidden = False
        xlSheet.Columns(77).Hidden = False
        xlSheet.Columns(78).Hidden = False
    End Sub

    Private Sub Button13_Click(sender As Object, e As RibbonControlEventArgs) Handles Button13.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加pg1
        xlSheet.Columns(81).Hidden = False
        xlSheet.Columns(82).Hidden = False
        xlSheet.Columns(83).Hidden = False
    End Sub

    Private Sub Button14_Click(sender As Object, e As RibbonControlEventArgs) Handles Button14.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加pg2
        xlSheet.Columns(86).Hidden = False
        xlSheet.Columns(87).Hidden = False
        xlSheet.Columns(88).Hidden = False
    End Sub

    Private Sub Button15_Click(sender As Object, e As RibbonControlEventArgs) Handles Button15.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加vb
        xlSheet.Columns(96).Hidden = False
        xlSheet.Columns(97).Hidden = False
        xlSheet.Columns(98).Hidden = False
    End Sub

    Private Sub Button19_Click(sender As Object, e As RibbonControlEventArgs) Handles Button19.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加qev
        xlSheet.Columns(101).Hidden = False
        xlSheet.Columns(102).Hidden = False
        xlSheet.Columns(103).Hidden = False
    End Sub

    Private Sub Button18_Click(sender As Object, e As RibbonControlEventArgs) Handles Button18.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加nv
        xlSheet.Columns(106).Hidden = False
        xlSheet.Columns(107).Hidden = False
        xlSheet.Columns(108).Hidden = False
    End Sub

    Private Sub Button20_Click(sender As Object, e As RibbonControlEventArgs) Handles Button20.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加cv1
        xlSheet.Columns(111).Hidden = False
        xlSheet.Columns(112).Hidden = False
        xlSheet.Columns(113).Hidden = False
    End Sub

    Private Sub Button21_Click(sender As Object, e As RibbonControlEventArgs) Handles Button21.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加cv2
        xlSheet.Columns(116).Hidden = False
        xlSheet.Columns(117).Hidden = False
        xlSheet.Columns(118).Hidden = False
    End Sub

    Private Sub Button17_Click(sender As Object, e As RibbonControlEventArgs) Handles Button17.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加scv
        xlSheet.Columns(121).Hidden = False
        xlSheet.Columns(122).Hidden = False
        xlSheet.Columns(123).Hidden = False
    End Sub

    Private Sub Button16_Click(sender As Object, e As RibbonControlEventArgs) Handles Button16.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加silencer
        xlSheet.Columns(126).Hidden = False
        xlSheet.Columns(127).Hidden = False
        xlSheet.Columns(128).Hidden = False
    End Sub

    Private Sub Button22_Click(sender As Object, e As RibbonControlEventArgs) Handles Button22.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加other1
        xlSheet.Columns(146).Hidden = False
        xlSheet.Columns(147).Hidden = False
        xlSheet.Columns(148).Hidden = False
    End Sub

    Private Sub Button23_Click(sender As Object, e As RibbonControlEventArgs) Handles Button23.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        '增加other2
        xlSheet.Columns(151).Hidden = False
        xlSheet.Columns(152).Hidden = False
        xlSheet.Columns(153).Hidden = False
    End Sub

    Private Sub EditBox1_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles EditBox1.TextChanged

    End Sub

    Private Sub Button24_Click(sender As Object, e As RibbonControlEventArgs) Handles Button24.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        Dim xlWorkbook As Excel.Workbook
        xlWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook
        xlSheet = xlWorkbook.ActiveSheet
        '保存至 C:\CPQExcelView\IQI\WTC
        Dim fileAddress = "C:\CPQExcelView\IQI\WTC\" + CStr(xlSheet.Cells(4, 13).Value) + ".xls"
        saveExcel(fileAddress)
    End Sub

    Private Sub Button25_Click(sender As Object, e As RibbonControlEventArgs) Handles Button25.Click
        Dim path = "C:\CPQExcelView\IQI\WTC\" + Trim(EditBox1.Text) + ".xls"
        If EditBox1.Text = "" Then
            MsgBox("Please input the CPQ number!")
        Else
            openFile(path)
        End If
    End Sub

    Private Sub Button27_Click(sender As Object, e As RibbonControlEventArgs) Handles Button27.Click
        Dim path = "C:\CPQExcelView\IQI\Done\" + Trim(EditBox1.Text) + ".xls"
        If EditBox1.Text = "" Then
            MsgBox("Please input the CPQ number!")
        Else
            openFile(Path)
        End If
    End Sub

    Private Sub Button26_Click(sender As Object, e As RibbonControlEventArgs) Handles Button26.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        Dim xlWorkbook As Excel.Workbook
        xlWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook
        xlSheet = xlWorkbook.ActiveSheet
        '习惯打开excel再move， Move file
        Dim path = "C:\CPQExcelView\IQI\WTC\" + CStr(xlSheet.Cells(4, 13).Value) + ".xls"
        Dim pathFinal = "C:\CPQExcelView\IQI\Done\" + CStr(xlSheet.Cells(4, 13).Value) + ".xls"
        moveExcel(Path, pathFinal)
    End Sub

    Private Sub Button28_Click(sender As Object, e As RibbonControlEventArgs) Handles Button28.Click
        '获取当前excel
        Dim xlSheet As Excel.Worksheet
        Dim xlWorkbook As Excel.Workbook
        xlWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook
        xlSheet = xlWorkbook.ActiveSheet
        Dim filename As String '获取文件路径
        If EditBox2.Text <> "" Then
            filename = "T:\Instrumentation\SMartINS\Input_Excel\" + Trim(EditBox2.Text) + ".xlsx"
        Else
            filename = "T:\Instrumentation\SMartINS\Input_Excel\" + CStr(xlSheet.Cells(4, 13).Value) + ".xlsx"
        End If
        xlWorkbook.SaveAs(Filename:=filename, FileFormat:=51)
        xlWorkbook.Close(True)
        xlWorkbook = Nothing
        'xlApp.Quit()
        GC.Collect()
    End Sub

    Private Sub Group7_DialogLauncherClick(sender As Object, e As RibbonControlEventArgs) Handles Group7.DialogLauncherClick
        Dim bb As New Form1
        bb.Show()
    End Sub

    Private Sub Button29_Click(sender As Object, e As RibbonControlEventArgs) Handles Button29.Click
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '获取excel行数
        Dim a = getLineNumber()
        Dim arr1 = getArray1(a)
        Dim accessories(a - 6) As String

        For i = 7 To a
            accessories(i - 7) = "Add "
            'SOV1支架
            Dim str1 As String
            If arr1(i - 7, 0, 1) = 0 Then
                str1 = ""
            Else
                Dim sqltxt As String
                sqltxt = "Select Bracket from SOV where IdCode = '" & arr1(i - 7, 0, 0) & "'"
                str1 = searchdata(sqltxt, arr1(i - 7, 0, 0))
                If str1 = "no need" Or str1 = "included" Then
                    str1 = ""
                ElseIf str1 = "lack" Then
                    str1 = "/lack SOV bracket!/"
                Else
                    str1 &= "*" & arr1(i - 7, 0, 1)
                End If
            End If
                MsgBox(str1)
            'SOV2支架
            Dim str2 As String
            If arr1(i - 7, 1, 1) = 0 Then
                str1 = ""
            Else
                Dim sqltxt As String
                sqltxt = "Select Bracket from SOV where IdCode = '" & arr1(i - 7, 1, 0) & "'"
                str1 = searchdata(sqltxt, arr1(i - 7, 1, 0))
                If str2 = "no need" Or str1 = "included" Then
                    str2 = ""
                ElseIf str2 = "lack" Then
                    str2 = "/lack SOV bracket!/"
                Else
                    str2 &= "*" & arr1(i - 7, 0, 1)
                End If
            End If
            MsgBox(str2)
        Next

    End Sub

    Private Sub Button30_Click(sender As Object, e As RibbonControlEventArgs) Handles Button30.Click
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet

        '获取excel行数
        Dim a = getLineNumber()

        xlSheet.Columns(13).ColumnWidth = 20        '设置列宽
        xlSheet.Columns(196).AutoFit  '设置自适应列宽
        xlSheet.Cells(1, 13).WrapText = True

        For i = 2 To 12
            xlSheet.Columns(i).hidden = True
        Next
        xlSheet.Columns(13).hidden = False
        Dim list(12) As Integer
        list = {14, 15, 191, 192, 193, 194, 195}
        Dim lenlist = UBound(list)
        For i = 0 To lenlist
            xlSheet.Columns(list(i)).hidden = True
        Next

        For i = 16 To 186 Step 5
            Dim sum = 0
            Dim sum1 = 0
            For j = 7 To a
                If InStr(LCase(CStr(xlSheet.Cells(j, 196).Value)), "comment") Then
                    sum1 = sum1 + 1
                End If
            Next
            If sum1 = 0 Then
                xlSheet.Columns(196).Hidden = True
            End If
            For j = 7 To a
                If InStr(LCase(CStr(xlSheet.Cells(j, i + 3).Value)), "comment") Then
                    sum = sum + 1
                End If
            Next
            If sum = 0 Then
                If xlSheet.Columns(i).Hidden = False Then
                    xlSheet.Columns(i).Hidden = True
                    xlSheet.Columns(i + 1).Hidden = True
                    xlSheet.Columns(i + 2).Hidden = True
                    xlSheet.Columns(i + 3).Hidden = True
                End If
            Else
                xlSheet.Columns(i + 3).AutoFit  '设置自适应列宽
            End If

            If xlSheet.Columns(i + 4).Hidden = False Then
                xlSheet.Columns(i + 4).Hidden = True
            End If
        Next

    End Sub

    Private Sub Button34_Click(sender As Object, e As RibbonControlEventArgs) Handles Button34.Click
        Dim xlSheet As Excel.Worksheet
        xlSheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim xlCell = Globals.ThisAddIn.Application.ActiveCell
        Dim cnum = Int(xlCell.Column)
        Dim crow = Int(xlCell.Row)
        '获取excel行数
        Dim a = getLineNumber()
        If cnum < 16 Or crow < 7 Or cnum > 190 Or crow > a Then
            MsgBox("Please choose the valid area!")
        ElseIf (cnum - 15) Mod 5 = 0 Then
            If InStr(LCase(CStr(xlSheet.Cells(crow, cnum - 1).Value)), "comment") Then
                MsgBox("Repeat click!")
            Else
                Dim intResult As Integer
                intResult = MessageBox.Show("Do you want to add comment?", "Warning!", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)
                If intResult = DialogResult.OK Then
                    xlSheet.Cells(crow, cnum - 1).Value = "Comment:  "
                    xlSheet.Cells(crow, cnum - 1).Interior.ColorIndex = 20  '设置单元格背景颜色
                    xlSheet.Cells(crow, cnum - 1).Font.Color = 255
                    xlSheet.Columns(cnum - 1).Hidden = False
                End If
            End If
        ElseIf (cnum - 15) Mod 5 = 1 Then
            If InStr(LCase(CStr(xlSheet.Cells(crow, cnum + 3).Value)), "comment") Then
                MsgBox("Repeat click!")
            Else
                Dim intResult As Integer
                intResult = MessageBox.Show("Do you want to add comment?", "Warning!", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)
                If intResult = DialogResult.OK Then
                    xlSheet.Cells(crow, cnum + 3).Value = "Comment:  "
                    xlSheet.Cells(crow, cnum + 3).Font.Color = 255
                    xlSheet.Cells(crow, cnum + 3).Interior.ColorIndex = 20  '设置单元格背景颜色
                    xlSheet.Columns(cnum + 3).Hidden = False
                End If
            End If
        ElseIf (cnum - 15) Mod 5 = 2 Then
            If InStr(LCase(CStr(xlSheet.Cells(crow, cnum + 2).Value)), "comment") Then
                MsgBox("Repeat click!")
            Else
                Dim intResult As Integer
                intResult = MessageBox.Show("Do you want to add comment?", "Warning!", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)
                If intResult = DialogResult.OK Then
                    xlSheet.Cells(crow, cnum + 2).Value = "Comment:  "
                    xlSheet.Cells(crow, cnum + 2).Interior.ColorIndex = 20  '设置单元格背景颜色
                    xlSheet.Cells(crow, cnum + 2).Font.Color = 255
                    xlSheet.Columns(cnum + 2).Hidden = False
                End If
            End If
        ElseIf (cnum - 15) Mod 5 = 3 Then
            If InStr(LCase(CStr(xlSheet.Cells(crow, cnum + 1).Value)), "comment") Then
                MsgBox("Repeat click!")
            Else
                Dim intResult As Integer
                intResult = MessageBox.Show("Do you want to add comment?", "Warning!", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)
                If intResult = DialogResult.OK Then
                    xlSheet.Cells(crow, cnum + 1).Value = "Comment:  "
                    xlSheet.Cells(crow, cnum + 1).Interior.ColorIndex = 20  '设置单元格背景颜色
                    xlSheet.Cells(crow, cnum + 1).Font.Color = 255
                    xlSheet.Columns(cnum + 1).Hidden = False
                End If
            End If
        ElseIf (cnum - 15) Mod 5 = 4 Then
            If InStr(LCase(CStr(xlSheet.Cells(crow, cnum).Value)), "comment") Then
                MsgBox("Repeat click!")
            Else
                Dim intResult As Integer
            intResult = MessageBox.Show("Do you want to add comment?", "Warning!", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)
                If intResult = DialogResult.OK Then
                    xlSheet.Cells(crow, cnum).Value = "Comment:  "
                    xlSheet.Cells(crow, cnum).Interior.ColorIndex = 20  '设置单元格背景颜色
                    xlSheet.Cells(crow, cnum).Font.Color = 255
                    xlSheet.Columns(cnum).Hidden = False
                End If
            End If
        End If
    End Sub
End Class
