Imports System.IO
Imports System.Data

Imports System.Text
Imports System.Security.Cryptography
'Imports System.Web.Security

Imports System.Xml
Module Mod_Function    '系統之公用功能

#Region "加密"
    Public Function Encryption(ByVal str As String) As String
        Return Convert.ToBase64String(System.Text.Encoding.Default.GetBytes(str)).Replace("+", "%2B")
    End Function
#End Region

#Region "解密"
    Public Function Decryption(ByVal str As String) As String
        Return System.Text.Encoding.Default.GetString(Convert.FromBase64String(str.Replace("%2B", "+")))
    End Function
#End Region

    '/ <summary>
    '/ 解密
    '/ </summary>
    '/ <param name="Text"></param>
    '/ <returns></returns>
    'Public Function Decrypt(ByVal Text As String) As String
    '    Return Decrypt(Text, "litianping")
    'End Function

    '/ <summary> 
    '/ 解密数据 
    '/ </summary> 
    '/ <param name="Text"></param> 
    '/ <param name="sKey"></param> 
    '/ <returns></returns> 

    'Public Function Decrypt(ByVal Text As String, ByVal sKey As String) As String
    '    Dim des As New DESCryptoServiceProvider()
    '    Dim len As Integer
    '    len = Text.Length / 2
    '    Dim inputByteArray(len - 1) As Byte
    '    Dim x, i As Integer
    '    For x = 0 To len - 1
    '        i = Convert.ToInt32(Text.Substring(x * 2, 2), 16)
    '        inputByteArray(x) = CType(i, Byte)
    '    Next
    '    des.Key = ASCIIEncoding.ASCII.GetBytes(System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(sKey, "md5").Substring(0, 8))
    '    des.IV = ASCIIEncoding.ASCII.GetBytes(System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(sKey, "md5").Substring(0, 8))
    '    Dim ms As New System.IO.MemoryStream()
    '    Dim cs As New CryptoStream(ms, des.CreateDecryptor(), CryptoStreamMode.Write)
    '    cs.Write(inputByteArray, 0, inputByteArray.Length)
    '    cs.FlushFinalBlock()
    '    Return Encoding.Default.GetString(ms.ToArray())
    'End Function



    Public Function CreateExcel2007(ByVal title As ArrayList, ByVal field As ArrayList, ByVal ds As DataSet) As NPOI.SS.UserModel.IWorkbook
        'Dim rowCount As Integer = 0
        'Dim sheetCount As Integer = 1
        'Dim sheet As NPOI.SS.UserModel.ISheet
        'Dim row As NPOI.SS.UserModel.IRow
        ''创建workbook
        'Dim workbook As NPOI.SS.UserModel.IWorkbook = New NPOI.XSSF.UserModel.XSSFWorkbook()
        ''创建sheet
        'sheet = workbook.CreateSheet("Sheet" & sheetCount)
        ''寫列標題  
        'Dim columncount As Integer = title.Count
        ''创建行
        'row = sheet.CreateRow(0)
        'For columi As Integer = 0 To columncount - 1
        '    '创建单元格
        '    Dim cell As NPOI.SS.UserModel.ICell = row.CreateCell(columi)
        '    '设置单元格值
        '    cell.SetCellValue(title.Item(columi).ToString)
        'Next
        'For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
        '    rowCount = rowCount + 1
        '    If rowCount = 10001 Then
        '        '凍結標題列
        '        sheet.CreateFreezePane(0, 1, 0, 1)
        '        '自動伸縮欄寬
        '        sheet.AutoSizeColumn(2)
        '        rowCount = 1
        '        sheetCount = sheetCount + 1
        '        sheet = workbook.CreateSheet("Sheet" & sheetCount)
        '        '寫列標題  
        '        row = sheet.CreateRow(0)
        '        For columi As Integer = 0 To columncount - 1
        '            '创建单元格
        '            Dim cell As NPOI.SS.UserModel.ICell = row.CreateCell(columi)
        '            '设置单元格值
        '            cell.SetCellValue(title.Item(columi).ToString)
        '        Next
        '    End If
        '    row = sheet.CreateRow(rowCount)
        '    For j As Integer = 0 To field.Count - 1
        '        '创建单元格
        '        Dim cell As NPOI.SS.UserModel.ICell = row.CreateCell(j)
        '        '设置单元格值
        '        cell.SetCellValue(ds.Tables(0).Rows(i)(field.Item(j).ToString).ToString)
        '    Next
        'Next
        ''凍結標題列
        'sheet.CreateFreezePane(0, 1, 0, 1)
        ''自動伸縮欄寬
        'sheet.AutoSizeColumn(2)
        'Return workbook
        '创建workbook
        Dim workbook As NPOI.SS.UserModel.IWorkbook = New NPOI.XSSF.UserModel.XSSFWorkbook()
        '创建sheet
        Dim sheet As NPOI.SS.UserModel.ISheet = workbook.CreateSheet("Sheet1")
        '寫列標題  
        Dim columncount As Integer = title.Count
        '创建行
        Dim row As NPOI.SS.UserModel.IRow = sheet.CreateRow(0)
        For columi As Integer = 0 To columncount - 1
            '创建单元格
            Dim cell As NPOI.SS.UserModel.ICell = row.CreateCell(columi)
            cell.CellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN
            cell.CellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.THIN
            cell.CellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.THIN
            cell.CellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.THIN
            '设置单元格值
            cell.SetCellValue(title.Item(columi).ToString)
        Next

        For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
            '创建行
            row = sheet.CreateRow(i + 1)
            For j As Integer = 0 To title.Count - 1
                '创建单元格
                Dim cell As NPOI.SS.UserModel.ICell = row.CreateCell(j)
                '设置单元格值
                If Not (ds.Tables(0).Rows(i)(field.Item(j).ToString) Is System.DBNull.Value) Then
                    Select Case ds.Tables(0).Columns(field.Item(j).ToString).DataType.Name.ToUpper()
                        Case "Integer".ToUpper, "Int32".ToUpper, "Int16".ToUpper, "Int64".ToUpper
                            cell.SetCellValue(Convert.ToInt32(ds.Tables(0).Rows(i)(field.Item(j).ToString).ToString.Trim()))
                        Case ("Double".ToUpper), "Decimal".ToUpper, "Float".ToUpper
                            cell.SetCellValue(Convert.ToDouble(ds.Tables(0).Rows(i)(field.Item(j).ToString).ToString.Trim()))
                        Case "String".ToUpper
                            cell.SetCellValue(Convert.ToString(ds.Tables(0).Rows(i)(field.Item(j).ToString).ToString.Trim()))
                        Case "DateTime".ToUpper
                            cell.SetCellValue(Convert.ToDateTime(ds.Tables(0).Rows(i)(field.Item(j).ToString()).ToString.Trim()).ToString("yyyy/MM/dd hh:mm:ss"))
                        Case "Boolean".ToUpper
                            cell.SetCellValue(Convert.ToString(ds.Tables(0).Rows(i)(field.Item(j).ToString).ToString.Trim()))
                        Case Else
                            cell.SetCellValue(Convert.ToString(ds.Tables(0).Rows(i)(field.Item(j).ToString).ToString.Trim()))
                    End Select
                End If
            Next
        Next
        Return workbook
    End Function

    Public Function CreateExcel2008(ByVal title As ArrayList, ByVal field As ArrayList, ByVal ds As DataSet) As NPOI.SS.UserModel.IWorkbook
        Dim workbook As NPOI.SS.UserModel.IWorkbook = New NPOI.XSSF.UserModel.XSSFWorkbook()
        '创建sheet
        Dim sheet As NPOI.SS.UserModel.ISheet = workbook.CreateSheet("Sheet1")
        '寫列標題  
        Dim columncount As Integer = title.Count
        '创建行
        Dim row As NPOI.SS.UserModel.IRow = sheet.CreateRow(0)
        For columi As Integer = 0 To columncount - 1
            '创建单元格
            Dim cell As NPOI.SS.UserModel.ICell = row.CreateCell(columi)
            cell.CellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN
            cell.CellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.THIN
            cell.CellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.THIN
            cell.CellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.THIN
            '设置单元格值
            cell.SetCellValue(title.Item(columi).ToString)
        Next

        For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
            '创建行
            row = sheet.CreateRow(i + 1)
            For j As Integer = 0 To title.Count - 1
                '创建单元格
                Dim cell As NPOI.SS.UserModel.ICell = row.CreateCell(j)
                '设置单元格值
                If Not (ds.Tables(0).Rows(i)(field.Item(j).ToString) Is System.DBNull.Value) Then
                    Select Case ds.Tables(0).Columns(field.Item(j).ToString).DataType.Name.ToUpper()
                        Case "Integer".ToUpper, "Int32".ToUpper, "Int16".ToUpper, "Int64".ToUpper
                            cell.SetCellValue(Convert.ToInt32(ds.Tables(0).Rows(i)(field.Item(j).ToString).ToString.Trim()))
                        Case ("Double".ToUpper), "Decimal".ToUpper, "Float".ToUpper
                            cell.SetCellValue(Convert.ToDouble(ds.Tables(0).Rows(i)(field.Item(j).ToString).ToString.Trim()))
                        Case "String".ToUpper
                            cell.SetCellValue(Convert.ToString(ds.Tables(0).Rows(i)(field.Item(j).ToString).ToString.Trim()))
                        Case "DateTime".ToUpper
                            cell.SetCellValue(Convert.ToDateTime(ds.Tables(0).Rows(i)(field.Item(j).ToString()).ToString.Trim()).ToString("yyyy/MM/dd HH:mm:ss"))
                        Case "Boolean".ToUpper
                            cell.SetCellValue(Convert.ToString(ds.Tables(0).Rows(i)(field.Item(j).ToString).ToString.Trim()))
                        Case Else
                            cell.SetCellValue(Convert.ToString(ds.Tables(0).Rows(i)(field.Item(j).ToString).ToString.Trim()))
                    End Select
                End If
            Next
        Next
        Return workbook
    End Function

#Region "將秒數轉成天數(xx天xx小時xx分)"

    ''' <summary>
    ''' 將秒數轉成天數
    ''' </summary>
    ''' <param name="second">秒</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SecToDay(ByVal second As Integer) As String '
        Dim D As Integer
        Dim H As Integer
        Dim M As Integer
        Dim S As Integer
        D = Fix(second / (60 * 60 * 24))
        H = (second / 3600) Mod 24
        M = Fix((second Mod 3600) / 60)
        S = second Mod 60
        SecToDay = D & "天" & H & "小時" & M & "分" ' & S & "秒"    ====> xx天xx小時xx分
    End Function

#End Region

#Region "產生日期時間字串"

    ''' <summary>
    ''' 產生日期時間
    ''' 例如：產生編號 HeadStr  092 08 01 12 45 30 字串
    ''' </summary>
    ''' <param name="HeadStr">標頭字串</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RunID(ByVal HeadStr As String) As String
        Dim IDStr As String
        Dim NowStr As Date
        Dim YearStr As String
        Dim MonthStr As String
        Dim DayStr As String
        Dim HourStr As String
        Dim MinuteStr As String
        Dim SecondStr As String
        NowStr = Now
        YearStr = Format(Val(Year(NowStr) - 1911), "000")
        MonthStr = Format(Month(NowStr), "00")
        DayStr = Format(Microsoft.VisualBasic.Day(NowStr), "00")
        HourStr = Format(Hour(NowStr), "00")
        MinuteStr = Format(Minute(NowStr), "00")
        SecondStr = Format(Second(NowStr), "00")
        IDStr = HeadStr & YearStr & MonthStr & DayStr & HourStr & MinuteStr & SecondStr
        RunID = IDStr
    End Function

#End Region

#Region "檢查Email文字串中是否包含@"

    ''' <summary>
    ''' 檢查Email文字串中是否包含@
    ''' </summary>
    ''' <param name="eMail">電子郵件位址</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function EMailCheck(ByVal eMail As String) As Boolean

        Dim result As Integer
        result = eMail.IndexOf("@")
        If result = -1 Then
            Return False
        End If

        Dim eMailSplit() As String = Split(eMail, "@")
        If eMailSplit(0).Length < 1 Or eMailSplit(eMailSplit.Length - 1).Length < 1 Then
            Return False
        End If

        Return True
    End Function

#End Region

#Region "計算傳入的日期時間與目前的差距為何"

    ''' <summary>
    ''' 計算傳入的日期時間與目前的差距為何
    ''' </summary>
    ''' <param name="SetEndDate">傳入日期時間</param>
    ''' <returns>秒</returns>
    ''' <remarks></remarks>
    Public Function CountSurplusTime(ByVal SetEndDate As Date) As Int64  '計算剩餘時間
        Dim v1 As Integer
        'SetEndDate = #3/20/2004#
        v1 = DateDiff(DateInterval.Second, Now, SetEndDate)
        Return v1

    End Function

#End Region

#Region "於數字前面補上空格功能方法"

    ''' <summary>
    ''' 於數字前面補上空格功能方法
    '''  8
    '''  9
    ''' 10
    ''' </summary>
    ''' <param name="bufNum">傳入數字</param>
    ''' <param name="blankspaceNum">所需控制數字總長度</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function patchBlankSpace(ByVal bufNum As Integer, ByVal blankspaceNum As String)

        Dim bufNumLength As Integer = bufNum.ToString().Length
        Dim resultString As String = ""

        If (bufNumLength >= blankspaceNum) Then

            Return bufNum.ToString()
        Else

            For i As Integer = 0 To blankspaceNum - bufNum.ToString().Length - 1

                resultString = resultString + " "

            Next

            resultString = resultString + bufNum.ToString()
            Return resultString

        End If

    End Function

#End Region

End Module
