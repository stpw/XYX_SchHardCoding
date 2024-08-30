Imports System.IO

Public Class GetAppConfig

#Region "常用的配置內容設定"
    '數據庫連接字符串
    Public Shared ReadOnly remoteConnect As String = System.Configuration.ConfigurationManager.ConnectionStrings("connstring").ToString
    '數據庫連接字符串
    Public Shared ReadOnly remoteEmailConnect As String = System.Configuration.ConfigurationManager.ConnectionStrings("connstring").ToString
    '郵件服務器地址
    Public Shared ReadOnly emailServerAddress As String = System.Configuration.ConfigurationManager.ConnectionStrings("EmailServerAddress").ToString
    '寄件者
    Public Shared ReadOnly emailFrom As String = System.Configuration.ConfigurationManager.ConnectionStrings("EmailFrom").ToString
    '密本
    Public Shared ReadOnly emailBcc As String = System.Configuration.ConfigurationManager.ConnectionStrings("EmailBCc").ToString
#End Region

#Region "通用方法獲取配置文件的設定值"

    ''' <summary>
    ''' 通用方法獲取配置文件的設定值
    ''' </summary>
    ''' <param name="keyName">key名稱</param>
    ''' <returns>配置文件的設定值</returns>
    ''' <remarks></remarks>
    Public Shared Function GetValue(ByVal keyName As String) As String
        Return System.Configuration.ConfigurationManager.ConnectionStrings(keyName).ToString
    End Function

#End Region

End Class
