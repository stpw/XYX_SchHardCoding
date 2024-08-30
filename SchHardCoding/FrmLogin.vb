Imports System.IO
Imports System.Net
Imports System.Net.Dns
Imports System.Net.NetworkInformation
Imports System.Xml

Public Class FrmLogin
    Dim connectionString As String = ""
    '/ <summary>
    '/ 聲明類
    '/ </summary>
    'Public Shared Autoupdate As ERP.Autoupdate = New ERP.Autoupdate
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    '輸完後,按Enter鍵,可跳到下一輸入點,Form之KeyPreview需設為True
    Private Sub Fm_BA3300_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If (e.KeyValue.ToString = 13) Then
            System.Windows.Forms.SendKeys.Send("{tab}")
        End If
    End Sub
    Private Sub FrmLogin_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        'Autoupdate.checkUpdate() '//检测更新

        'Dim strUpdateXmlPath As String = Application.StartupPath + "\update\conf\update.xml"
        'connection.con = Mod_Function.Decrypt(ERP.Autoupdate.getConfigValue(strUpdateXmlPath, "ConnectionString"))
        connection.con = GetAppConfig.remoteConnect
        Label5.Text = Mid(connection.con, 1, 20)
    End Sub

    Private Sub bt_login_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_login.Click
        Dim sql_txt, SQL As String
        Dim p_count As Integer = 0
        Dim IS_MIDFORM As New Form '窗體
        Dim IS_FRMNAME As String '程式名稱
        Dim Message As String '窗體名稱
        'SQL = "SELECT ComputerID FROM Acme.dbo.CanUserAdmin WHERE ComputerID='" & Environment.MachineName.ToString & "'"
        'If db.FillDataSet(SQL).Tables(0).Rows.Count = 0 Then
        '    If Txt_UserID.Text.Trim.ToUpper = "ADMIN" Then
        '        MsgBox("請不要試圖盜用管理員帳號！")
        '        Exit Sub
        '    End If
        'End If

        sql_txt = " select count(userid)  " & _
                 " from Acme.dbo.tbl_i_users " & _
                  " where userid='" + Txt_UserID.Text + "' " & _
                  "  AND permit=1  "
        p_count = db.GetExecSQL(sql_txt)
        If p_count = 0 Then
            MsgBox("帳號不存在或者人員已離職！")
            Exit Sub
        End If

        sql_txt = " select count(userid)  " & _
                 " from Acme.dbo.tbl_i_users " & _
                  " where userid='" + Txt_UserID.Text + "' " & _
                  "   and password   =substring(sys.fn_VarBinToHexStr(hashbytes('MD5','" + txt_pwd.Text + "')),3,32) "
        p_count = db.GetExecSQL(sql_txt)
        If p_count = 0 Then
            MsgBox("帳號/密碼不對,請重新輸入")
        Else
            '--------------------------------
            sql_txt = " select userid,username " & _
                      " from Acme.dbo.tbl_i_users " & _
                      " where userid='" + Txt_UserID.Text + "'  AND permit=1 " & _
                      "   and password   =  substring(sys.fn_VarBinToHexStr(hashbytes('MD5','" + txt_pwd.Text + "')),3,32)  "
            '取得登陸信息 , 保留登入之帳號,可供各程式使用
            With db.FillDataSet(sql_txt).Tables(0)
                IS_USER.emp_no = .Rows(0).Item("userid")
                IS_USER.UserID = .Rows(0).Item("userid")
                IS_USER.IPID = Environment.MachineName.ToString
                IS_USER.emp_name = .Rows(0).Item("USERNAME")
                Fm_main.txt_Use.Text = .Rows(0).Item("userid").ToString.Trim + "-" + .Rows(0).Item("USERNAME")
            End With
            Me.DialogResult = vbOK
            'Application.Run(New Fm_main())
            Dim connectionString As String = ""
            'Dim strUpdateXmlPath As String = Application.StartupPath + "\update\conf\update.xml"
            'connectionString = Mod_Function.Decrypt(ERP.Autoupdate.getConfigValue(strUpdateXmlPath, "ConnectionString"))
            connectionString = GetAppConfig.remoteConnect
            Try
                'SQL = "SELECT * FROM Acme.dbo.webURL WHERE connection='" & connectionString & "' "
                'If db.FillDataSet(SQL).Tables(0).Rows.Count > 0 Then
                '    WEBURL.URL = db.FillDataSet(SQL).Tables(0).Rows(0).Item("URL")
                '    WEBURL.RptURL = db.FillDataSet(SQL).Tables(0).Rows(0).Item("RptURL")
                '    WEBURL.LoginUserId = db.FillDataSet(SQL).Tables(0).Rows(0).Item("LoginUserId")
                '    WEBURL.PassWord = db.FillDataSet(SQL).Tables(0).Rows(0).Item("PassWord")
                '    IS_FRMNAME = WEBURL.URL + "Login.aspx?EmpId=" & IS_USER.UserID & ""
                '    Fm_main.WebBrowser1.Navigate(IS_FRMNAME)
                'Else
                '    MsgBox("請確認使用資料庫是否正確")
                '    Exit Sub
                'End If

                Dim subKey As Microsoft.Win32.RegistryKey
                subKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\UNIMICRON\SD\MIS", True)
                subKey.SetValue("LastLoginUser", Txt_UserID.Text.Trim)
            Catch ex As Exception
                Message = ErrorMessage(ex.Message.ToString)
                MsgBox(Message)
                Exit Sub
            End Try
        End If
    End Sub
#Region "去掉字符串中的特殊字符"
    Public Function ErrorMessage(ByVal bufstr As Object) As Object
        Return bufstr.Replace(vbCrLf, "").Replace("'", "")
    End Function
#End Region
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        '離開並關閉執行緒
        Environment.Exit(Environment.ExitCode)
        Application.Exit()
        Close()
    End Sub

    Private Sub MsgBox_AutoClose(ByVal z_sec As Integer, ByVal z_msg As String)
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000 * z_sec
        MsgBox(z_msg, MsgBoxStyle.OkOnly, "MsgBox")
        Me.Timer1.Enabled = False
    End Sub
    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim hWnd As Integer
        hWnd = FindWindow(vbNullString, "MsgBox")
        If hWnd Then
            PostMessage(hWnd, &H10, 0&, 0&)
        End If
    End Sub

    Private Sub FrmLogin_Activated(sender As System.Object, e As System.EventArgs) Handles MyBase.Activated
        Dim subKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\UNIMICRON\SD\MIS")

        If IsNothing(subKey) Then
            Microsoft.Win32.Registry.LocalMachine.CreateSubKey("SOFTWARE\UNIMICRON\SD\MIS")
            subKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\UNIMICRON\SD\MIS", True)
            subKey.SetValue("LastLoginUser", System.Environment.UserName)
            Txt_UserID.Text = System.Environment.UserName
        Else
            Dim value As String = subKey.GetValue("LastLoginUser")
            Txt_UserID.Text = value
        End If
        txt_pwd.Focus()
    End Sub

    Private Sub FrmLogin_FormClosed(sender As System.Object, e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If DialogResult <> vbOK Then
            '離開並關閉執行緒
            Environment.Exit(Environment.ExitCode)
            Application.Exit()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '離開並關閉執行緒
        Environment.Exit(Environment.ExitCode)
        Application.Exit()
        Close()
    End Sub

    
End Class