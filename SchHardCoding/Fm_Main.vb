'Imports System.Net
'Imports System.Net.Dns
'Imports System.Net.NetworkInformation
Imports System.Text.RegularExpressions
Imports System.IO
Public Class Fm_main
    Dim dsline As DataSet = New DataSet()
    Dim dt As DataTable = New DataTable()
    
    Dim matchStrs As List(Of String) = New List(Of String)
    Private Sub Fm_main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If FrmLogin.ShowDialog <> vbOK Then
            Application.Exit()
            Me.Close()
        End If
        'Dim dcEmpNo As DataColumn = New DataColumn("操作人", GetType(String))
        'Dim dcFileName As DataColumn = New DataColumn("檔案", GetType(String))
        'Dim dcLineNo As DataColumn = New DataColumn("行號", GetType(String))
        'Dim dcCode As DataColumn = New DataColumn("程式碼", GetType(String))
        'Dim dcSchTime As DataColumn = New DataColumn("查找時間", GetType(String))
        dt.Columns.Add("操作人", GetType(String))
        dt.Columns.Add("檔案", GetType(String))
        dt.Columns.Add("行號", GetType(String))
        dt.Columns.Add("程式碼", GetType(String))
        dt.Columns.Add("查找時間", GetType(String))

        dsline.Clear()
        dsline.Tables.Add(dt)

        rbEptExegYes.Checked = True
        'matchStrs.Add("\d*\.\d*\.\d*\.\d*")
        matchStrs.Add("\'[vV][iI]\'")
        Label4.Text = "Lips：點擊""選擇檔案""按鈕，" + vbCrLf + "選中需查找的檔案，即可查找完成"
        'matchStr.Add("\.[cC][oO][mM]")
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Try
            Dim NewCondi As String = InputBox("請輸入查找字串(支持正則表達式)", "確認")
            lbxSchCon.Items.Add(NewCondi.Trim)
            matchStrs.Add(NewCondi.Trim)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub lbxSchCon_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lbxSchCon.KeyDown
        Try
            If e.KeyCode = Keys.Delete Or e.KeyCode = Keys.Back Then 'delete鍵碼
                For Each SelectedItem As Object In lbxSchCon.SelectedItems
                    'lbxSchCon.Items.Remove(SelectedItem)
                    matchStrs.Remove(SelectedItem.ToString)
                Next
                Dim indices As ListBox.SelectedIndexCollection = lbxSchCon.SelectedIndices
                Dim count As Integer = indices.Count
                lbxSchCon.BeginUpdate()
                For i As Integer = 0 To count - 1
                    lbxSchCon.Items.RemoveAt(indices(0))
                Next
                lbxSchCon.EndUpdate()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
        Try
            For Each SelectedItem As Object In lbxSchCon.SelectedItems
                'lbxSchCon.Items.Remove(SelectedItem)
                matchStrs.Remove(SelectedItem.ToString)
            Next
            Dim indices As ListBox.SelectedIndexCollection = lbxSchCon.SelectedIndices
            Dim count As Integer = indices.Count
            lbxSchCon.BeginUpdate()
            For i As Integer = 0 To count - 1
                lbxSchCon.Items.RemoveAt(indices(0))
            Next
            lbxSchCon.EndUpdate()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnSelFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelFile.Click
        '2024/08/06 沈騰 選取文件方式
        'Try
        '    Dim NowTime As String = Format(Now(), "yyyy/MM/dd HH:mm:ss")
        '    Dim FileOpen As New OpenFileDialog()
        '    FileOpen.Multiselect = True
        '    FileOpen.Title = "請選擇需查找的文件"
        '    FileOpen.Filter = "Visual Basic 程式碼檔案(*.vb;*.resx;*.settings;*.xsd;*.wsdl)|*.vb;*.resx;*.settings;*.xsd;*.wsdl"
        '    Dim SelectFile, fileName As String
        '    If FileOpen.ShowDialog() = DialogResult.OK Then
        '        ''Dim fileReader As System.IO.StreamReader
        '        'fileReader = My.Computer.FileSystem.OpenTextFileReader("C:\\testfile.txt")
        '        dt.Clear()
        '        For Each SelectFile In FileOpen.FileNames
        '            Dim indexLine As Integer = 0
        '            fileName = System.IO.Path.GetFileName(SelectFile)
        '            Dim LineStrs As List(Of String) = IO.File.ReadAllLines(SelectFile).ToList
        '            For Each line As String In LineStrs
        '                indexLine = LineStrs.IndexOf(line)
        '                line = line.Trim
        '                For Each MatchStr As String In matchStrs
        '                    If rbEptExegYes.Checked = True And Microsoft.VisualBasic.Left(line, 1) = "'" Then
        '                    Else
        '                        Match(line, MatchStr, fileName, indexLine, NowTime)
        '                    End If
        '                Next
        '            Next
        '        Next
        '        DataGridView1.DataSource = dt
        '        LbNum.Text = dt.Rows.Count
        '        If LbNum.Text <> "0" Then
        '            LbNum.ForeColor = Color.Red
        '        Else
        '            LbNum.ForeColor = Color.FromArgb(0, 192, 0)
        '        End If

        '        dsline.Clear()
        '        dsline.Tables.Add(dt)

        '    Else
        '        Exit Sub
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
        Try
            dt.Clear()
            Dim NowTime As String = Format(Now(), "yyyy/MM/dd HH:mm:ss")
            If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
                ' List files in the folder.
                Dim fileNames = My.Computer.FileSystem.GetFiles(
        FolderBrowserDialog1.SelectedPath, FileIO.SearchOption.SearchAllSubDirectories, "*.vb")
                '參數：SearchTopLevelOnly，當前文件夾
                '參數：SearchAllSubDirectories，包含子目錄
                Dim fileName As String
                For Each SelectFile As String In fileNames
                    Dim indexLine As Integer = 0
                    fileName = System.IO.Path.GetFileName(SelectFile)
                    Dim LineStrs As List(Of String) = IO.File.ReadAllLines(SelectFile).ToList
                    For Each line As String In LineStrs
                        indexLine = LineStrs.IndexOf(line)
                        line = line.Trim
                        For Each MatchStr As String In matchStrs
                            If rbEptExegYes.Checked = True And Microsoft.VisualBasic.Left(line, 1) = "'" Then
                            Else
                                Match(line, MatchStr, fileName, indexLine, NowTime)
                            End If
                        Next
                    Next
                Next
                DataGridView1.DataSource = dt
                LbNum.Text = dt.Rows.Count
                If LbNum.Text <> "0" Then
                    LbNum.ForeColor = Color.Red
                Else
                    LbNum.ForeColor = Color.FromArgb(0, 192, 0)
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub Match(ByVal line As String, ByVal expr As String, ByVal fileName As String, ByVal indexLine As Integer, ByVal NowTime As String)
        Dim mc As MatchCollection = Regex.Matches(line, expr)
        If mc.Count <> 0 Then
            Dim dr As DataRow = dt.NewRow()
            dr("操作人") = IS_USER.emp_no + "-" + IS_USER.emp_name
            dr("檔案") = fileName
            dr("行號") = indexLine
            dr("程式碼") = line
            dr("查找時間") = NowTime
            dt.Rows.Add(dr)
        End If
    End Sub

    Private Sub rbHasExegYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbEptExegYes.CheckedChanged
        If rbEptExegYes.Checked = True Then
            rbEptExegNo.Checked = False
        Else
            rbEptExegNo.Checked = True
        End If
    End Sub

    Private Sub rbHasExegNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbEptExegNo.CheckedChanged
        If rbEptExegNo.Checked = True Then
            rbEptExegYes.Checked = False
        Else
            rbEptExegYes.Checked = True
        End If
    End Sub
    '結束系統
    Private Sub Fm_main_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Dim ii As MsgBoxResult
        ii = MsgBox("確定離開?", MsgBoxStyle.OkCancel, "結束系統")
        If ii = MsgBoxResult.Cancel Then
            e.Cancel = True
        End If
    End Sub
    '2024/04/25 沈騰 Excel導出功能
    Private Sub btnExl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExl.Click
        If Not My.Computer.FileSystem.DirectoryExists(P_DownPath) Then
            My.Computer.FileSystem.CreateDirectory(P_DownPath)
        End If
        Dim AlHeader As New ArrayList
        Dim AlData As New ArrayList
        For i = 0 To DataGridView1.Columns.Count - 1
            AlHeader.Add(DataGridView1.Columns(i).HeaderText)
            AlData.Add(TryCast(DataGridView1.Columns(i), DataGridViewColumn).DataPropertyName)
        Next
        Dim mySDate As DateTime = System.DateTime.Now
        Try
            Dim myDate As DateTime = System.DateTime.Now
            Dim myFileName As String = String.Format(P_DownPath & "\查找VI_{0}{1}{2}_{3}{4}{5}.xlsx", myDate.Year, myDate.Month.ToString("00"), myDate.Day.ToString("00"), myDate.Hour.ToString("00"), myDate.Minute.ToString("00"), myDate.Second.ToString("00"))
            Dim myFile As FileStream = New FileStream(myFileName, FileMode.Create)
            Mod_Function.CreateExcel2007(AlHeader, AlData, dsLine).Write(myFile)
            myFile.Close()
            Dim Zstr As String = ""
            Zstr = "Excel 檔案已產生!!" & vbNewLine
            Zstr &= "檔案存放於    " & myFileName & vbNewLine
            Zstr &= "" & vbNewLine
            Zstr &= "是否開啟 Excel 檔 ?  "
            If MsgBox(Zstr, MsgBoxStyle.YesNo, "匯出Excel檔") = MsgBoxResult.Yes Then
                Dim OtherDoc1 As New System.Diagnostics.Process
                OtherDoc1.StartInfo.FileName = myFileName   '"檔案名稱(含路徑)"
                OtherDoc1.Start()
            End If
            Zstr = Nothing
        Catch ex As Exception
            MsgBox(String.Format("Excel 檔案產生失敗!!失敗原因:{0}", ex.ToString()), MsgBoxStyle.Information, "存檔完成")
        End Try
    End Sub
End Class
