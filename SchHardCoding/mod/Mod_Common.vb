Imports System.IO
Imports System.Net
Imports System.Net.Dns
Imports System.Net.NetworkInformation
Imports System.Data.OleDb
Imports System.Text
Module Mod_Common
    Dim db As SqlDb = New SqlDb()
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

    Public P_ServerType As String
    Public P_DownPath As String = "C:\Temp"   '匯出Excel之目錄檔
    Public P_DownPath1 As String = "C:\Temp"

#Region "抓取成本系統確認年月"
    Function load_datepkerYM() As String
        Dim datepkerYM As String = Now.ToString("yyyy/MM")
        If db.GetExecSQL("exec CAC_isTableExists 'ClassParams'") <> 1 Then
            Return db.GetExecSQL("select ParamsOne from dbParams where Defines ='CAC_SetMonth'")
        End If
        Return datepkerYM
    End Function
#End Region

#Region "去掉字符串中的特殊字符"
    Public Function ErrorMessage(ByVal bufstr As Object) As Object
        Return bufstr.Replace(vbCrLf, "").Replace("'", "")
    End Function
#End Region

#Region "判斷字串是否為DBULL"
    Public Function StringIsDBnull(ByVal bufstr As Object) As Object
        If (bufstr Is System.DBNull.Value) Then
            Return ""
        Else
            Return bufstr
        End If
    End Function
#End Region

#Region "判斷字串是否為Nothing"

    ''' <summary>
    ''' 判斷字串是否為Nothing
    ''' </summary>
    ''' <param name="bufstr">要判斷的字串變數</param>
    ''' <returns>回傳True表示為Nothing , 回傳False表示不為Nothing</returns>
    ''' <remarks></remarks>
    Public Function StringIsNothing(ByVal bufstr) As Boolean

        If (bufstr Is Nothing) Then
            Return True
        Else
            Return False
        End If

    End Function

#End Region

#Region "Frm_CLEAR:清除窗體中文本框的內容"
    '清除窗體中文本框的內容
    Public Sub Frm_CLEAR(ByRef VAR_OBJ As Object)
        For Each is_t As Object In VAR_OBJ.Controls
            Select Case is_t.GetType.ToString
                Case "System.Windows.Forms.TextBox"
                    CType(is_t, TextBox).Clear()
                Case "System.Windows.Forms.MaskedTextBox"
                    Frm_CLEAR(is_t)
                Case "System.Windows.Forms.GroupBox"
                    Frm_CLEAR(is_t)
                Case "System.Windows.Forms.Panel"
                    Frm_CLEAR(is_t)
                Case "System.Windows.Forms.TableLayoutPanel"
                    Frm_CLEAR(is_t)
                Case "System.Windows.Forms.DataGridview"
                    Frm_CLEAR(is_t)
                Case Else

            End Select

            If (TypeOf is_t Is GroupBox) Or (TypeOf is_t Is TableLayoutPanel) Then
                For Each is_t2 As Object In is_t.Controls
                    Select Case is_t2.GetType.ToString
                        Case "System.Windows.Forms.TextBox"
                            CType(is_t2, TextBox).Clear()
                        Case "System.Windows.Forms.MaskedTextBox"
                            Frm_CLEAR(is_t2)
                        Case "System.Windows.Forms.GroupBox"
                            Frm_CLEAR(is_t2)
                        Case "System.Windows.Forms.Panel"
                            Frm_CLEAR(is_t2)
                        Case "System.Windows.Forms.TableLayoutPanel"
                            Frm_CLEAR(is_t2)
                        Case Else

                    End Select
                Next
            End If
        Next
    End Sub
#End Region

    '獲取系統中寫死路徑
    Public Function FUN_GET_URLAddress(ByRef SystemID As String, ByRef ItemName As String) As String
        Dim is_sql As String
        is_sql = " SELECT URLAddress FROM dbo.ERPURL WHERE SystemID='" & SystemID & "' AND ItemName='" & ItemName & "' "
        Dim Path As String = ""
        With db.FillDataSet(is_sql).Tables(0)
            If .Rows.Count > 0 Then
                Path = .Rows(0).Item("URLAddress").ToString
            End If
        End With

        Return Path

    End Function

#Region "Frm_init_frm:控制控件的可编辑性---ReadOnly,IsReadOnly"
    Public Sub Frm_init_frm(ByRef VAR_OBJ As Object, ByVal var_switch As Boolean)
        For Each is_t As Object In VAR_OBJ.Controls
            Select Case is_t.GetType.ToString
                Case "System.Windows.Forms.TextBox"
                    is_t.ReadOnly = var_switch
                    is_t.forecolor = Color.Black
                Case "System.Windows.Forms.CheckBox"
                    is_t.forecolor = Color.Black
                    is_t.enabled = Not var_switch
                Case "System.Windows.Forms.MaskedTextBox"
                    is_t.ReadOnly = var_switch
                    is_t.forecolor = Color.Black
                Case "System.Windows.Forms.DataGridView"
                    is_t.ReadOnly = var_switch
                    is_t.forecolor = Color.Black
                Case "System.Windows.Forms.DateTimePicker"
                    is_t.forecolor = Color.Black
                    is_t.enabled = Not var_switch
                Case "System.Windows.Forms.ComboBox"
                    is_t.forecolor = Color.Black
                    is_t.enabled = Not var_switch
                Case "mySpecialComboBox.mySpecialComboBox"    '自訂元件,設定IsReadOnly=True/false,解決原Enable=false時,字體變灰看不清之問題!
                    is_t.IsReadOnly = var_switch              '設Enable=false時,字體變灰看不清
                Case "myObject.myComboBox"                    '自訂元件,設定IsReadOnly=True/false,解決原Enable=false時,字體變灰看不清之問題!
                    is_t.IsReadOnly = var_switch              '設Enable=false時,字體變灰看不清
                Case "myObject.myCheckBox"
                    is_t.IsReadOnly = var_switch
                Case "myObject.myRadioButton"
                    is_t.IsReadOnly = var_switch
                Case "myObject.myDateTimePicker"
                    is_t.IsReadOnly = var_switch
            End Select

            If (TypeOf is_t Is GroupBox) Or (TypeOf is_t Is TableLayoutPanel) Then
                For Each is_txt As Object In is_t.Controls
                    Select Case is_txt.GetType.ToString
                        Case "System.Windows.Forms.TextBox"
                            is_txt.ReadOnly = var_switch
                            is_txt.forecolor = Color.Black
                        Case "System.Windows.Forms.CheckBox"
                            is_txt.forecolor = Color.Black
                            is_txt.enabled = Not var_switch
                        Case "System.Windows.Forms.MaskedTextBox"
                            is_txt.ReadOnly = var_switch
                            is_txt.forecolor = Color.Black
                        Case "System.Windows.Forms.DataGridView"
                            is_txt.ReadOnly = var_switch
                            is_txt.forecolor = Color.Black
                        Case "System.Windows.Forms.DateTimePicker"
                            is_txt.forecolor = Color.Black
                            is_txt.enabled = Not var_switch
                        Case "System.Windows.Forms.ComboBox"
                            is_txt.forecolor = Color.Black
                            is_txt.enabled = Not var_switch
                        Case "mySpecialComboBox.mySpecialComboBox"   '自訂元件,設定IsReadOnly=True/false,解決原Enable=false時,字體變灰看不清之問題!
                            is_txt.IsReadOnly = var_switch           '設Enable=false時,字體變灰看不清
                        Case "myObject.myComboBox"                   '自訂元件,設定IsReadOnly=True/false,解決原Enable=false時,字體變灰看不清之問題!
                            is_txt.IsReadOnly = var_switch           '設Enable=false時,字體變灰看不清
                        Case "myObject.myCheckBox"
                            is_txt.IsReadOnly = var_switch
                        Case "myObject.myRadioButton"
                            is_txt.IsReadOnly = var_switch
                        Case "myObject.myDateTimePicker"
                            is_txt.IsReadOnly = var_switch
                    End Select
                Next
            End If
        Next
    End Sub
#End Region

#Region "將字符串轉化成數字"
    '將字符串轉化成數字
    Public Function str2sng(ByVal var_str As Object, Optional ByVal var_format As String = "") As Single
        Dim is_res As Single
        Dim z_str As String
        z_str = var_str & ""
        If IsDBNull(var_str) Or (Trim(z_str).Length = 0) Then
            is_res = 0
        Else
            is_res = CSng(var_str)
        End If
        If var_format <> "" Then
            Return Format(is_res, var_format)
        Else
            Return is_res
        End If
    End Function
#End Region

#Region "Set_Combox公用程式"
    Public Sub Set_Combox(ByRef z_Combox As Object, ByRef Array_ID As Array, ByRef Array_Name As Array)
        '呼叫方式
        Dim DT As DataTable = New DataTable
        Dim dr As DataRow
        DT.Columns.Add("z_ID", GetType(String))
        DT.Columns.Add("z_Name", GetType(String))
        Dim ii As Integer = 0
        For ii = 0 To UBound(Array_ID)
            dr = DT.NewRow()
            dr.Item(0) = Array_ID(ii)
            dr.Item(1) = Array_Name(ii)
            DT.Rows.Add(dr)
        Next
        z_Combox.DisplayMember = "z_Name"
        z_Combox.ValueMember = "z_ID"
        z_Combox.DataSource = DT
    End Sub

#End Region

#Region "Set_Cob_LindId-綁定線別"
    '綁定線別
    Public Sub Set_Cob_Line(ByRef var_comb As ComboBox)
        var_comb.DataSource = Nothing
        var_comb.Items.Clear()
        Dim is_sql As String
        is_sql = "select LineId,LineName=cast(LineId as varchar)+'_'+Rtrim(LineName) from ClassLine(Nolock) order by LineId "
        var_comb.DataSource = db.FillDataSet(is_sql).Tables(0)
        var_comb.DisplayMember = "LineName"
        var_comb.ValueMember = "LineId"
        var_comb.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        var_comb.AutoCompleteSource = AutoCompleteSource.ListItems
        var_comb.SelectedIndex = -1
    End Sub
#End Region

#Region "Set_Cob_PNType-綁定料號型態"
    '綁定料號型態
    Public Sub Set_Cob_PNType(ByRef var_comb As ComboBox)
        Dim Array_ID() As String = {"0", "1"}
        Dim Array_Name() As String = {"量產", "樣品"}
        Call Set_Combox(var_comb, Array_ID, Array_Name)
        var_comb.SelectedIndex = -1
    End Sub
#End Region

#Region "Set_Cob_DeptCode-綁定部門別"
    '綁定部門別
    Public Sub Set_Cob_DeptCode(ByRef var_comb As ComboBox, ByVal var_YM As String)
        var_comb.DataSource = Nothing
        var_comb.Items.Clear()
        Dim is_sql As String
        If var_YM <> "/01" Then
            is_sql = "Select UnitId,UnitName=UnitId+'_'+UnitName from dbo.CAC_DepartProc(nolock) where CostMonth = '" & var_YM & "' order by UnitId"
            var_comb.DataSource = db.FillDataSet(is_sql).Tables(0)
            var_comb.DisplayMember = "UnitName"
            var_comb.ValueMember = "UnitId"
            var_comb.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            var_comb.AutoCompleteSource = AutoCompleteSource.ListItems
            var_comb.SelectedIndex = -1
        End If
    End Sub
#End Region

#Region "Set_Cob_ProcCode-綁定製程名稱"
    '綁定製程名稱
    Public Sub Set_Cob_ProcCode(ByRef var_comb As ComboBox)
        var_comb.DataSource = Nothing
        var_comb.Items.Clear()
        Dim is_sql As String
        is_sql = "Select CustProcCode,ProcName=CustProcCode+'_'+ProcName from dbo.ProcBasic(nolock) order by ProcCode"
        var_comb.DataSource = db.FillDataSet(is_sql).Tables(0)
        var_comb.DisplayMember = "ProcName"
        var_comb.ValueMember = "CustProcCode"
        var_comb.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        var_comb.AutoCompleteSource = AutoCompleteSource.ListItems
        var_comb.SelectedIndex = -1
    End Sub
#End Region

#Region "Set_Cob_CostClass-綁定成本類別"
    '綁定成本類別
    Public Sub Set_Cob_CostClass(ByRef var_comb As ComboBox)
        var_comb.DataSource = Nothing
        var_comb.Items.Clear()
        Dim is_sql As String
        is_sql = "select ClassId,ClassName=cast(ClassId as varchar)+'_'+Rtrim(ClassName) from CAC_CostClass(nolock) where  Classid<>6"
        var_comb.DataSource = db.FillDataSet(is_sql).Tables(0)
        var_comb.DisplayMember = "ClassName"
        var_comb.ValueMember = "ClassId"
        var_comb.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        var_comb.AutoCompleteSource = AutoCompleteSource.ListItems
        var_comb.SelectedIndex = -1
    End Sub
#End Region
    '獲取操作碼--財務系統
    Public Function FUN_GET_GlobalId(ByRef empid As String) As String
        Dim vGlobalId, vR As String
        vGlobalId = ""
        Randomize()
        vR = Int(Rnd() * 1000).ToString
        If Len(vR) = 2 Then
            vR = "0" + vR
        End If
        If Len(vR) = 1 Then
            vR = "00" + vR
        End If
        vGlobalId = vR + empid.Trim + Mid(Format(Now, "yyyyMMddHHmmssms"), 4, 14)
        Return vGlobalId
    End Function
    '取三分編碼
    Function FUN_UTA_GET_GPRID(ByVal UseId As String, ByVal z_var1 As String) As String
        Dim is_S As String
        is_S = " select ISNULL(GPRID,'') GPRID  " & _
               " from UTA_JourBas " & _
               " where UseId= '" & UseId & "'" & _
               " and JourId='" & z_var1 & "'"
        If db.FillDataSet(is_S).Tables(0).Rows.Count > 0 Then
            Return db.GetExecSQL(is_S)
        Else
            Return ""
        End If

    End Function

    '部門別是否失效
    Public Function FUN_UTA_IsDepartDisable(ByRef DepartId As String) As Boolean
        Dim is_S As String
        Dim Result As Boolean
        Dim dt As DataTable
        is_S = "  Select 1 From CompanyStruct(nolock) " & _
               " Where UnitId='" & DepartId & "' and IsDisable=1 "

        dt = db.FillDataSet(is_S).Tables(0)
        If dt.Rows.Count > 0 Then
            Result = True
        Else
            Result = False
        End If
        Return Result
    End Function

    Public Sub BindCbo(ByVal sql As String, ByVal text As String, ByVal value As String, ByRef var_comb As ComboBox)
        Dim dt As New DataTable
        dt.Columns.Add("DisplayMember", System.Type.GetType("System.String"))
        dt.Columns.Add("ValueMember", System.Type.GetType("System.String"))
        Dim ds As DataSet = db.FillDataSet(sql)
        Dim i = 0
        For i = 0 To ds.Tables(0).Rows.Count - 1
            Dim dr As DataRow = dt.NewRow()
            dr(0) = ds.Tables(0).Rows(i).Item(value).ToString & "_" & ds.Tables(0).Rows(i).Item(text).ToString
            dr(1) = ds.Tables(0).Rows(i).Item(value).ToString
            dt.Rows.Add(dr)
        Next

        var_comb.DataSource = Nothing
        var_comb.Items.Clear()
        var_comb.DataSource = dt
        var_comb.DisplayMember = "DisplayMember"
        var_comb.ValueMember = "ValueMember"
        var_comb.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        var_comb.AutoCompleteSource = AutoCompleteSource.ListItems
        var_comb.SelectedIndex = -1
    End Sub

    Public Sub BindCombox(ByVal sql As String, ByVal text As String, ByVal value As String, ByRef var_comb As ComboBox)
        Dim dt As New DataTable
        dt.Columns.Add("DisplayMember", System.Type.GetType("System.String"))
        dt.Columns.Add("ValueMember", System.Type.GetType("System.String"))
        dt.Rows.Add(" ", " ")
        Dim ds As DataSet = db.FillDataSet(sql)
        Dim i = 0
        For i = 0 To ds.Tables(0).Rows.Count - 1
            Dim dr As DataRow = dt.NewRow()
            dr(0) = ds.Tables(0).Rows(i).Item(text).ToString
            dr(1) = ds.Tables(0).Rows(i).Item(value).ToString
            dt.Rows.Add(dr)
        Next

        var_comb.DataSource = Nothing
        var_comb.Items.Clear()
        var_comb.DataSource = dt
        var_comb.DisplayMember = "DisplayMember"
        var_comb.ValueMember = "ValueMember"
        var_comb.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        var_comb.AutoCompleteSource = AutoCompleteSource.ListItems
        var_comb.SelectedIndex = 0
    End Sub
    '此數據綁定用於ComboBox 中 DropDownStyle屬性是DropDownList下拉選不可修改的時候
    Public Sub BandDropColumn_3(ByVal sql As String, ByVal text As String, ByVal value As String, ByRef var_comb As ComboBox)
        Dim dt As New DataTable
        dt.Columns.Add("DisplayMember", System.Type.GetType("System.String"))
        dt.Columns.Add("ValueMember", System.Type.GetType("System.String"))
        dt.Rows.Add(" ", " ")
        Dim ds As DataSet = db.FillDataSet(sql)
        Dim i = 0
        For i = 0 To ds.Tables(0).Rows.Count - 1
            Dim dr As DataRow = dt.NewRow()
            dr(0) = ds.Tables(0).Rows(i).Item(text).ToString
            dr(1) = ds.Tables(0).Rows(i).Item(value).ToString
            dt.Rows.Add(dr)
        Next

        var_comb.DataSource = Nothing
        var_comb.Items.Clear()
        var_comb.DataSource = dt
        var_comb.DisplayMember = "DisplayMember"
        var_comb.ValueMember = "ValueMember"
        'var_comb.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        'var_comb.AutoCompleteSource = AutoCompleteSource.ListItems
        var_comb.SelectedIndex = 1

    End Sub
    '此數據綁定用於ComboBox 中 DropDownStyle屬性是DropDownList下拉選不可修改的時候
    Public Sub BandDropColumn_2(ByVal sql As String, ByVal text As String, ByVal value As String, ByRef var_comb As ComboBox)
        Dim dt As New DataTable
        dt.Columns.Add("DisplayMember", System.Type.GetType("System.String"))
        dt.Columns.Add("ValueMember", System.Type.GetType("System.String"))
        dt.Rows.Add(" ", " ")
        Dim ds As DataSet = db.FillDataSet(sql)
        Dim i = 0
        For i = 0 To ds.Tables(0).Rows.Count - 1
            Dim dr As DataRow = dt.NewRow()
            dr(0) = ds.Tables(0).Rows(i).Item(text).ToString
            dr(1) = ds.Tables(0).Rows(i).Item(value).ToString
            dt.Rows.Add(dr)
        Next

        var_comb.DataSource = Nothing
        var_comb.Items.Clear()
        var_comb.DataSource = dt
        var_comb.DisplayMember = "DisplayMember"
        var_comb.ValueMember = "ValueMember"
        'var_comb.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        'var_comb.AutoCompleteSource = AutoCompleteSource.ListItems
        var_comb.SelectedIndex = 0

    End Sub

    '下拉--財務系統
    Public Sub bandDrop(ByVal sql As String, ByVal text As String, ByVal value As String, ByVal var_comb As ComboBox)
        sql = "select '' AS " + value + ", ''  AS " + text + " UNION ALL " + sql
        var_comb.DataSource = Nothing
        var_comb.Items.Clear()
        var_comb.DataSource = db.FillDataSet(sql).Tables(0)
        var_comb.DisplayMember = text
        var_comb.ValueMember = value
        var_comb.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        var_comb.AutoCompleteSource = AutoCompleteSource.ListItems
        var_comb.SelectedIndex = -1
    End Sub
    '下拉--財務系統
    Public Sub bandDropColumn(ByVal sql As String, ByVal text As String, ByVal value As String, ByVal var_comb As DataGridViewComboBoxColumn)
        var_comb.DataSource = Nothing
        var_comb.Items.Clear()
        var_comb.DataSource = db.FillDataSet(sql).Tables(0)
        var_comb.DisplayMember = text
        var_comb.ValueMember = value
        'var_comb.Selected = False
    End Sub
    '獲取財務系統中的權限等級
    Public Function FUN_GET_UTA_AuthorLevel(ByRef empid As String) As String
        Dim is_sql As String
        is_sql = " SELECT Accountant  FROM UTA_Private t1(NOLOCK) " & _
        " WHERE   " & _
        "  t1.EmpId = '" & empid & "' "
        Return db.GetExecSQL(is_sql)
    End Function
    '獲取財務系統中的權限等級
    Public Function FUN_GET_UTA_Manager(ByRef empid As String) As String
        Dim is_sql As String
        is_sql = " SELECT Directorid FROM dbo.PSN_DepDirector WHERE UnitId='00UF0000'"
        Return db.GetExecSQL(is_sql)
    End Function

#Region "UTA 生產傳票 孫婷婷"
    Public Function SetJourId(ByRef empid As String) As String
        Dim dt As DataTable
        Dim SQLStringList As New ArrayList '存儲更新語句
        Dim is_sql As String

        is_sql = " select AccSId from UTA_JourBas(nolock)where isnull(JourId,'')='' " & _
                 "and Accountant='" & FUN_GET_UTA_AuthorLevel(empid) & "' and   Booker ='" & empid & "'"
        dt = db.FillDataSet(is_sql).Tables(0)
        If dt.Rows.Count = 0 Then
            Return "false"
            Exit Function
        End If
        Dim vJourId_ALL = ""
        '取得傳票號碼
        Dim GlobalId = StringIsDBnull(dt.Rows(0).Item("AccSId"))
        Dim Userid = FUN_GET_FAS_COM()
        vJourId_ALL = FUN_GET_JourId(Userid, GlobalId)
        If vJourId_ALL = "" Then
            Return ""
            Exit Function
        End If
        SQLStringList.Clear()
        Dim p_str = "  update UTA_JourBas set JourId='" & vJourId_ALL & "' " & _
              "where UseId='" & Userid & "' and   AccSId='" & GlobalId & "' " & _
               " delete UTA_JourDtl where UseId='" & Userid & "' and   AccSId='" & GlobalId & "' " & _
               " and   (isnull(DAmount,0) + isnull(CAmount,0))=0 "
        SQLStringList.Add(p_str)
        db.ExecuteSqlTran(SQLStringList)

        SQLStringList.Clear()
        p_str = "  delete UTA_JourDtl where UseId='" & Userid & "' and  AccSId='" & GlobalId & "' " & _
               " and  (isnull(DAmount,0) + isnull(CAmount,0))=0  "
        SQLStringList.Add(p_str)
        db.ExecuteSqlTran(SQLStringList)

        Return vJourId_ALL
    End Function
#End Region

#Region "抓取財務系統公司別"
    Function FUN_GET_FAS_COM() As String
        Dim str As String 'sql語句
        str = "select UseId from UTA_AccUse where defau=1"
        With db.FillDataSet(str).Tables(0)
            If .Rows.Count > 0 Then
                Return .Rows(0).Item("UseId").ToString()
            Else
                Return ""
            End If
        End With
    End Function
#End Region
    '獲取傳票項次
    Public Function FUN_GET_Item(ByRef UseId As String, ByRef AccSId As String) As String
        Dim is_S As String
        Dim Item As Integer
        Dim dt As DataTable
        is_S = " select MaxItem= isnull(max(Item),0)  " & _
               " from UTA_JourDtl " & _
               " where UseId= '" & UseId & "'" & _
               " and AccSId='" & AccSId & "'"
        dt = db.FillDataSet(is_S).Tables(0)
        If dt.Rows.Count > 0 Then
            Item = CInt(dt.Rows(0).Item("MaxItem").ToString) + 1
        Else
            Item = 1
        End If
        Return Item
    End Function
    '獲取匯率 lijuan
    Public Function FUN_GET_RateToNT(jourdate As String) As String
        Dim is_S, flag As String
        Dim dt As DataTable
        flag = 0
        is_S = " select MoneyCode from UTA_AccUse where Defau=1 "
        dt = db.FillDataSet(is_S).Tables(0)
        If dt.Rows.Count > 0 Then
            dt = db.FillDataSet("exec UTA_SpGetRate '" + dt.Rows(0).Item("MoneyCode").ToString + "','" + jourdate + "'").Tables(0)
            If dt.Rows.Count > 0 Then
                flag = dt.Rows(0).Item("RateToNT").ToString
            Else
                flag = "1"
            End If
        End If
        Return flag
    End Function
    '日匯率抓取
    Public Function FUN_GET_Rate(Moneycode As String, vDate As String) As String
        Dim is_S, flag As String
        Dim dt As DataTable
        flag = 0

        is_S = " exec UTA_SpGetRate '" & Moneycode & "','" & vDate & "' "
        dt = db.FillDataSet(is_S).Tables(0)
        If dt.Rows.Count > 0 Then
            flag = dt.Rows(0).Item("RateTONT").ToString
        End If
        Return flag
    End Function

    '修改，刪除時判斷是否有立賬 
    Public Function ChkGLNum(kind As String, item As String, accsid As String, useid As String) As String
        Dim sql, flag As String
        Dim dt As DataTable
        sql = "    DECLARE @GLNum uChrGlobalId, @IsTranBill INT"
        sql = sql + " IF '" + kind + "' = 0"
        sql = sql + " BEGIN"
        sql = sql + "  IF EXISTS(SELECT GLNum"
        sql = sql + "            FROM UTA_JourDtl(NOLOCK)"
        sql = sql + "             WHERE AccSId ='" + accsid + "'"
        sql = sql + "                AND UseId ='" + useid + "'"
        sql = sql + "              AND ISNULL(GLNum, '') != '')"
        sql = sql + "   BEGIN"
        sql = sql + "     SELECT 1"
        sql = sql + "   END"
        sql = sql + "   ELSE IF EXISTS(SELECT UseId"
        sql = sql + "                   FROM UTA_JourDtl(NOLOCK)"
        sql = sql + "                  WHERE AccSId ='" + accsid + "'"
        sql = sql + "                    AND UseId ='" + useid + "'"
        sql = sql + "                    AND IsTranBill != 0)"
        sql = sql + "   BEGIN"
        sql = sql + "    SELECT 2"
        sql = sql + "   END"
        sql = sql + "   ELSE"
        sql = sql + "   BEGIN"
        sql = sql + "    SELECT 0"
        sql = sql + "  END"
        sql = sql + " END"
        sql = sql + " ELSE"
        sql = sql + " BEGIN"
        sql = sql + "   SELECT @GLNum = GLNum,"
        sql = sql + "        @IsTranBill = IsTranBill"
        sql = sql + "    FROM UTA_JourDtl(NOLOCK)"
        sql = sql + "   WHERE AccSId ='" + accsid + "'"
        sql = sql + "     AND UseId ='" + useid + "'"
        sql = sql + "    AND Item ='" + item + "'"

        sql = sql + "   IF ISNULL(@GLNum, '') != ''"
        sql = sql + "   BEGIN"
        sql = sql + "     SELECT 1"
        sql = sql + "  END"
        sql = sql + "   ELSE IF @IsTranBill = 1"
        sql = sql + "   BEGIN"
        sql = sql + "     SELECT 2"
        sql = sql + "   END"
        sql = sql + "   BEGIN"
        sql = sql + "    SELECT 0"
        sql = sql + "   END"
        sql = sql + "  END "
        dt = db.FillDataSet(sql).Tables(0)

        flag = dt.Rows(0)(0).ToString()

        Return flag

    End Function

    '抓取傳票號碼
    Public Function FUN_GET_JourId(ByRef UseId As String, ByRef AccSId As String) As String
        Dim is_S As String
        is_S = " exec UTA_GetJourId '" & UseId & "','" & AccSId & "' "
        Return db.GetExecSQL(is_S)
    End Function

    '檢查是否鎖定
    Function CheckLock(ByVal YM As String) As Integer
        Dim is_S As String
        is_S = " select CostMonth " & _
               " from CAC_CostLock(nolock) " & _
               " where CostMonth ='" & YM & "'" & _
               " and isLock = 1"
        Return db.FillDataSet(is_S).Tables(0).Rows.Count
    End Function
    '檢查資料是否存在
    Function CheckYMexists(ByVal YM As String, ByVal tableName As String, ByVal Mesage As String, Optional Kind As Integer = 0) As Integer
        Dim is_S As String
        Dim SQLStringList As New ArrayList
        Dim Result As Integer = 0
        Dim RCount As Integer
        is_S = " select top 1 CostMonth " & _
               " from " & tableName & "(nolock)" & _
               " where CostMonth='" & YM & "'"
        RCount = db.FillDataSet(is_S).Tables(0).Rows.Count
        If Kind = 0 Then '有做刪除
            If RCount > 0 Then '有資料
                If MsgBox("警 告" & Chr(13) & Chr(10) & "此月份" & Mesage & "資料已存在您要刪除重新寫入嗎?", vbYesNo, "提示") = vbYes Then
                    SQLStringList.Add("delete " & tableName & _
                    " where CostMonth ='" & YM & "'")
                    db.ExecuteSqlTran(SQLStringList) '刪除后做重新寫入
                Else
                    Result = 1 '存在不重新寫入
                End If
            End If
        ElseIf Kind = 1 Then '無刪除
            If MsgBox("警 告" & Chr(13) & Chr(10) & "你確定要沿用上期" & Mesage & "的設定嗎?", vbYesNo, "提示") = vbYes Then
                Result = 1 '不沿用上期
            End If
        ElseIf Kind = 2 Then
            If MsgBox("警 告" & Chr(13) & Chr(10) & "本期" & Mesage & "未建立" & Chr(13) & Chr(10) & "設定你確定要沿用上期" & Mesage & "的設定嗎?", vbYesNo, "提示") = vbYes Then
                Result = 2
            End If
        Else
        End If
        Return Result
    End Function

    Function FindSort(ByVal Sort As Integer, Optional ByVal Kind As Integer = 0) As String
        Dim Result As String = ""
        If Kind = 0 Then
            Select Case Sort
                Case 0 : Result = " order by LineID,ProcCode,PartNum,Revision,Layer desc,ClassId"
                Case 1 : Result = " order by ProcCode,PartNum,Revision,Layer desc,ClassId"
                Case 2 : Result = " order by PartNum,Revision,Layer desc,ProcCode,ClassId"
                Case 3 : Result = " order by ClassId,PartNum,Revision,Layer desc,ProcCode"
            End Select
        ElseIf Kind = 4 Then
            Select Case Sort
                Case 0 : Result = " order by LineID,PartNum,Revision,Layer DESC,ClassId"
                Case 1 : Result = " order by PartNum,Revision,Layer DESC,ClassId"
                Case 2 : Result = " order by ClassId,PartNum,Revision,Layer DESC"
            End Select
        ElseIf Kind = 7 Then
            Select Case Sort
                Case 0 : Result = " order by LineId,ProcCode,PartNum,Revision,Layer DESC,ClassId"
                Case 1 : Result = " order by ProcCode,PartNum,Revision,Layer DESC,ClassId"
                Case 2 : Result = " order by PartNum,Revision,RoutSerial,Layer DESC,ProcCode,ClassId"
                Case 3 : Result = " order by ClassId,PartNum,Revision,Layer DESC,ProcCode"
                Case 4 : Result = " order by PartNum,Revision,RoutSerial,Layer DESC,ClassId"
            End Select
        End If
        Return Result
    End Function

    Function CheckContinue() As Boolean
        Dim Result = False
        If MsgBox("如資料量多可先設查詢條件以縮短顯示時間?" & Chr(13) & Chr(10) & "是否繼續顯示資料", vbYesNo, "提示") = vbYes Then
            Result = True
        End If
        Return Result
    End Function

    '繁體轉換簡體
    Function t2s(ByVal var_str As DataSet) As DataSet
        For i As Integer = 0 To var_str.Tables(0).Rows.Count - 1
            For j As Integer = 0 To var_str.Tables(0).Columns.Count - 1
                var_str.Tables(0).Rows(i)(j) = StrConv(var_str.Tables(0).Rows(i)(j).ToString, VbStrConv.SimplifiedChinese, 2052)
            Next
        Next
        Return var_str
    End Function
End Module
