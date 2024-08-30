Module mod_menu
    Public db As SqlDb = New SqlDb()


    Const TBNAME As String = "MENU"
    Const TFELD1 As String = "menutext"
    Const TFELD2 As String = "menutag"
    Const authcount As Integer = 20


    Structure PerMenu
        Dim menutext As ArrayList '存儲菜單列表
        Dim menutag As ArrayList '存儲菜單列表tag
    End Structure

    Structure LoginUser
        Dim comp_id As String
        Dim plan_id As String
        Dim emp_no As String
        Dim emp_name As String
        Dim UserID As String
        Dim IPID As String
        Dim Psno As String
        Dim tp As String
        Dim NETWORK_ACC As String
        Dim St As String
        Dim SystemID As String
        Dim USER_ID As String
        Dim GS_COMP_NAME As String
        Dim IS_SALES_ROLE As String
        Dim GS_RPT_PATH As String
        Dim GS_ACCESSARY_PATH As String
        Dim GS_MAIL_IP As String
        Dim GS_PC_IP As String
        Dim GS_MAIL_PORT As String
        Dim GS_RPT_CTITLE As String
        Dim GS_RPT_ETITLE As String
        Dim GS_ERP_MAIL_ADDR As String
        Dim GS_PAMSTR As String
        Dim GS_ROLE As String
        Dim GS_PC_ID As String
        Dim GS_PC_IP_STR As String
        Dim ROLE_TYPE As String
        Dim COST_NO As String
        Dim ITEMID As String
        Dim GS_DEP_ID As String
        Dim FROMFLAG As String
        Dim PUR_TYPE As String
        Dim GS_RIGHT As String
        Dim GS_REPORT_TITLE1 As String
        Dim GS_COMP_FULL_NAME As String
        Dim APPLY_COST_NO As String
        Dim IS_WUGUAN As String
        Dim IS_FIN As String
        Dim IS_CAIGOU As String
        Dim CIR_COMP As String
        Dim ERP_API As String
        Dim COMP_UM As String
        Dim GS_URL As String '線上簽核路勁
        Dim FROMname As String
        Dim ROLE_ID As String
        Dim CAM_FTP As String
        Dim ISVER As String
        Dim NKS_CAM As String
    End Structure
    Structure CompyInfo
        Dim Comp_Ftel As String
        Dim Comp_Ffax As String
        Dim Comp_Faddr As String
        Dim Comp_FullCname As String
        Dim Comp_FullEname As String
    End Structure

    Structure connectionString

        Dim con As String

    End Structure


    Structure StausInfo
        Dim cmd As String
    End Structure

    Structure WEBURLInfo
        Dim URL As String
        Dim RptURL As String
        Dim LoginUserId As String
        Dim PassWord As String
    End Structure

    Public IS_COMPY As New CompyInfo
    Public IS_USER As New LoginUser
    Public IS_Staus As New StausInfo
    Public WEBURL As New WEBURLInfo
    Public connection As New connectionString




    'p00	執行
    'p01	新增
    'p02	查詢
    'p03	更改
    'p04	刪除
    'p05	列印
    'p06	審核1
    'p07	審核2
    'p08	審核3
    'p09	核准1
    'p10	核准2
    'p11	核准3
    '做權限數組下標用,便於記憶
    Enum Rights_index
        Can_EXEC = 0
        Can_ADD = 1
        Can_SERH = 2
        Can_EDIT = 3
        Can_DEL = 4
        Can_PRINT = 5
        Can_CHECK1 = 6
        Can_CHECK2 = 7
        Can_CHECK3 = 8
        Can_VERIFY1 = 9
        Can_VERIFY2 = 10
        Can_VERIFY3 = 11
    End Enum


    Structure FORM_RIGHTS
        Dim Can_EXEC_P00 As Boolean
        Dim Can_ADD_P01 As Boolean
        Dim Can_SERH_P02 As Boolean
        Dim Can_EDIT_P03 As Boolean
        Dim Can_DEL_P04 As Boolean
        Dim Can_PRINT_P05 As Boolean
        Dim Can_CHECK1_P06 As Boolean
        Dim Can_CHECK2_P07 As Boolean
        Dim Can_CHECK3_P08 As Boolean
        Dim Can_VERIFY1_P09 As Boolean
        Dim Can_VERIFY2_P10 As Boolean
        Dim Can_VERIFY3_P11 As Boolean
    End Structure

    Dim mu As New PerMenu
    ',item= CASE ISNULL(URLAddress,'') WHEN '' THEN '' ELSE substring(URLAddress,charindex('/',URLAddress)+1,charindex('.aspx',URLAddress)-charindex('/',URLAddress)-1) END
    Function load_Mmenu(ByVal ms As MenuStrip, ByVal parentID As String)
        Dim sqlstr As String
        sqlstr = "select ItemID,ItemName" & _
                ",SystemID,NodeId from ACME.dbo.tbl_items WHERE  SystemID In (SELECT SystemID FROM sys_winfromSys) and NodeId=0 "
        With db.FillDataSet(sqlstr)
            If .Tables(0).Rows.Count > 0 Then
                With .Tables(0)
                    For i As Integer = 0 To .Rows.Count - 1
                        If .Rows(i).Item("NodeId") = parentID Then
                            Dim menuItem1 As New ToolStripMenuItem()
                            menuItem1.Text = .Rows(i).Item("ItemName")
                            menuItem1.Tag = .Rows(i).Item("ItemID")
                            ms.Items.Add(DirectCast(menuItem1, ToolStripItem))
                            CreateMenuItem(menuItem1, .Rows(i).Item("ItemID"))
                        End If
                    Next
                End With
            End If
        End With

    End Function
    ',item= CASE ISNULL(URLAddress,'') WHEN '' THEN '' ELSE substring(URLAddress,charindex('/',URLAddress)+1,charindex('.aspx',URLAddress)-charindex('/',URLAddress)-1) END
    Function CreateMenuItem(ByVal ms As ToolStripMenuItem, ByVal parentID As String)
        Dim sqlstr As String
        sqlstr = "select ItemID,ItemName=ItemName+CASE ISNULL(M,'') WHEN '' THEN '' ELSE '-'+ISNULL(M,'')  END " & _
                ",SystemID,NodeId from ACME.dbo.tbl_items WHERE SystemID In (SELECT SystemID FROM sys_winfromSys) and NodeId= '" & parentID & "'  order by  M"
        With db.FillDataSet(sqlstr)
            If .Tables(0).Rows.Count > 0 Then
                With .Tables(0)
                    For i As Integer = 0 To .Rows.Count - 1
                        Dim mitem As New ToolStripMenuItem()
                        mitem.Text = .Rows(i).Item("ItemName")
                        mitem.Tag = .Rows(i).Item("ItemID")
                        CreateMenuItem(mitem, .Rows(i).Item("ItemID"))
                        ms.DropDownItems.Add(mitem)
                    Next
                End With
            End If
        End With

    End Function

    '將權限轉化成列表
    ',item=CASE ISNULL(URLAddress,'') WHEN '' THEN '' ELSE substring(URLAddress,charindex('/',URLAddress)+1,charindex('.aspx',URLAddress)-charindex('/',URLAddress)-1) end
    Function loadlist2al(ByVal uid As String) As PerMenu
        Dim sqlstr As String
        Dim al As New PerMenu
        al.menutext = New ArrayList
        sqlstr = "select a.ItemID " & _
                 " from ACME.dbo.tbl_items a,ACME.dbo.tbl_userpermits b " & _
                 "where  a.SystemID In (SELECT SystemID FROM ACME.dbo.sys_winfromSys) and a.itemId=b.itemId and UserId='" & uid & "' " & _
                 " UNION ALL  SELECT ItemID FROM ACME.dbo.tbl_items WHERE ItemID IN (SELECT NodeId FROM ACME.dbo.tbl_items) " & _
                 " UNION ALL SELECT ItemID FROM ACME.dbo.tbl_i_groupitems WHERE groupid IN (SELECT groupid FROM ACME.dbo.tbl_i_usergroup  WHERE userid='" & uid & "') " & _
                 " UNION ALL select ItemID from ACME.dbo.tbl_items where '" & uid.ToUpper.Trim & "'='ADMIN' "
        With db.FillDataSet(sqlstr)
            If .Tables(0).Rows.Count > 0 Then
                With .Tables(0)
                    For i As Integer = 0 To .Rows.Count - 1
                        al.menutext.Add(.Rows(i).Item("ItemID"))
                    Next
                End With
            End If
        End With
        Return al

    End Function

    '===根據權限列表顯示某個登陸用戶的菜單
    'Public Sub loadmenu_m(ByVal uid As String, ByRef ms As MenuStrip)
    '    Dim al As New PerMenu
    '    al.menutext = New ArrayList
    '    '將菜單轉化成arraylist存儲
    '    al = loadlist2al(uid)

    '    For Each mi As ToolStripMenuItem In ms.Items
    '        '若菜單文本為'-'分隔符或者在菜單arraylist中能夠找到此菜單文本，則該菜單項可用
    '        If CBool(al.menutext.Contains(mi.Tag)) Then
    '            mi.Enabled = True
    '        Else
    '            mi.Enabled = False
    '        End If

    '        'If mi.Tag.ToString.Trim = "" Then mi.Enabled = True
    '        If mi.Enabled Then
    '            If mi.DropDownItems.Count > 0 Then
    '                loadmenu_s(mi, al)
    '            Else
    '                '動態添加事件
    '                AddHandler mi.Click, AddressOf Fm_main.MenuStrip1_Click
    '            End If
    '        End If
    '    Next
    'End Sub

    'Private Sub loadmenu_s(ByVal mi As ToolStripMenuItem, ByVal al As PerMenu)
    '    Dim mt As ToolStripMenuItem
    '    For Each mt In mi.DropDownItems
    '        mt.Enabled = CBool(al.menutext.Contains(mt.Tag))

    '        If mt.Tag.ToString.Trim = "" Then
    '            mt.Enabled = True
    '        End If

    '        If mt.Enabled Then
    '            If mt.DropDownItems.Count > 0 Then
    '                loadmenu_s(mt, al)
    '            Else
    '                AddHandler mt.Click, AddressOf Fm_main.MenuStrip1_Click
    '            End If
    '        End If
    '    Next
    'End Sub

    Private Sub SUB_CloseMenu_s(ByVal mi As MenuItem)
        Dim mt As MenuItem
        For Each mt In mi.MenuItems
            mt.Enabled = False
        Next
    End Sub

    '將菜單列表轉化成列表
    '------------------------------------------------------------------------------------------------------------------------------
    Private Sub Menu2AL(ByVal ms As MainMenu)
        Dim mi As MenuItem
        mu.menutext = New ArrayList
        mu.menutag = New ArrayList

        For Each mi In ms.MenuItems
            If mi.Text = "-" Then Continue For
            mu.menutext.Add(mi.Text)
            mu.menutag.Add(mi.Tag)

            If mi.MenuItems.Count > 0 Then
                menuitem(mi)
            End If
        Next
    End Sub

    Private Sub menuitem(ByVal mi As MenuItem)
        Dim mt As MenuItem
        For Each mt In mi.MenuItems
            mu.menutext.Add(mt.Text)
            mu.menutag.Add(mt.Tag)
            If mt.MenuItems.Count > 0 Then menuitem(mt)
        Next
    End Sub

    '將菜單展開成樹狀
    '參數 roottext:根節點顯示的文字
    Public Sub menu2tv(ByVal roottext As String, ByRef tv As TreeView, ByVal ms As MainMenu)
        tv.Nodes.Clear()
        '先增加一個根節點
        Dim tvroot As New TreeNode
        tvroot.Text = roottext
        tv.Nodes.Add(tvroot)

        Dim mi As MenuItem
        For Each mi In ms.MenuItems
            Dim nodet As New TreeNode
            nodet.Text = mi.Text
            nodet.Tag = mi.Tag
            tvroot.Nodes.Add(nodet)
            If mi.MenuItems.Count > 0 Then
                Dim nodes As New TreeNode
                Menu2tv_s(mi, nodet)
            End If
        Next

    End Sub

    Private Sub Menu2tv_s(ByVal mi As MenuItem, ByRef nodet As TreeNode)
        Dim mt As MenuItem
        For Each mt In mi.MenuItems
            Dim nodest As New TreeNode
            nodest.Text = mt.Text
            nodest.Tag = mt.Tag
            nodet.Nodes.Add(nodest)
            If mt.MenuItems.Count > 0 Then Menu2tv_s(mt, nodest)
        Next
    End Sub

    '將權限字符串轉化成為權限struct
    Function FUN_RIGHGSTR2RIGHTSTRU(ByVal VAR_RIGHTS As String) As FORM_RIGHTS
        Dim IS_STRU_RIGHTS As New FORM_RIGHTS

        IS_STRU_RIGHTS.Can_EXEC_P00 = Mid(VAR_RIGHTS, 1, 1) = "Y"
        IS_STRU_RIGHTS.Can_ADD_P01 = Mid(VAR_RIGHTS, 2, 1) = "Y"
        IS_STRU_RIGHTS.Can_SERH_P02 = Mid(VAR_RIGHTS, 3, 1) = "Y"
        IS_STRU_RIGHTS.Can_EDIT_P03 = Mid(VAR_RIGHTS, 4, 1) = "Y"
        IS_STRU_RIGHTS.Can_DEL_P04 = Mid(VAR_RIGHTS, 5, 1) = "Y"
        IS_STRU_RIGHTS.Can_PRINT_P05 = Mid(VAR_RIGHTS, 6, 1) = "Y"
        IS_STRU_RIGHTS.Can_CHECK1_P06 = Mid(VAR_RIGHTS, 7, 1) = "Y"
        IS_STRU_RIGHTS.Can_CHECK2_P07 = Mid(VAR_RIGHTS, 8, 1) = "Y"
        IS_STRU_RIGHTS.Can_CHECK3_P08 = Mid(VAR_RIGHTS, 9, 1) = "Y"
        IS_STRU_RIGHTS.Can_VERIFY1_P09 = Mid(VAR_RIGHTS, 10, 1) = "Y"
        IS_STRU_RIGHTS.Can_VERIFY2_P10 = Mid(VAR_RIGHTS, 11, 1) = "Y"
        IS_STRU_RIGHTS.Can_VERIFY3_P11 = Mid(VAR_RIGHTS, 12, 1) = "Y"
        Return IS_STRU_RIGHTS
    End Function

    '獲取某個員工，某個窗體的權限,2011.09.14 ADD
    '每個窗體的load事件中調用
    '參數為公司別，廠別，工號，程式名，返回的權限數組
    Sub FUN_GET_EMP_RIGHTS(ByVal VAR_UserId As String, ByVal VAR_itemId As String, ByRef VAR_RS() As Boolean)
        Dim str As String 'sql語句
        Dim pid As String '權限字符串
        Dim fname As String '臨時存字段名稱

        str = "select userId,itemId,systemid "
        For i As Integer = 0 To authcount
            fname = "P" & Right("00" & i, 2)
            str &= ",isnull(" & fname & ",0)  as " & fname
        Next
        str &= " from tbl_userpermits where (UserId='" & VAR_UserId & "' or '" & VAR_UserId.ToUpper & "'='ADMIN') and itemId = '" & VAR_itemId & "'"

        pid = ""
        With db.FillDataSet(str).Tables(0)
            If .Rows.Count > 0 Then
                For k As Integer = 0 To authcount
                    fname = "P" & Right("00" & k, 2)
                    VAR_RS(k) = CBool(.Rows(0).Item(fname).ToString)
                Next
            Else
                pid = ""
            End If
        End With
    End Sub

    '菜單點擊后寫入log,工號,系統代號,菜單代號,主機名稱
    Sub FUN_systemlog(ByVal UserId As String, ByVal systemId As String, ByVal itemId As String, ByVal hostname As String)
        Dim sql, memo As String
        memo = "winfrom系統登入"
        sql = "insert into acme.dbo.tbl_SysLogin_His ("
        sql &= " userid,modulid,"
        sql &= " itemid,host_name,"
        sql &= " memo"
        sql &= " ) values("
        sql &= " '" & UserId & "',"
        sql &= " '" & systemId & "',"
        sql &= " '" & itemId & "',"
        sql &= " '" & hostname & "',"
        sql &= " '" & memo & "'"
        sql &= ")"
        db.ExecuteSql(sql)

    End Sub
End Module
