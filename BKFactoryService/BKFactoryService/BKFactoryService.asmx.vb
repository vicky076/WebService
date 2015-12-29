Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Reflection

<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class BKFactoryService
    Inherits System.Web.Services.WebService

    Dim myDbfnc As New MyCardConn.Dbfuction
    Dim conn As New SqlClient.SqlConnection()

    Public Sub ConnStr(ByVal ConnNum As Integer)
        Select Case My.Settings.ConnectMode
            Case 1
                conn = myDbfnc.open100Db
            Case 2
                conn = myDbfnc.openTestDb
            Case 3
                Select Case ConnNum
                    Case 11
                        conn = myDbfnc.open11Db
                    Case 12
                        conn = myDbfnc.open12Db
                    Case 13
                        conn = myDbfnc.open13Db
                    Case 14
                        conn = myDbfnc.open14Db
                    Case 15
                        conn = myDbfnc.open15Db
                        'Case 16
                        '    conn = myDbfnc.open16Db
                    Case 17
                        conn = myDbfnc.open17Db
                    Case 19
                        conn = myDbfnc.open19Db
                    Case 20
                        conn = myDbfnc.open20Db
                End Select
            Case Else
                conn = myDbfnc.open100Db

        End Select
    End Sub

    ''' <summary>
    ''' Billing後台-付費廠商類別維護
    ''' 查詢
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetView_BK_StorePoints_FactoryType() As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線) 
        Dim sqlcomm As String
        sqlcomm = "select  Sn,FactoryTypeDesc from PointsBilling_DB.dbo.View_BK_StorePoints_FactoryType"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet()
        Try
            da.Fill(ds)

        Catch ex As Exception
            Throw New Exception("GetView_BK_StorePoints_FactoryType|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds
    End Function

    ''' <summary>
    ''' Billing後台-付費廠商類別維護
    ''' 新增付費廠商類別
    ''' </summary>
    ''' <param name="FactoryTypeDesc"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_FactoryType_Insert(ByVal FactoryTypeDesc As String, ByVal UserStamp As String) As InsertRE
        ConnStr(12)
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_FactoryType_Insert", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位
        cmd.Parameters.Add("@FactoryTypeDesc", SqlDbType.NVarChar, 30).Value = FactoryTypeDesc
        cmd.Parameters.Add("@UserId", SqlDbType.VarChar, 30).Value = UserStamp

        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnValue As New InsertRE
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("新增付費廠商類別失敗|SPS_StorePoints_FactoryType_Insert|" & ex.Message)
        Finally
            conn.Close()
        End Try

        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        Return ReturnValue

    End Function
    ''' <summary>
    ''' Billing後台-付費廠商類別維護
    ''' 修改付費廠商類別
    ''' </summary>
    ''' <param name="FactoryTypeDesc"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_FactoryType_Update(ByVal Sn As Integer, ByVal FactoryTypeDesc As String, ByVal UserStamp As String) As UpdateRE
        ConnStr(12)
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_FactoryType_Update", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位
        cmd.Parameters.Add("@SN", SqlDbType.Int, 4).Value = Sn
        cmd.Parameters.Add("@FactoryTypeDesc", SqlDbType.NVarChar, 30).Value = FactoryTypeDesc
        cmd.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserStamp


        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnValue As New UpdateRE
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("修改付費廠商類別失敗|SPS_StorePoints_FactoryType_Update|" & ex.Message)
        Finally
            conn.Close()
        End Try

        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        Return ReturnValue

    End Function

    ''' <summary>
    ''' Billing後台-IP黑名單管理
    ''' 查詢
    ''' </summary>
    ''' <param name="IP"></param>
    ''' <param name="Status"></param>
    ''' <param name="CheckIP"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetView_BK_StorePoints_BlackIPList(ByVal IP As String, ByVal Status As String, ByVal CheckIP As String) As DataSet
        ConnStr(12)
        Dim com As New SqlCommand()
        com.Connection = conn
        If IP = "" And Status = "-1" Then
            com.CommandText = "select distinct Sn,StartIP, EndIP, StartIPNum, EndIPNum, Status from PointsBilling_DB.dbo.View_BK_StorePoints_BlackIPList"
        ElseIf IP = "" And Status <> "-1" Then
            com.CommandText = "select distinct Sn,StartIP, EndIP, StartIPNum, EndIPNum, Status from PointsBilling_DB.dbo.View_BK_StorePoints_BlackIPList where Status=@Status"
            com.Parameters.Add("@Status", SqlDbType.Int, 1).Value = CInt(Status)
        ElseIf IP <> "" And Status = "-1" Then
            com.CommandText = "select distinct Sn,StartIP, EndIP, StartIPNum, EndIPNum, Status from PointsBilling_DB.dbo.View_BK_StorePoints_BlackIPList where @CheckIP between StartIPNum and EndIPNum"
            com.Parameters.Add("@CheckIP", SqlDbType.BigInt, 8).Value = CheckIP
        ElseIf IP <> "" And Status <> "-1" Then
            com.CommandText = "select distinct Sn,StartIP, EndIP, StartIPNum, EndIPNum, Status from PointsBilling_DB.dbo.View_BK_StorePoints_BlackIPList where Status=@Status and @CheckIP between StartIPNum and EndIPNum"
            com.Parameters.Add("@Status", SqlDbType.Int, 1).Value = CInt(Status)
            com.Parameters.Add("@CheckIP", SqlDbType.BigInt, 8).Value = CheckIP
        End If

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet()
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("GetView_BK_StorePoints_BlackIPList|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds
    End Function
    ''' <summary>
    ''' Billing後台-IP黑名單管理
    ''' 新增
    ''' </summary>
    ''' <param name="StartIP"></param>
    ''' <param name="EndIP"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_BlackIPList_Insert(ByVal StartIP As String, ByVal EndIP As String, ByVal ProcDesc As String, ByVal UserStamp As String) As InsertRE
        ConnStr(12)
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_BlackIPList_Insert", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位
        cmd.Parameters.Add("@StartIP", SqlDbType.VarChar, 25).Value = StartIP
        cmd.Parameters.Add("@EndIP", SqlDbType.VarChar, 25).Value = EndIP
        cmd.Parameters.Add("@ProcDesc", SqlDbType.VarChar, 30).Value = ProcDesc
        cmd.Parameters.Add("@CreateUser", SqlDbType.VarChar, 30).Value = UserStamp
        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnValue As New InsertRE
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("新增IP黑名單失敗|SPS_StorePoints_BlackIPList_Insert|" & ex.Message)
        Finally
            conn.Close()
        End Try

        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        Return ReturnValue

    End Function
    ''' <summary>
    ''' Billing後台-IP黑名單管理
    ''' 修改
    ''' </summary>
    ''' <param name="Sn"></param>
    ''' <param name="ProcDesc"></param>
    ''' <param name="Status"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_BlackIPList_Update(ByVal Sn As Integer, ByVal ProcDesc As String, ByVal Status As Integer, ByVal UserStamp As String) As UpdateRE
        ConnStr(12)
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_BlackIPList_Update", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位
        cmd.Parameters.Add("@SN", SqlDbType.Int, 4).Value = Sn
        cmd.Parameters.Add("@ProcDesc", SqlDbType.NVarChar, 30).Value = ProcDesc
        cmd.Parameters.Add("@Status", SqlDbType.Int, 1).Value = Status
        cmd.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserStamp
        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("修改IP黑名單失敗|SPS_StorePoints_BlackIPList_Update|" & ex.Message)
        Finally
            conn.Close()
        End Try
        Dim ReturnValue As New UpdateRE

        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        Return ReturnValue

    End Function
    ''' <summary>
    ''' Billing後台-IP黑名單管理
    ''' 依據修改按鈕狀態查詢
    ''' </summary>
    ''' <param name="AutoGenerateEdit">是否有修改按鈕</param>
    ''' <param name="StartIP"></param>
    ''' <param name="EndIP"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetView_BK_StorePoints_BlackIPListForEX(ByVal AutoGenerateEdit As Boolean, ByVal StartIP As String, ByVal EndIP As String) As DataSet
        ConnStr(12)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "select StartIP, EndIP,LogStatus,ProcDesc,CreateUser,CreateDate from PointsBilling_DB.dbo.View_BK_StorePoints_BlackIPList where @StartIP=StartIP and @EndIP=EndIP order by CreateDate"
        If AutoGenerateEdit = True Then
            com.Parameters.Add("@StartIP", SqlDbType.VarChar, 15).Value = StartIP
            com.Parameters.Add("@EndIP", SqlDbType.VarChar, 15).Value = EndIP
        Else
            com.Parameters.Add("@StartIP", SqlDbType.VarChar, 15).Value = StartIP
            com.Parameters.Add("@EndIP", SqlDbType.VarChar, 15).Value = EndIP
        End If
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet()
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("GetView_BK_StorePoints_BlackIPListForEX|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds
    End Function

    ''' <summary>
    ''' Billing後台-用戶號碼白名單管理
    ''' 查詢
    ''' </summary>
    ''' <param name="rdpStartDate">開始日期是否有值</param>
    ''' <param name="rdpEndDate">結束日期是否有值</param>
    ''' <param name="rdpStartDate_value">開始日期</param>
    ''' <param name="rdpEndDate_value">結束日期</param>
    ''' <param name="CustId">用戶號碼</param>
    ''' <param name="Status">狀態</param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetView_BK_StorePoints_WhiteCustIdList(ByVal rdpStartDate As Boolean, ByVal rdpEndDate As Boolean, ByVal rdpStartDate_value As String, ByVal rdpEndDate_value As String, ByVal CustId As String, ByVal Status As String) As DataSet
        ConnStr(12)
        Dim com As New SqlCommand()
        com.Connection = conn
        If rdpStartDate = False And rdpEndDate = False Then
            If CustId = "" And Status = "-1" Then
                com.CommandText = "select distinct Sn,CustId, DayAllowPrice, MonthAllowPirce, StartTime,EndTime,Status from PointsBilling_DB.dbo.View_BK_StorePoints_WhiteCustIdList order by CustId"

            ElseIf CustId = "" And Status <> "-1" Then
                com.CommandText = "select distinct Sn,CustId, DayAllowPrice, MonthAllowPirce, StartTime,EndTime,Status from PointsBilling_DB.dbo.View_BK_StorePoints_WhiteCustIdList where Status=@Status order by CustId"
                com.Parameters.Add("@Status", SqlDbType.Int, 1).Value = CInt(Status)

            ElseIf CustId <> "" And Status = "-1" Then
                com.CommandText = "select distinct Sn,CustId, DayAllowPrice, MonthAllowPirce, StartTime,EndTime,Status from PointsBilling_DB.dbo.View_BK_StorePoints_WhiteCustIdList where @CustId=CustId order by CustId"
                com.Parameters.Add("@CustId", SqlDbType.VarChar, 25).Value = CustId

            ElseIf CustId <> "" And Status <> "-1" Then
                com.CommandText = "select distinct Sn,CustId, DayAllowPrice, MonthAllowPirce, StartTime,EndTime,Status from PointsBilling_DB.dbo.View_BK_StorePoints_WhiteCustIdList where @CustId=CustId and Status=@Status order by CustId"
                com.Parameters.Add("@CustId", SqlDbType.VarChar, 25).Value = CustId
                com.Parameters.Add("@Status", SqlDbType.Int, 1).Value = CInt(Status)

            End If
        ElseIf rdpStartDate And rdpEndDate Then
            If CustId = "" And Status = "-1" Then
                com.CommandText = "select distinct Sn,CustId, DayAllowPrice, MonthAllowPirce, StartTime,EndTime,Status from PointsBilling_DB.dbo.View_BK_StorePoints_WhiteCustIdList where (@StartDate <= StartTime) and (@EndDate >= EndTime) order by CustId"
                com.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = rdpStartDate_value
                com.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = rdpEndDate_value
            ElseIf CustId = "" And Status <> "-1" Then
                com.CommandText = "select distinct Sn,CustId, DayAllowPrice, MonthAllowPirce, StartTime,EndTime,Status from PointsBilling_DB.dbo.View_BK_StorePoints_WhiteCustIdList where Status=@Status and (@StartDate <= StartTime) and (@EndDate >= EndTime) order by CustId"
                com.Parameters.Add("@Status", SqlDbType.Int, 1).Value = CInt(Status)
                com.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = rdpStartDate_value
                com.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = rdpEndDate_value
            ElseIf CustId <> "" And Status = "-1" Then
                com.CommandText = "select distinct Sn,CustId, DayAllowPrice, MonthAllowPirce, StartTime,EndTime,Status from PointsBilling_DB.dbo.View_BK_StorePoints_WhiteCustIdList where @CustId=CustId and (@StartDate <= StartTime) and (@EndDate >= EndTime) order by CustId"
                com.Parameters.Add("@CustId", SqlDbType.VarChar, 25).Value = CustId
                com.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = rdpStartDate_value
                com.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = rdpEndDate_value
            ElseIf CustId <> "" And Status <> "-1" Then
                com.CommandText = "select distinct Sn,CustId, DayAllowPrice, MonthAllowPirce, StartTime,EndTime,Status from PointsBilling_DB.dbo.View_BK_StorePoints_WhiteCustIdList where @CustId=CustId and Status=@Status and (@StartDate <= StartTime) and (@EndDate >= EndTime) order by CustId"
                com.Parameters.Add("@CustId", SqlDbType.VarChar, 25).Value = CustId
                com.Parameters.Add("@Status", SqlDbType.Int, 1).Value = CInt(Status)
                com.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = rdpStartDate_value
                com.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = rdpEndDate_value
            End If
        End If


        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet()
        Try
            da.Fill(ds)

        Catch ex As Exception
            Throw New Exception("GetView_BK_StorePoints_WhiteCustIdList|" & ex.Message)
        Finally
            conn.Close()
        End Try
        Return ds
    End Function
    ''' <summary>
    ''' Billing後台-用戶號碼白名單管理
    ''' 新增
    ''' </summary>
    ''' <param name="CustId"></param>
    ''' <param name="ProcDesc"></param>
    ''' <param name="DayAllowPrice"></param>
    ''' <param name="MonthAllowPirce"></param>
    ''' <param name="StartTime"></param>
    ''' <param name="EndTime"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_WhiteCustIdList_Insert(ByVal CustId As String, ByVal ProcDesc As String, ByVal DayAllowPrice As String, ByVal MonthAllowPirce As String, ByVal StartTime As Date, ByVal EndTime As Date, ByVal UserStamp As String) As InsertRE
        ConnStr(12)
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_WhiteCustIdList_Insert", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位
        cmd.Parameters.Add("@CustId", SqlDbType.VarChar, 25).Value = CustId
        cmd.Parameters.Add("@ProcDesc", SqlDbType.NVarChar, 30).Value = ProcDesc
        cmd.Parameters.Add("@DayAllowPrice", SqlDbType.Int).Value = DayAllowPrice
        cmd.Parameters.Add("@MonthAllowPirce", SqlDbType.Int).Value = MonthAllowPirce
        cmd.Parameters.Add("@StartTime", SqlDbType.DateTime).Value = StartTime
        cmd.Parameters.Add("@EndTime", SqlDbType.DateTime).Value = EndTime
        cmd.Parameters.Add("@CreateUser", SqlDbType.VarChar, 30).Value = UserStamp
        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnValue As New InsertRE
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("新增用戶號碼白名單|SPS_StorePoints_BlackIPList_Insert|" & ex.Message)
        Finally
            conn.Close()
        End Try

        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        Return ReturnValue

    End Function
    ''' <summary>
    ''' Billing後台-用戶號碼白名單管理
    ''' 修改
    ''' </summary>
    ''' <param name="Sn"></param>
    ''' <param name="ProcDesc"></param>
    ''' <param name="DayAllowPrice"></param>
    ''' <param name="MonthAllowPirce"></param>
    ''' <param name="StartTime"></param>
    ''' <param name="EndTime"></param>
    ''' <param name="Status"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_WhiteCustIdList_Update(ByVal Sn As Integer, ByVal ProcDesc As String, ByVal DayAllowPrice As String, ByVal MonthAllowPirce As String, ByVal StartTime As Date, ByVal EndTime As Date, ByVal Status As String, ByVal UserStamp As String) As UpdateRE
        ConnStr(12)
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_WhiteCustIdList_Update", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位
        cmd.Parameters.Add("@SN", SqlDbType.Int, 4).Value = Sn
        cmd.Parameters.Add("@ProcDesc", SqlDbType.NVarChar, 30).Value = ProcDesc
        cmd.Parameters.Add("@DayAllowPrice", SqlDbType.Int).Value = DayAllowPrice
        cmd.Parameters.Add("@MonthAllowPirce", SqlDbType.Int).Value = MonthAllowPirce
        cmd.Parameters.Add("@StartTime", SqlDbType.DateTime).Value = StartTime
        cmd.Parameters.Add("@EndTime", SqlDbType.DateTime).Value = EndTime
        cmd.Parameters.Add("@Status", SqlDbType.Int, 1).Value = Status
        cmd.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserStamp
        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnValue As New UpdateRE
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("修改用戶號碼白名單|SPS_StorePoints_WhiteCustIdList_Update|" & ex.Message)
        Finally
            conn.Close()
        End Try

        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        Return ReturnValue

    End Function
    ''' <summary>
    ''' Billing後台-用戶號碼白名單管理
    ''' 依據修改按鈕狀態查詢
    ''' </summary>
    ''' <param name="AutoGenerateEdit">是否有修改按鈕</param>
    ''' <param name="CustId"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetView_BK_StorePoints_WhiteCustIdListForEX(ByVal AutoGenerateEdit As Boolean, ByVal CustId As String) As DataSet
        ConnStr(12)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "select CustId,LogStatus,ProcDesc,CreateUser,CreateDate from PointsBilling_DB.dbo.View_BK_StorePoints_WhiteCustIdList where @CustId=CustId order by CreateDate"
        If AutoGenerateEdit = True Then
            com.Parameters.Add("@CustId", SqlDbType.VarChar, 25).Value = CustId
        Else
            com.Parameters.Add("@CustId", SqlDbType.VarChar, 25).Value = CustId
        End If

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet()
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("GetView_BK_StorePoints_WhiteCustIdListForEX|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds
    End Function
    ''' <summary>
    ''' 20140123 qq 把判斷抓取特定廠商移除，可以查全部廠商
    ''' Billing後台-廠商結帳配合條件設定
    ''' 依權限GameServiceId取得特定廠商
    ''' </summary>
    ''' <param name="GameServiceId"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_GameFactorySP_View(ByVal GameServiceId As String) As DataSet
        ConnStr(14)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "select Distinct GameFactoryId,GameFactoryName from Mycard_backup.dbo.swvw_GetGameFactoryJoinGameService_PointBilling order by GameFactoryName"
        'com.CommandText = "select Distinct GameFactoryId,GameFactoryName from Mycard_backup.dbo.swvw_GetGameFactoryJoinGameService_PointBilling where GameServiceId in (" & GameServiceId & ") order by GameFactoryName"
        com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 20).Value = GameServiceId
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢特定遊戲廠商失敗|Get_GameFactorySP_View|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds
    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 取得遊戲廠商資料
    ''' </summary>
    ''' <param name="GameFactoryId"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_GameFactory_View(ByVal GameFactoryId As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線) 
        'Dim sqladp As New System.Data.SqlClient.SqlDataAdapter("Select GameFactoryId,GameFactoryName,Status From PointsBilling_DB.dbo.View_BK_StorePoints_GameFactory where GameFactoryId=@GameFactoryId And Status=1", cn)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "Select GameFactoryId,GameFactoryName,HwaId,Status From PointsBilling_DB.dbo.View_BK_StorePoints_GameFactory where GameFactoryId=@GameFactoryId And Status=1"
        com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢遊戲廠商失敗|Get_GameFactory_View|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function
    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 取得遊戲服務
    ''' </summary>
    ''' <param name="GameFactoryId"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_GameService_View(ByVal GameFactoryId As String, ByVal hidGSerId As String) As DataSet
        ConnStr(14)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "Select Distinct GameServiceId,GameName From Mycard_backup.dbo.swvw_GetGameFactoryJoinGameService_PointBilling where GameFactoryId=@GameFactoryId"
        'com.CommandText = "Select Distinct GameServiceId,GameName From Mycard_backup.dbo.swvw_GetGameFactoryJoinGameService_PointBilling where GameServiceId in (" & hidGSerId & ") And GameFactoryId=@GameFactoryId"
        com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢遊戲服務失敗|Get_GameService_View|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function
    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 取得金流服務
    ''' </summary>
    ''' <param name="GameServiceId"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_FacService_View(ByVal GameServiceId As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandType = CommandType.StoredProcedure
        com.CommandText = "PointsBilling_DB.dbo.SPS_StorePoints_CPFactorySharedRate_Query"
        com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢付費方式失敗|Get_FacService_View|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function
    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 取得產品服務
    ''' </summary>
    ''' <param name="GameFactoryId"></param>
    ''' <param name="GameServiceId"></param>
    ''' <param name="FactoryId"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_ProductService_View(ByVal GameFactoryId As String, ByVal GameServiceId As String, ByVal FactoryId As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandTimeout = 300
        com.CommandText = "Select Sn,FactoryId,FactoryName,FactoryServiceName,GameProductName,Price,SharedRate,GameFactoryId,GameServiceId,Status From PointsBilling_DB.dbo.View_BK_StorePoints_StoreProductService where GameFactoryId=@GameFactoryId And GameServiceId=@GameServiceId And FactoryId=@FactoryId And Status=1 order by FactoryServiceName"
        com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId
        com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢產品服務失敗|Get_ProductService_View|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function
    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 修改
    ''' </summary>
    ''' <param name="Sn"></param>
    ''' <param name="SharedRate"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg 
    ''' </remarks>
    <WebMethod()> _
    Public Function ProductService_Update(ByVal Sn As Integer, ByVal SharedRate As Double, ByVal UserStamp As String) As ProductServiceUpdateRE
        ConnStr(12)

        Dim Cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_StoreProductServiceSharedRate_Update", conn)

        Cmd.CommandType = CommandType.StoredProcedure

        Cmd.Parameters.Add("@Sn", SqlDbType.Int).Value = Sn

        Cmd.Parameters.Add("@SharedRate", SqlDbType.Decimal).Value = SharedRate
        Cmd.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserStamp

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ProductServiceUpdateRE
        Try
            conn.Open()
            Cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("攤分比更新失敗|SPS_StorePoints_StoreProductServiceSharedRate_Update|" & ex.Message)
        Finally
            conn.Close()
        End Try

        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台-線上購卡遊戲服務分類管理 
    ''' 查詢
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetView_StorePoints_GameServiceType() As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線) 
        Dim sqlcomm As String
        sqlcomm = "select Sn,GameServiceTypeDesc,Status,Priority from PointsBilling_DB.dbo.View_StorePoints_GameServiceType"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet()
        Try
            da.Fill(ds)

        Catch ex As Exception
            Throw New Exception("GetView_StorePoints_GameServiceType|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds
    End Function
    ''' <summary>
    ''' Billing後台-線上購卡遊戲服務分類管理
    ''' 新增
    ''' </summary>
    ''' <param name="GameServiceTypeDesc"></param>
    ''' <param name="Status"></param>
    ''' <param name="Priority"></param>
    ''' <param name="CreateUser"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_GameServiceType_Insert(ByVal GameServiceTypeDesc As String, ByVal Status As Integer, ByVal Priority As Integer, ByVal CreateUser As String) As InsertRE
        ConnStr(12)
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_GameServiceType_Insert", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位
        cmd.Parameters.Add("@GameServiceTypeDesc", SqlDbType.NVarChar, 30).Value = GameServiceTypeDesc
        cmd.Parameters.Add("@Status", SqlDbType.Int).Value = Status
        cmd.Parameters.Add("@Priority", SqlDbType.Int).Value = Priority
        cmd.Parameters.Add("@CreateUser", SqlDbType.VarChar, 30).Value = CreateUser

        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnValue As New InsertRE
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("新增類別失敗|SPS_StorePoints_GameServiceType_Insert|" & ex.Message)
        Finally
            conn.Close()
        End Try

        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        Return ReturnValue

    End Function
    ''' <summary>
    ''' Billing後台-線上購卡遊戲服務分類管理
    ''' 修改
    ''' </summary>
    ''' <param name="Sn"></param>
    ''' <param name="GameServiceTypeDesc"></param>
    ''' <param name="Status"></param>
    ''' <param name="Priority"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_GameServiceType_Update(ByVal Sn As Integer, ByVal GameServiceTypeDesc As String, ByVal Status As Integer, ByVal Priority As Integer, ByVal UserStamp As String) As UpdateRE
        ConnStr(12)
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_GameServiceType_Update", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位
        cmd.Parameters.Add("@Sn", SqlDbType.Int, 4).Value = Sn
        cmd.Parameters.Add("@GameServiceTypeDesc", SqlDbType.NVarChar, 30).Value = GameServiceTypeDesc
        cmd.Parameters.Add("@Status", SqlDbType.Int).Value = Status
        cmd.Parameters.Add("@Priority", SqlDbType.Int).Value = Priority
        cmd.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserStamp
        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnValue As New UpdateRE
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("修改類別失敗|SPS_StorePoints_GameServiceType_Update|" & ex.Message)
        Finally
            conn.Close()
        End Try

        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        Return ReturnValue

    End Function
    ''' <summary>
    ''' Billing後台-Billing銷售報表
    ''' 查詢
    ''' </summary>
    ''' <param name="FactoryId"></param>
    ''' <param name="DeptSn">單位別</param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetView_BK_Billing_FactoryService(ByVal FactoryId As String, ByVal DeptSn As Integer) As DataSet
        ConnStr(14)
        Dim com As New SqlCommand()
        com.Connection = conn
        If FactoryId = "" Then
            com.CommandText = "Select distinct FactoryId,FactoryName From PointsBilling_DB.dbo.View_BK_Billing_FactoryService_dept where DeptSn=@DeptSn"
            com.Parameters.Add("@DeptSn", SqlDbType.Int).Value = DeptSn
        Else
            com.CommandText = "Select distinct FactoryServiceId,FactoryServiceName From PointsBilling_DB.dbo.View_BK_Billing_FactoryService_dept where FactoryId=@FactoryId and DeptSn=@DeptSn"
            com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
            com.Parameters.Add("@DeptSn", SqlDbType.Int).Value = DeptSn
        End If

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet()
        Try
            da.Fill(ds)

        Catch ex As Exception
            Throw New Exception("GetddlFacServiceView|" & ex.Message)
        Finally
            conn.Close()
        End Try
        Return ds

    End Function
    ''' <summary>
    ''' Billing後台-Billing銷售報表
    ''' 查詢報表
    ''' </summary>
    ''' <param name="StartDate"></param>
    ''' <param name="EndDate"></param>
    ''' <param name="FactoryId"></param>
    ''' <param name="FactoryServiceId"></param>
    ''' <param name="ReportPriority"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_SPS_StorePoints_BillingOnlineSellReport(ByVal StartDate As String, ByVal EndDate As String, ByVal FactoryId As String, ByVal FactoryServiceId As String, ByVal ReportPriority As Integer, ByVal DeptSn As Integer) As BillingOnlineSellReportRE
        ConnStr(14)
        Dim Cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_BillingOnlineSellReport", conn)
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 480
        Cmd.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
        Cmd.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
        Cmd.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        Cmd.Parameters.Add("@FactoryServiceId ", SqlDbType.VarChar, 30).Value = FactoryServiceId
        Cmd.Parameters.Add("@ReportPriority", SqlDbType.Int, 1).Value = ReportPriority
        Cmd.Parameters.Add("@DeptSN", SqlDbType.Int).Value = DeptSn

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New BillingOnlineSellReportRE
        Dim da As New SqlDataAdapter(Cmd)
        Dim ds As New DataSet()
        Try
            da.Fill(ds)
        Catch ex As Exception

            Throw New Exception("查詢報表資料失敗|Get_BillingOnlineSellReport|" & ex.Message)

        Finally
            conn.Close()
        End Try
        If ReturnMsgNo.Value = 1 Then
            ReturnValue.ReturnDS = ds

        End If
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value

        Return ReturnValue
    End Function
    ''' <summary>
    ''' Billing後台-CP對應付費方式批次新增
    ''' 取得遊戲廠商
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_GameFactory_ForDDL() As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線) 
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "Select GameFactoryId,GameFactoryName From PointsBilling_db.dbo.View_BK_StorePoints_GameFactory order by GameFactoryName"

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢遊戲廠商失敗|Get_GameFactory_ForDDL|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function
    ''' <summary>
    ''' Billing後台-CP對應付費方式批次新增
    ''' 取得遊戲廠商(區分單位別)
    ''' </summary>
    ''' <param name="DeptSn">單位別</param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_GameFactory_ForDDL_Dept(ByVal DeptSn As Integer) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線) 
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "Select GameFactoryId,GameFactoryName From PointsBilling_db.dbo.View_BK_StorePoints_GameFactory_dept where DeptSn=@DeptSn order by GameFactoryName"
        com.Parameters.Add("@DeptSn", SqlDbType.Int).Value = DeptSn

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢遊戲廠商失敗|Get_GameFactory_ForDDL|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function
    ''' <summary>
    ''' Billing後台-CP對應付費方式批次新增
    ''' 取得遊戲服務
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_GameService_ForDDL(ByVal GameFactoryId As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "Select GameServiceId,GameName From PointsBilling_db.dbo.View_BK_StorePoints_GameService where GameFactoryId=@GameFactoryId order by GameName"
        com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢遊戲服務失敗|Get_GameService_ForDDL|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function
    ''' <summary>
    ''' Billing後台-CP對應付費方式批次新增
    ''' 取得付費方式
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_Factory_View(ByVal DeptSn As Integer) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "Select distinct FactoryId,FactoryName From PointsBilling_db.dbo.View_BK_StorePoints_FactoryLimit where DeptSn=@DeptSn order by FactoryName"
        com.Parameters.Add("@DeptSN", SqlDbType.Int).Value = DeptSn
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢付費方式失敗|Get_Factory_View|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function
    ''' <summary>
    ''' Billing後台-CP對應付費方式批次新增
    ''' 取得額度
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_Bound_View(ByVal FactoryId As String, ByVal DeptSn As Integer) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "Select * From PointsBilling_db.dbo.View_BK_StorePoints_FactoryLimit where FactoryId=@FactoryId and DeptSn=@DeptSn"
        com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        com.Parameters.Add("@DeptSN", SqlDbType.Int).Value = DeptSn
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢額度失敗|Get_Bound_View|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function
    ''' <summary>
    ''' Billing後台-CP對應付費方式批次新增
    ''' 新增
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function CPAndFactoryLimit_Insert(ByVal FactoryId As String, ByVal GameFactoryId As String, ByVal GameServiceId As String, ByVal UpperFristBound As Integer, ByVal DayUpperMobileBound As Integer, ByVal MonthUpperMobileBound As Integer, ByVal DayUpperMailBound As Integer, ByVal MonthUpperMailBound As Integer, ByVal DayUpperCustIdBound As Integer, ByVal MonthUpperCustIdBound As Integer, ByVal TempCodeLimit As Integer, ByVal IPTradeLimit As Integer, ByVal IncludeBadAccountStatus As Integer, ByVal CreateUser As String) As ReturnRE
        ConnStr(12)
        Dim Cmd As New SqlClient.SqlCommand("PointsBilling_db.dbo.SPS_StorePoints_CPAndFactoryLimit_Insert", conn)
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 100

        Cmd.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        Cmd.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId
        Cmd.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        Cmd.Parameters.Add("@UpperFristBound", SqlDbType.Int).Value = UpperFristBound
        Cmd.Parameters.Add("@DayUpperMobileBound", SqlDbType.Int).Value = DayUpperMobileBound
        Cmd.Parameters.Add("@MonthUpperMobileBound", SqlDbType.Int).Value = MonthUpperMobileBound
        Cmd.Parameters.Add("@DayUpperMailBound", SqlDbType.Int).Value = DayUpperMailBound
        Cmd.Parameters.Add("@MonthUpperMailBound", SqlDbType.Int).Value = MonthUpperMailBound
        Cmd.Parameters.Add("@DayUpperCustIdBound", SqlDbType.Int).Value = DayUpperCustIdBound
        Cmd.Parameters.Add("@MonthUpperCustIdBound", SqlDbType.Int).Value = MonthUpperCustIdBound
        Cmd.Parameters.Add("@TempCodeLimit", SqlDbType.Int).Value = TempCodeLimit
        Cmd.Parameters.Add("@IPTradeLimit", SqlDbType.Int).Value = IPTradeLimit
        Cmd.Parameters.Add("@IncludeBadAccountStatus", SqlDbType.Int).Value = IncludeBadAccountStatus
        Cmd.Parameters.Add("@CreateUser", SqlDbType.VarChar, 30).Value = CreateUser

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnRE
        Try
            conn.Open()
            Cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("CPAndFactoryLimit_Insert|" & ex.Message)
        Finally
            conn.Close()
        End Try

        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        Return ReturnValue

    End Function
    ''' <summary>
    ''' Billing後台-billing流程總失敗比率查詢
    ''' 取得遊戲廠商
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_GameFactory_ForDDL14() As DataSet
        ConnStr(14)
        ' 建立命令(預存程序名,使用連線) 
        Dim sqladp As New System.Data.SqlClient.SqlDataAdapter("Select GameFactoryId,GameFactoryName From PointsBilling_db.dbo.View_BK_StorePoints_GameFactory", conn)
        Dim da As New DataSet
        Try
            sqladp.Fill(da)
        Catch ex As Exception
            Throw New Exception("查詢遊戲廠商失敗|Get_GameFactory_ForDDL|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return da

    End Function

    ''' <summary>
    ''' Billing後台-billing流程總失敗比率查詢
    ''' 取得遊戲服務
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_GameService_ForDDL14(ByVal GameFactoryId As String) As DataSet
        ConnStr(14)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "Select GameServiceId,GameName From PointsBilling_db.dbo.View_BK_StorePoints_GameService where GameFactoryId=@GameFactoryId"
        com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢遊戲服務失敗|Get_GameService_ForDDL|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function

    ''' <summary>
    ''' Billing後台-billing流程總失敗比率查詢
    ''' 取得遊戲服務
    ''' </summary>
    ''' <returns>DataSet,ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳DataSet,ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function BillingTradeFailQuery(ByVal GameFactoryId As String, ByVal GameServiceId As String, ByVal StarDate As String, ByVal EndDate As String, ByVal DeptSn As Integer) As TradeFailRE
        ConnStr(14)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_db.dbo.SPS_StorePoints_BillingTradeFailQuery"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 300
        com.CommandType = CommandType.StoredProcedure
        '輸入參數值
        com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId
        com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 20).Value = GameServiceId
        com.Parameters.Add("@StarDate", SqlDbType.DateTime).Value = StarDate
        com.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
        com.Parameters.Add("@DeptSN", SqlDbType.Int).Value = DeptSn

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New TradeFailRE
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("BillingTradeFailQuery|" & ex.Message)
        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        ReturnValue.ReturnDS = ds

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台-雙向簡訊報表
    ''' 遊戲廠商
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_GameFactory_View_all() As DataSet
        ConnStr(12)
        Dim com As New SqlClient.SqlCommand
        com.Connection = conn
        com.CommandType = CommandType.Text
        com.CommandText = "Select GameFactoryId,GameFactoryName From PointsBilling_DB.dbo.View_BK_StorePoints_GameFactory order by GameFactoryName"

        ' 建立命令(預存程序名,使用連線) 
        Dim sqladp As New System.Data.SqlClient.SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            sqladp.Fill(ds)
        Catch ex As Exception
            Throw New Exception("Get_GameFactory_View|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds
    End Function
    ''' <summary>
    ''' Billing後台-雙向簡訊報表
    ''' 找遊戲廠商Value
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_GameFactory_ValueByText(ByVal GameFactoryName As String) As String
        ConnStr(12)
        Dim com As New SqlClient.SqlCommand
        com.Connection = conn
        com.CommandType = CommandType.Text
        com.CommandText = "Select GameFactoryId,GameFactoryName From PointsBilling_DB.dbo.View_BK_StorePoints_GameFactory where GameFactoryName=@GameFactoryName order by GameFactoryName"
        com.Parameters.Add("@GameFactoryName", SqlDbType.NVarChar, 30).Value = GameFactoryName
        ' 建立命令(預存程序名,使用連線) 
        Dim sqladp As New System.Data.SqlClient.SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            sqladp.Fill(ds)
        Catch ex As Exception
            Throw New Exception("Get_GameFactory_ViewByText|" & ex.Message)
        Finally
            conn.Close()
        End Try
        Dim GameFactoryValue As String
        If ds.Tables(0).Rows.Count > 0 Then
            GameFactoryValue = ds.Tables(0).Rows(0).Item("GameFactoryId")
        Else
            GameFactoryValue = ""
        End If
        Return GameFactoryValue
    End Function
    ''' <summary>
    ''' Billing後台-雙向簡訊報表
    ''' 找遊戲服務
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_GameService_View_ByGameFactoryId(ByVal GameFactoryId As String) As DataSet
        ConnStr(12)
        Dim com As New SqlClient.SqlCommand
        com.Connection = conn
        com.CommandType = CommandType.Text
        com.CommandText = "Select GameFactoryId,GameServiceId,GameName From PointsBilling_DB.dbo.View_BK_StorePoints_GameService where GameFactoryId=@GameFactoryId"
        com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId
        ' 建立命令(預存程序名,使用連線) 
        Dim sqladp As New System.Data.SqlClient.SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            sqladp.Fill(ds)
        Catch ex As Exception
            Throw New Exception("Get_GameService_View|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds
    End Function
    ''' <summary>
    ''' Billing後台-雙向簡訊報表
    ''' 付費方式
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_StorePoints_Factory_View() As DataSet
        ConnStr(12)
        Dim com As New SqlClient.SqlCommand
        com.Connection = conn
        com.CommandType = CommandType.Text
        com.CommandText = "Select FactoryId,FactoryName From PointsBilling_DB.dbo.View_BK_StorePoints_Factory_ForTwoWayMsg order by FactoryName"

        ' 建立命令(預存程序名,使用連線) 
        Dim sqladp As New System.Data.SqlClient.SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            sqladp.Fill(ds)
        Catch ex As Exception
            Throw New Exception("Get_StorePoints_Factory_View|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds
    End Function
    ''' <summary>
    ''' Billing後台-雙向簡訊報表
    ''' 
    ''' </summary>
    ''' <returns>DataSet,ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳DataSet,ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function TwoWayMsg_SecondMsgRecord_ForReport(ByVal StartDate As String, ByVal EndDate As String, ByVal TradeSeq As String, ByVal CustMobile As String, ByVal GameFactoryName As String, ByVal FactoryName As String, ByVal PayStatus As Integer, ByVal GameName As String, ByVal FactoryServiceName As String, ByVal ReportType As Integer) As ReportRE
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_TwoWayMsg_SecondMsgRecord_ForReport_Search"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure
        '輸入參數值
        com.Parameters.Add("@StartDate", SqlDbType.VarChar, 20).Value = StartDate
        com.Parameters.Add("@EndDate", SqlDbType.VarChar, 20).Value = EndDate
        com.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 16).Value = TradeSeq
        com.Parameters.Add("@CustMobile", SqlDbType.VarChar, 30).Value = CustMobile
        com.Parameters.Add("@GameFactoryId", SqlDbType.NVarChar, 30).Value = GameFactoryName
        com.Parameters.Add("@EC_FactoryId", SqlDbType.NVarChar, 30).Value = FactoryName
        com.Parameters.Add("@PayStatus", SqlDbType.Int).Value = PayStatus
        com.Parameters.Add("@GameServiceId", SqlDbType.NVarChar, 30).Value = GameName
        com.Parameters.Add("@FactoryID", SqlDbType.NVarChar, 30).Value = FactoryServiceName
        com.Parameters.Add("@ReportType", SqlDbType.Int).Value = ReportType

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReportRE
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception

            Throw New Exception(ex.Message)

        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        ReturnValue.ReturnDS = ds

        Return ReturnValue
    End Function
    ''' <summary>
    ''' Billing後台-雙向簡訊報表
    ''' 取得遊戲廠商AJAX
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_GameFactory_View_AJAX(ByVal prefixText As String) As DataSet
        ConnStr(12)
        Dim com As New SqlClient.SqlCommand("Select GameFactoryId,GameFactoryName From PointsBilling_DB.dbo.View_BK_StorePoints_GameFactory Where GameFactoryName like @GameFactoryName order by GameFactoryName", conn)
        com.Parameters.AddWithValue("@GameFactoryName", prefixText.Trim + "%")

        ' 建立命令(預存程序名,使用連線) 
        Dim sqladp As New System.Data.SqlClient.SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            sqladp.Fill(ds)
        Catch ex As Exception

            Throw New Exception("Get_GameFactory_View_AJAX|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds
    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 查詢拆帳比
    ''' </summary>
    ''' <param name="QueryType">1:NOW 2:Futrue 3:history</param>
    ''' <param name="GameFactoryId"></param>
    ''' <param name="GameServiceId"></param>
    ''' <param name="FactoryId"></param>
    ''' <returns>DataSet,ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳DataSet,ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_StoreProductServiceSharedRate_Query(ByVal QueryType As Integer, ByVal GameFactoryId As String, ByVal GameServiceId As String, ByVal FactoryId As String) As ReportRE
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_StoreProductServiceSharedRate_Query"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure
        com.Parameters.Add("@QueryType", SqlDbType.Int).Value = QueryType
        com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId
        com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 20).Value = GameServiceId
        com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReportRE
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception

            Throw New Exception(ex.Message)

        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
        ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
        ReturnValue.ReturnDS = ds

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 新增產品服務攤分比
    ''' </summary>
    ''' <param name="UpdateType">1:NOW 2:Futrue</param>
    ''' <param name="Sn"></param>
    ''' <param name="StoreProductServiceSn">產品服務SN</param>
    ''' <param name="SharedRate">攤分比</param>
    ''' <param name="StartDate"></param>
    ''' <param name="EndDate"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_StoreProductServiceSharedRate_Insert(ByVal UpdateType As Integer, ByVal Sn As Integer, ByVal StoreProductServiceSn As Integer, ByVal SharedRate As Double, ByVal StartDate As Date, ByVal EndDate As Date, ByVal UserStamp As String) As ReturnRE
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_StoreProductServiceSharedRate_Insert"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure
        com.Parameters.Add("@UpdateType", SqlDbType.Int).Value = UpdateType
        com.Parameters.Add("@Sn", SqlDbType.Int).Value = Sn
        com.Parameters.Add("@StoreProductServiceSn", SqlDbType.Int).Value = StoreProductServiceSn
        com.Parameters.Add("@SharedRate", SqlDbType.Decimal).Value = SharedRate
        com.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
        com.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
        com.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserStamp

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnRE
        Try
            conn.Open()
            com.ExecuteNonQuery()
        Catch ex As Exception

            Throw New Exception(ex.Message)

        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 修改產品服務攤分比
    ''' </summary>
    ''' <param name="UpdateType">1:NOW 2:Futrue</param>
    ''' <param name="Sn"></param>
    ''' <param name="StoreProductServiceSn">產品服務SN</param>
    ''' <param name="SharedRate">攤分比</param>
    ''' <param name="StartDate"></param>
    ''' <param name="EndDate"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_StoreProductServiceSharedRate_Update(ByVal UpdateType As Integer, ByVal Sn As Integer, ByVal StoreProductServiceSn As Integer, ByVal SharedRate As Double, ByVal StartDate As String, ByVal EndDate As String, ByVal UserStamp As String) As ReturnRE
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_StoreProductServiceSharedRate_Update"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure
        com.Parameters.Add("@UpdateType", SqlDbType.Int).Value = UpdateType
        com.Parameters.Add("@Sn", SqlDbType.Int).Value = Sn
        com.Parameters.Add("@StoreProductServiceSn", SqlDbType.Int).Value = StoreProductServiceSn
        com.Parameters.Add("@SharedRate", SqlDbType.Decimal).Value = SharedRate
        com.Parameters.Add("@StartDate", SqlDbType.VarChar, 25).Value = StartDate
        com.Parameters.Add("@EndDate", SqlDbType.VarChar, 25).Value = EndDate
        com.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserStamp

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnRE
        Try
            conn.Open()
            com.ExecuteNonQuery()
        Catch ex As Exception

            Throw New Exception(ex.Message)

        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 修改產品服務攤分比 20120801更新
    ''' </summary>
    ''' <param name="Sn"></param>
    ''' <param name="StoreProductServiceSn">產品服務SN</param>
    ''' <param name="SharedRate">攤分比</param>
    ''' <param name="StartDate"></param>
    ''' <param name="EndDate"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_StoreProductServiceSharedRate_Update_New(ByVal Sn As Integer, ByVal StoreProductServiceSn As Integer, ByVal SharedRate As Double, ByVal StartDate As String, ByVal EndDate As String, ByVal UserStamp As String) As ReturnRE
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_StoreProductServiceSharedRate_Update"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure
        com.Parameters.Add("@Sn", SqlDbType.Int).Value = Sn
        com.Parameters.Add("@StoreProductServiceSn", SqlDbType.Int).Value = StoreProductServiceSn
        com.Parameters.Add("@SharedRate", SqlDbType.Decimal).Value = SharedRate
        com.Parameters.Add("@StartDate", SqlDbType.VarChar, 25).Value = StartDate
        com.Parameters.Add("@EndDate", SqlDbType.VarChar, 25).Value = EndDate
        com.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserStamp

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnRemindMsg As SqlClient.SqlParameter = com.Parameters.Add("@RemindMsg", SqlDbType.NVarChar, 100)
        ReturnRemindMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnRE
        Try
            conn.Open()
            com.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        'ReturnValue.ReturnRemindMsg = ReturnRemindMsg.Value

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 取得遊戲廠商資料(未續約)
    ''' </summary>
    ''' <param name="GameFactoryId"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_GameFactory_View_Outer(ByVal GameFactoryId As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "Select Distinct GameFactoryId,GameFactoryName,HwaId,Status From PointsBilling_DB.dbo.View_StorePoints_StoreProductServiceSharedRate_Outer where GameFactoryId=@GameFactoryId And Status=1"
        com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception

            Throw New Exception("查詢未續約遊戲廠商失敗|Get_GameFactory_View_Outer|" & ex.Message)

        End Try

        Return ds

    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 取得遊戲服務(未續約)
    ''' </summary>
    ''' <param name="GameFactoryId"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_GameService_View_Outer(ByVal GameFactoryId As String, ByVal hidGSerId As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "Select Distinct GameServiceId,GameName From PointsBilling_DB.dbo.View_StorePoints_StoreProductServiceSharedRate_Outer where GameServiceId in (" & hidGSerId & ") And GameFactoryId=@GameFactoryId And Status=1"
        com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception

            Throw New Exception("查詢未續約遊戲服務失敗|Get_GameService_View|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 取得金流服務(未續約)
    ''' </summary>
    ''' <param name="GameServiceId"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_FacService_View_Outer(ByVal GameServiceId As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "Select Distinct FactoryId,FactoryName From PointsBilling_DB.dbo.View_StorePoints_StoreProductServiceSharedRate_Outer where GameServiceId=@GameServiceId And Status=1"
        com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 20).Value = GameServiceId

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("取得未續約金流服務|Get_FacService_View_Outer|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 取得產品服務(未續約)
    ''' </summary>
    ''' <param name="GameFactoryId"></param>
    ''' <param name="GameServiceId"></param>
    ''' <param name="FactoryId"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_ProductService_View_Outer(ByVal GameFactoryId As String, ByVal GameServiceId As String, ByVal FactoryId As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandTimeout = 300
        com.CommandText = "Select Distinct Sn,StoreProductServiceSn,FactoryId,FactoryName,FactoryServiceName,GameProductName,SharedRate,GameFactoryId,GameServiceId,Status,StartDate,EndDate From PointsBilling_DB.dbo.View_StorePoints_StoreProductServiceSharedRate_Outer where GameFactoryId=@GameFactoryId And GameServiceId=@GameServiceId And FactoryId=@FactoryId And Status=1 order by FactoryServiceName,GameProductName,EndDate desc"
        com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId
        com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢未續約產品服務失敗|Get_ProductService_View|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function

    ''' <summary>
    ''' PayPal主動查詢交易狀態
    ''' 取得查詢時間範圍(小時)
    ''' </summary>
    ''' <param name="FactoryId">金流代號</param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function View_BK_StorePoints_FactoryForTimePeriod(ByVal FactoryId As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandTimeout = 300
        com.CommandText = "Select Sn,FactoryId,TimePeriod From PointsBilling_DB.dbo.View_BK_StorePoints_Factory where FactoryId=@FactoryId"
        com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢時間範圍失敗|View_BK_StorePoints_FactoryForTimePeriod|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function

    ''' <summary>
    ''' PayPal主動查詢交易狀態
    ''' 取得PayPal付費失敗的交易
    ''' </summary>
    ''' <param name="CheckDate">日期時間</param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function View_BK_StorePoints_Factory_Paypal(ByVal CheckDate As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandTimeout = 300
        com.CommandText = "Select TradeSeq,Price,FactoryTradeCode,CreateDate From PointsBilling_DB.dbo.View_BK_StorePoints_Factory_Paypal where CreateDate >= @CheckDate"
        com.Parameters.Add("@CheckDate", SqlDbType.VarChar, 20).Value = CheckDate

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢PayPal付費失敗的交易失敗|View_BK_StorePoints_Factory_Paypal|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function
    ''' <summary>
    ''' 查詢寄送郵件的Email
    ''' </summary>
    ''' <param name="GroupSn">群組代碼</param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Mycard_Backup_Remind_VB_Mail_View(ByVal GroupSn As Integer) As DataSet
        ConnStr(14)
        Dim command As New SqlClient.SqlCommand
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        command.CommandText = "Select * from MyCard_Backup.dbo.Mycard_Backup_Remind_VB_Mail_View  where Sn=@GroupSn order by UserMail"
        command.Parameters.Add("@GroupSn", SqlDbType.Int).Value = GroupSn
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            conn.Open()
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("撈取Mail時發生錯誤：" & ex.Message)
        Finally
            conn.Close()
        End Try
        Return DSssid
    End Function

    ''' <summary>
    ''' Billing後台-依付費方式批次修改產品服務
    ''' 付費廠商
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_View_BK_StorePoints_Factory() As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "Select FactoryId,FactoryName From PointsBilling_DB.dbo.View_BK_StorePoints_Factory"

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception

            Throw New Exception("查詢付費廠商失敗|Get_View_BK_StorePoints_Factory|" & ex.Message)

        End Try
        Return ds

    End Function

    ''' <summary>
    ''' Billing後台-依付費方式批次修改產品服務
    ''' 付費廠商品項名稱
    ''' </summary>
    ''' <param name="FactoryId">付費廠商</param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_View_BK_StorePoints_FactoryService(ByVal FactoryId As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "Select FactoryServiceId,FactoryServiceName From PointsBilling_DB.dbo.View_BK_StorePoints_FactoryService where FactoryId=@FactoryId"
        com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception

            Throw New Exception("查詢付費廠商品項名稱失敗|Get_View_BK_StorePoints_FactoryService|" & ex.Message)

        End Try
        Return ds

    End Function

    ''' <summary>
    ''' Billing後台-依付費方式批次修改產品服務
    ''' 查詢遊戲服務
    ''' </summary>
    ''' <param name="FactoryServiceId"></param>
    ''' <param name="FactoryServicePrice"></param>
    ''' <param name="Status"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_View_BK_StorePoints_StoreProductServiceForBK(ByVal FactoryServiceId As String, ByVal FactoryServicePrice As Integer, ByVal Status As Integer) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        Dim sqlstr As String = "Select Sn,GameName+'-'+GameProductName as GameName ,Status From PointsBilling_DB.dbo.View_BK_StorePoints_StoreProductServiceForBK where 1=1 "
        If FactoryServiceId <> "" Then
            sqlstr = sqlstr & " and FactoryServiceId=@FactoryServiceId "
            com.Parameters.Add("@FactoryServiceId", SqlDbType.VarChar, 30).Value = FactoryServiceId
        End If
        If FactoryServicePrice <> -1 Then
            sqlstr = sqlstr & " and FactoryServicePrice=@FactoryServicePrice "
            com.Parameters.Add("@FactoryServicePrice", SqlDbType.Int).Value = FactoryServicePrice
        End If
        If Status <> -1 Then
            sqlstr = sqlstr & " and Status=@Status "
            com.Parameters.Add("@Status", SqlDbType.Int).Value = Status
        End If

        com.CommandText = sqlstr

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception

            Throw New Exception("查詢遊戲服務失敗|Get_View_BK_StorePoints_FactoryService|" & ex.Message)

        End Try
        Return ds

    End Function

    ''' <summary>
    ''' Billing後台-依付費方式批次修改產品服務
    ''' 修改產品服務狀態
    ''' </summary>
    ''' <param name="Sn"></param>
    ''' <param name="Status">狀態</param>
    ''' <param name="UserStamp">修改人</param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_StoreProductService_UpdateStatus(ByVal Sn As Integer, ByVal Status As Integer, ByVal UserStamp As String) As ReturnResult
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_StoreProductService_UpdateStatus"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure

        com.Parameters.Add("@Sn", SqlDbType.Int).Value = Sn
        com.Parameters.Add("@Status", SqlDbType.Int).Value = Status
        com.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserStamp
        '參數-回傳訊息
        Dim ReturnMsg As New SqlParameter("@ReturnMsg", SqlDbType.VarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        com.Parameters.Add(ReturnMsg)
        '參數-回傳值
        Dim ReturnMsgNo As New SqlParameter("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        com.Parameters.Add(ReturnMsgNo)

        Dim ReturnValue As New ReturnResult
        Try
            conn.Open()
            com.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("修改失敗|SPS_StorePoints_StoreProductService_UpdateStatus|" & ex.Message)
        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台-CP廠商付費方式查詢
    ''' 查詢
    ''' </summary>
    ''' <param name="FactoryId">付費廠商代號</param>
    ''' <param name="ONLINEsaveTYPE">一般Billing、線上購卡儲值</param>
    ''' <param name="Status">上下架</param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetView_BK_StorePoints_StoreProductService_CPPayment(ByVal FactoryId As String, ByVal ONLINEsaveTYPE As String, ByVal Status As String) As DataSet
        ConnStr(14)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        Dim sqlstr As String = "Select distinct FactoryName,GameFactoryName,GameName,ONLINEsaveTYPE,Status,FactoryId From PointsBilling_DB.dbo.View_BK_StorePoints_StoreProductService_CPPayment where 1=1 "
        If FactoryId <> "" Then
            sqlstr = sqlstr & " and FactoryId=@FactoryId "
            com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        End If
        If ONLINEsaveTYPE <> "" Then
            sqlstr = sqlstr & " and ONLINEsaveTYPE=@ONLINEsaveTYPE "
            com.Parameters.Add("@ONLINEsaveTYPE", SqlDbType.VarChar, 14).Value = ONLINEsaveTYPE
        End If
        If Status <> "" Then
            sqlstr = sqlstr & " and Status=@Status "
            com.Parameters.Add("@Status", SqlDbType.VarChar, 4).Value = Status
        End If
        sqlstr = sqlstr & " Order By GameFactoryName,GameName,ONLINEsaveTYPE,Status desc"
        com.CommandText = sqlstr

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception

            Throw New Exception("查詢CP廠商付費方式失敗|GetView_BK_StorePoints_StoreProductService_CPPayment|" & ex.Message)

        End Try
        Return ds

    End Function

    ''' <summary>
    ''' Billing後台-CP廠商群組上架查詢
    ''' 依金額查詢卡群組
    ''' </summary>
    ''' <param name="CARD_PRICE">金額</param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetOGB_TB_CARD_GROUPByPrice(ByVal CARD_PRICE As Decimal) As DataSet
        ConnStr(14)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        Dim sqlstr As String = "Select CARD_GROUP,CARD_PRICE,CARD_GROUP_DESC From OGB_DB_CARD.dbo.View_OGB_TB_CardGroup where 1=1 "
        If CARD_PRICE <> -1 Then
            sqlstr = sqlstr & " and CARD_PRICE=@CARD_PRICE "
            com.Parameters.Add("@CARD_PRICE", SqlDbType.Money).Value = CARD_PRICE
        End If

        com.CommandText = sqlstr

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception

            Throw New Exception("查詢卡群組失敗|GetOGB_TB_CARD_GROUPByPrice|" & ex.Message)

        End Try
        Return ds

    End Function

    ''' <summary>
    ''' Billing後台-CP廠商群組上架查詢
    ''' 查詢
    ''' </summary>
    ''' <param name="CARD_GROUP">卡群組</param>
    ''' <param name="Status">web儲值上下架</param>
    ''' <param name="ImGameStatus">ingame上下架</param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetView_OGB_TB_CardGroup_WebIngameStatus(ByVal CARD_GROUP As String, ByVal Status As String, ByVal ImGameStatus As String, ByVal WebInGameStatus As String) As DataSet
        ConnStr(14)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        Dim sqlstr As String = "Select CARD_GROUP_DESC,GameFacName,DSCR,Status,ImGameStatus,WebInGameStatus,CARD_GROUP From OGB_DB_CARD.dbo.View_OGB_TB_CardGroup_WebIngameStatus where 1=1 "
        If CARD_GROUP <> "" Then
            If IsNumeric(CARD_GROUP) Then
                CARD_GROUP = CARD_GROUP.PadLeft(4, "0")
                sqlstr = sqlstr & " and CARD_GROUP Like @CARD_GROUP "
                com.Parameters.AddWithValue("@CARD_GROUP", "%" & CARD_GROUP)
            Else
                sqlstr = sqlstr & " and CARD_GROUP=@CARD_GROUP "
                com.Parameters.Add("@CARD_GROUP", SqlDbType.VarChar, 10).Value = CARD_GROUP
            End If
        End If
        If Status <> "" Then
            sqlstr = sqlstr & " and Status=@Status "
            com.Parameters.Add("@Status", SqlDbType.VarChar, 4).Value = Status
        End If
        If ImGameStatus <> "" Then
            sqlstr = sqlstr & " and ImGameStatus=@ImGameStatus "
            com.Parameters.Add("@ImGameStatus", SqlDbType.VarChar, 4).Value = ImGameStatus
        End If
        If WebInGameStatus <> "" Then
            sqlstr = sqlstr & " and WebInGameStatus=@WebInGameStatus "
            com.Parameters.Add("@WebInGameStatus", SqlDbType.VarChar, 4).Value = WebInGameStatus
        End If

        sqlstr = sqlstr & " Order By CARD_GROUP_DESC,GameFacName,DSCR"
        com.CommandText = sqlstr

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception

            Throw New Exception("查詢CP廠商群組上架失敗|GetView_OGB_TB_CardGroup_WebIngameStatus|" & ex.Message)

        End Try
        Return ds

    End Function

    ''' <summary>
    ''' WinForm-Mail每月發出CP廠商配合實體點數一覽表
    ''' 查詢
    ''' </summary>
    ''' <param name="CARD_GROUP_Type">卡群組: M點，N超商，O國外，P大馬，C點交易</param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetView_OGB_TB_CardGroup_WebIngameStatusForWinForm(ByVal CARD_GROUP_Type As String) As DataSet
        ConnStr(14)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        Dim sqlstr As String = "Select GameFacName,DSCR,Status,ImGameStatus,CARD_GROUP,WebInGameStatus From OGB_DB_CARD.dbo.View_OGB_TB_CardGroup_WebIngameStatus where 1=1 "
        If CARD_GROUP_Type <> "" Then
            sqlstr = sqlstr & " and Left(CARD_GROUP,1) = @CARD_GROUP AND Left(CARD_GROUP,2) != 'MA'"
            com.Parameters.AddWithValue("@CARD_GROUP", CARD_GROUP_Type.Trim)
        End If

        sqlstr = sqlstr & " Order By GameFacName,DSCR,CARD_GROUP"
        com.CommandText = sqlstr

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception

            Throw New Exception("查詢CP廠商群組失敗|GetView_OGB_TB_CardGroup_WebIngameStatusForWinForm|" & ex.Message)

        End Try
        Return ds

    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 查詢CP所屬的付費廠商
    ''' </summary>
    ''' <param name="GameFactoryId"></param>
    ''' <param name="GameServiceId"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetView_BK_StorePoints_StoreProductServiceForRateForCP(ByVal GameFactoryId As String, ByVal GameServiceId As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        Dim sqlstr As String = "Select distinct FactoryId,FactoryName From PointsBilling_DB.dbo.View_BK_StorePoints_StoreProductServiceForRate where 1=1 "
        If GameFactoryId <> "" Then
            sqlstr = sqlstr & " and GameFactoryId=@GameFactoryId"
            com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId
        End If
        If GameServiceId <> "" Then
            sqlstr = sqlstr & " and GameServiceId=@GameServiceId"
            com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        End If

        com.CommandText = sqlstr
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢CP所屬的付費廠商失敗|GetView_BK_StorePoints_StoreProductServiceForRateForCP|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds
    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 查詢所有付費廠商對應遊戲
    ''' </summary>
    ''' <param name="GameFactoryId"></param>
    ''' <param name="GameServiceId"></param>
    ''' <param name="FactoryId"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetView_BK_StorePoints_StoreProductServiceForRate(ByVal GameFactoryId As String, ByVal GameServiceId As String, ByVal FactoryId As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        Dim sqlstr As String = "Select  * From PointsBilling_DB.dbo.View_BK_StorePoints_StoreProductServiceForRate where 1=1 "
        If GameFactoryId <> "" Then
            sqlstr = sqlstr & " and GameFactoryId=@GameFactoryId"
            com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId
        End If
        If GameServiceId <> "" Then
            sqlstr = sqlstr & " and GameServiceId=@GameServiceId"
            com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        End If
        If GameFactoryId <> "" Then
            sqlstr = sqlstr & " and FactoryId=@FactoryId"
            com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        End If

        com.CommandText = sqlstr
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢所有付費廠商對應遊戲失敗|GetView_BK_StorePoints_StoreProductServiceForRate|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds
    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 查詢合約日期
    ''' </summary>
    ''' <param name="GameFactoryId"></param>
    ''' <param name="GameServiceId"></param>
    ''' <param name="FactoryId"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetView_BK_StorePoints_GameServiceContractDate(ByVal GameFactoryId As String, ByVal GameServiceId As String, ByVal FactoryId As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        Dim sqlstr As String = "Select * From PointsBilling_DB.dbo.View_BK_StorePoints_GameServiceContractDate where 1=1 "
        If GameFactoryId <> "" Then
            sqlstr = sqlstr & " and GameFactoryId=@GameFactoryId"
            com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId
        End If
        If GameServiceId <> "" Then
            sqlstr = sqlstr & " and GameServiceId=@GameServiceId"
            com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        End If
        If GameFactoryId <> "" Then
            sqlstr = sqlstr & " and FactoryId=@FactoryId"
            com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        End If
        sqlstr = sqlstr & " order by StartDate desc"
        com.CommandText = sqlstr
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢合約日期|GetView_BK_StorePoints_GameServiceContractDate|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds
    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 查詢折數
    ''' </summary>
    ''' <param name="GameFactoryId"></param>
    ''' <param name="GameServiceId"></param>
    ''' <param name="FactoryId"></param>
    ''' <param name="GSContractSn">合約日期Sn</param>
    ''' <returns>DataSet,ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳DataSet,ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_GameServiceContractRate_Query(ByVal GameFactoryId As String, ByVal GameServiceId As String, ByVal FactoryId As String, ByVal GSContractSn As Integer) As ReportRE
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_GameServiceContractRate_Query"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure

        com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId
        com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 20).Value = GameServiceId
        com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        com.Parameters.Add("@GSContractSn", SqlDbType.Int).Value = GSContractSn

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReportRE
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢折數失敗|SPS_StorePoints_GameServiceContractRate_Query|" & ex.Message)
        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
        ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
        ReturnValue.ReturnDS = ds

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 新增合約日期
    ''' </summary>
    ''' <param name="GameFactoryId"></param>
    ''' <param name="GameServiceId"></param>
    ''' <param name="FactoryId"></param>
    ''' <param name="StartDate"></param>
    ''' <param name="EndDate"></param>
    ''' <param name="CreateUser"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_GameServiceContractDate_Insert(ByVal GameFactoryId As String, ByVal GameServiceId As String, ByVal FactoryId As String, ByVal StartDate As Date, ByVal EndDate As Date, ByVal CreateUser As String) As ReturnRE
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_GameServiceContractDate_Insert"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure

        com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 30).Value = FactoryId
        com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 30).Value = GameFactoryId
        com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        com.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
        com.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
        com.Parameters.Add("@CreateUser", SqlDbType.VarChar, 50).Value = CreateUser

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 50)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnRE
        Try
            conn.Open()
            com.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("新增合約日期|" & ex.Message)
        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 修改合約日期
    ''' </summary>
    ''' <param name="GameFactoryId"></param>
    ''' <param name="GameServiceId"></param>
    ''' <param name="FactoryId"></param>
    ''' <param name="Sn"></param>
    ''' <param name="StartDate"></param>
    ''' <param name="EndDate"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_GameServiceContractDate_Update(ByVal GameFactoryId As String, ByVal GameServiceId As String, ByVal FactoryId As String, ByVal Sn As Integer, ByVal StartDate As Date, ByVal EndDate As Date, ByVal UserStamp As String) As ReturnRE
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_GameServiceContractDate_Update"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure

        com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 30).Value = FactoryId
        com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 30).Value = GameFactoryId
        com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        com.Parameters.Add("@Sn", SqlDbType.Int).Value = Sn
        com.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
        com.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
        com.Parameters.Add("@UserStamp", SqlDbType.VarChar, 50).Value = UserStamp

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 50)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnRE
        Try
            conn.Open()
            com.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("修改合約日期|" & ex.Message)
        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 修改合約折數
    ''' </summary>
    ''' <param name="GSContractDateSn"></param>
    ''' <param name="StoreProductServiceId"></param>
    ''' <param name="ShareRate"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_GameServiceContractRate_Update(ByVal GSContractDateSn As Integer, ByVal StoreProductServiceId As String, ByVal ShareRate As Double, ByVal UserStamp As String) As ReturnRE
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_GameServiceContractRate_Update"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure

        com.Parameters.Add("@GSContractDateSn", SqlDbType.Int).Value = GSContractDateSn
        com.Parameters.Add("@StoreProductServiceId", SqlDbType.VarChar, 10).Value = StoreProductServiceId
        com.Parameters.Add("@ShareRate", SqlDbType.Decimal).Value = ShareRate
        com.Parameters.Add("@UserStamp", SqlDbType.VarChar, 50).Value = UserStamp

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 50)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnRE
        Try
            conn.Open()
            com.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("修改合約折數|" & ex.Message)
        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 新增合約折數
    ''' </summary>
    ''' <param name="GSContractDateSn"></param>
    ''' <param name="StoreProductServiceId"></param>
    ''' <param name="GameServiceId"></param>
    ''' <param name="GameProductId"></param>
    ''' <param name="FactoryServiceId"></param>
    ''' <param name="ShareRate"></param>
    ''' <param name="CreateUser"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_GameServiceContractRate_Insert(ByVal GSContractDateSn As Integer, ByVal StoreProductServiceId As String, ByVal GameServiceId As String, ByVal GameProductId As String, ByVal FactoryServiceId As String, ByVal ShareRate As Double, ByVal CreateUser As String) As ReturnRE
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_GameServiceContractRate_Insert"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure

        com.Parameters.Add("@GSContractDateSn", SqlDbType.Int).Value = GSContractDateSn
        com.Parameters.Add("@StoreProductServiceId", SqlDbType.VarChar, 10).Value = StoreProductServiceId
        com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        com.Parameters.Add("@GameProductId", SqlDbType.VarChar, 30).Value = GameProductId
        com.Parameters.Add("@FactoryServiceId", SqlDbType.VarChar, 30).Value = FactoryServiceId
        com.Parameters.Add("@ShareRate", SqlDbType.Decimal).Value = ShareRate
        com.Parameters.Add("@CreateUser", SqlDbType.VarChar, 50).Value = CreateUser

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 50)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnRE
        Try
            conn.Open()
            com.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("新增合約折數|" & ex.Message)
        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台-廠商結帳配合條件設定
    ''' 新增日期及合約折數
    ''' </summary>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_GameServiceContractDate_Insert_BeginTrans(ByVal GameFactoryId As String, ByVal GameServiceId As String, ByVal FactoryIdString As String, ByVal StartDate As Date, ByVal EndDate As Date, ByVal ShareRate As Double, ByVal CreateUser As String) As ReturnRE
        Dim Msg As String = ""
        Dim ReturnMsgForLoop As String = ""
        Dim ReturnValue As New ReturnRE

        Dim FactoryString() As String = Split(FactoryIdString, "|")

        For i As Integer = 0 To FactoryString.Length - 1
            Dim FactoryId() As String = Split(FactoryString(i), "^")

            ConnStr(12)
            BeginTrans(Msg)
            '新增合約日期
            Dim sqlcomm As String
            sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_GameServiceContractDate_Insert"
            Dim com As New SqlCommand(sqlcomm, conn)
            com.CommandTimeout = 100
            com.CommandType = CommandType.StoredProcedure

            com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 30).Value = FactoryId(0)
            com.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 30).Value = GameFactoryId
            com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
            com.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
            com.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
            com.Parameters.Add("@CreateUser", SqlDbType.VarChar, 50).Value = CreateUser

            Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
            ReturnMsgNo.Direction = ParameterDirection.Output
            Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 50)
            ReturnMsg.Direction = ParameterDirection.Output
            Dim ReturnGSContractDateSn As SqlClient.SqlParameter = com.Parameters.Add("@GSContractDateSn", SqlDbType.Int)
            ReturnGSContractDateSn.Direction = ParameterDirection.Output

            Try
                ConnOpen(com)
                com.ExecuteNonQuery()
            Catch ex As Exception
                RollbackTrans(Msg)
                ReturnMsgNo.Value = -99
                ReturnMsg.Value = "新增合約日期|" & ex.Message
            Finally
                ConnClose()
            End Try

            If ReturnMsgNo.Value = 1 Then

                '查詢付費廠商對應品項
                Dim com1 As New SqlCommand()
                com1.Connection = conn
                Dim sqlstr As String = "Select  * From PointsBilling_DB.dbo.View_BK_StorePoints_StoreProductServiceForRate where 1=1 "
                If GameFactoryId <> "" Then
                    sqlstr = sqlstr & " and GameFactoryId=@GameFactoryId"
                    com1.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId
                End If
                If GameServiceId <> "" Then
                    sqlstr = sqlstr & " and GameServiceId=@GameServiceId"
                    com1.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
                End If
                If GameFactoryId <> "" Then
                    sqlstr = sqlstr & " and FactoryId=@FactoryId"
                    com1.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId(0)
                End If

                com1.CommandText = sqlstr
                Dim da As New SqlDataAdapter(com1)
                Dim ds As New DataSet

                Try
                    ConnOpen(com1)
                    da.Fill(ds)
                Catch ex As Exception
                    RollbackTrans(Msg)
                    ReturnMsgNo.Value = -99
                    ReturnMsg.Value = "查詢付費廠商對應品項|" & ex.Message
                Finally
                    ConnClose()
                End Try

                If ReturnMsgNo.Value = 1 Then

                    '新增品項折數
                    Dim CheckStatus As Boolean = True
                    For j As Integer = 0 To ds.Tables(0).Rows.Count - 1
                        Dim sqlcomm2 As String
                        sqlcomm2 = "PointsBilling_DB.dbo.SPS_StorePoints_GameServiceContractRate_Insert"
                        Dim com2 As New SqlCommand(sqlcomm2, conn)
                        com2.CommandTimeout = 100
                        com2.CommandType = CommandType.StoredProcedure

                        com2.Parameters.Add("@GSContractDateSn", SqlDbType.Int).Value = ReturnGSContractDateSn.Value
                        com2.Parameters.Add("@StoreProductServiceId", SqlDbType.VarChar, 10).Value = ds.Tables(0).Rows(j).Item("StoreProductServiceId")
                        com2.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = ds.Tables(0).Rows(j).Item("GameServiceId")
                        com2.Parameters.Add("@GameProductId", SqlDbType.VarChar, 30).Value = ds.Tables(0).Rows(j).Item("GameProductId")
                        com2.Parameters.Add("@FactoryServiceId", SqlDbType.VarChar, 30).Value = ds.Tables(0).Rows(j).Item("FactoryServiceId")
                        com2.Parameters.Add("@ShareRate", SqlDbType.Decimal).Value = ShareRate
                        com2.Parameters.Add("@CreateUser", SqlDbType.VarChar, 50).Value = CreateUser

                        Dim ReturnMsgNo2 As SqlClient.SqlParameter = com2.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
                        ReturnMsgNo2.Direction = ParameterDirection.Output
                        Dim ReturnMsg2 As SqlClient.SqlParameter = com2.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 50)
                        ReturnMsg2.Direction = ParameterDirection.Output

                        Try
                            ConnOpen(com2)
                            com2.ExecuteNonQuery()
                        Catch ex As Exception
                            RollbackTrans(Msg)
                            ReturnMsgNo2.Value = -99
                            ReturnMsg2.Value = "新增合約折數|" & ex.Message
                        Finally
                            ConnClose()
                        End Try
                        If ReturnMsgNo2.Value <> 1 Then
                            RollbackTrans(Msg)
                            CheckStatus = False
                            ReturnMsgForLoop = ReturnMsgForLoop & ds.Tables(0).Rows(0).Item("FactoryName") & " " & ReturnMsg2.Value & "(" & ReturnMsgNo2.Value & ")<br>"
                            Exit For
                        End If
                        System.IO.File.AppendAllText(Server.MapPath("WorkLog.txt"), Now.ToString("yyyy/MM/dd HH:mm:ss") & "  " & ReturnGSContractDateSn.Value & "," & ds.Tables(0).Rows(j).Item("StoreProductServiceId") & "," & ds.Tables(0).Rows(j).Item("GameServiceId") & "," & ds.Tables(0).Rows(j).Item("GameProductId") & "," & ds.Tables(0).Rows(j).Item("FactoryServiceId") & "," & ShareRate & vbCrLf)
                    Next
                    If CheckStatus Then
                        CommitTrans(Msg)
                        ReturnMsgForLoop = ReturnMsgForLoop & ds.Tables(0).Rows(0).Item("FactoryName") & " 新增成功<br>"
                    End If
                Else
                    RollbackTrans(Msg)
                    ReturnMsgForLoop = ReturnMsgForLoop & FactoryId(1) & " " & ReturnMsg.Value & "(" & ReturnMsgNo.Value & ")<br>"
                End If

            Else
                RollbackTrans(Msg)
                ReturnMsgForLoop = ReturnMsgForLoop & FactoryId(1) & " " & ReturnMsg.Value & "(" & ReturnMsgNo.Value & ")<br>"
            End If
        Next

        ReturnValue.ReturnMsgNo = 1
        ReturnValue.ReturnMsg = ReturnMsgForLoop
        Return ReturnValue
    End Function

    ''' <summary>
    ''' CP廠商遊戲服務管理
    ''' </summary>
    ''' <param name="GameServiceID"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function View_BK_StorePoints_GameServiceLoginData(ByVal GameServiceID As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim sqlcomm As String
        sqlcomm = "select GameServiceId,GameName,Sn,IPTypeDesc,StoreIP from PointsBilling_DB.dbo.View_BK_StorePoints_GameServiceLoginData Where GameServiceId=@GameServiceId"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 300
        com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceID
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet()
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("View_BK_StorePoints_GameServiceLoginData|" & ex.Message)
        Finally
            conn.Close()
        End Try
        Return ds
    End Function

    ''' <summary>
    ''' Billing後台-CP廠商遊戲服務管理
    ''' 新增IP
    ''' </summary>
    ''' <param name="GameServiceId"></param>
    ''' <param name="IPTypeDesc"></param>
    ''' <param name="StoreIP"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_GameServiceLoginData_Insert(ByVal GameServiceId As String, ByVal IPTypeDesc As String, ByVal StoreIP As String, ByVal UserStamp As String) As ReturnRE
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_GameServiceLoginData_Insert"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure

        com.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        com.Parameters.Add("@IPTypeDesc", SqlDbType.NVarChar, 10).Value = IPTypeDesc
        com.Parameters.Add("@StoreIP", SqlDbType.VarChar, 20).Value = StoreIP
        com.Parameters.Add("@CreateUser", SqlDbType.VarChar, 30).Value = UserStamp
        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnRE
        Try
            conn.Open()
            com.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("新增IP|" & ex.Message)
        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台-CP廠商遊戲服務管理
    ''' 修改IP
    ''' </summary>
    ''' <param name="Sn"></param>
    ''' <param name="IPTypeDesc"></param>
    ''' <param name="StoreIP"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_GameServiceLoginData_Update(ByVal Sn As String, ByVal IPTypeDesc As String, ByVal StoreIP As String, ByVal UserStamp As String) As ReturnRE
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_GameServiceLoginData_Update"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure

        com.Parameters.Add("@SN", SqlDbType.Int, 4).Value = Sn
        com.Parameters.Add("@IPTypeDesc", SqlDbType.NVarChar, 10).Value = IPTypeDesc
        com.Parameters.Add("@StoreIP", SqlDbType.VarChar, 20).Value = StoreIP
        com.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserStamp
        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnRE
        Try
            conn.Open()
            com.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("修改IP|" & ex.Message)
        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台/Billing後台維護作業/付費廠商對應單位別功能
    ''' 查詢已設定付費廠商
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_View_BK_StorePoints_FactoryInDept_Inner(ByVal DEPTSN As String) As DataSet
        ConnStr(12)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandText = "Select FactoryId,FactoryName,DEPTSN From PointsBilling_DB.dbo.View_BK_StorePoints_FactoryInDept_Inner Where DEPTSN=@DeptSN"
        com.Parameters.Add("@DeptSN", SqlDbType.Int).Value = DEPTSN

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢已設定付費廠商|Get_View_BK_StorePoints_FactoryInDept_Inner|" & ex.Message)
        End Try
        Return ds
    End Function

    ''' <summary>
    ''' Billing後台/Billing後台維護作業/付費廠商對應單位別功能
    ''' 查詢未設定部門的金流
    ''' </summary>
    ''' <param name="DeptSN"></param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_FactoryInDept_Outter(ByVal DeptSN As Integer) As FactoryInDeptRE
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_FactoryInDept_Outter"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure

        com.Parameters.Add("@DeptSn", SqlDbType.Int).Value = DeptSN

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New FactoryInDeptRE

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet()
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("SPS_StorePoints_FactoryInDept_Outter|" & ex.Message)
        Finally
            conn.Close()
        End Try

        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        ReturnValue.ReturnDS = ds

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台/Billing後台維護作業/付費廠商對應單位別功能
    ''' 新增金流廠商對應部門
    ''' </summary>
    ''' <param name="DeptSN"></param>
    ''' <param name="FactoryID"></param>
    ''' <param name="CreateUser"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_FactoryInDept_Insert(ByVal DeptSN As Integer, ByVal FactoryId As String, ByVal CreateUser As String) As FactoryInDeptRE
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_FactoryInDept_Insert"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure

        com.Parameters.Add("@DeptSn", SqlDbType.Int).Value = DeptSN
        com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        com.Parameters.Add("@CreateUser", SqlDbType.VarChar, 50).Value = CreateUser
        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New FactoryInDeptRE
        Try
            conn.Open()
            com.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("SPS_StorePoints_FactoryInDept_Insert|" & ex.Message)
        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台/Billing後台維護作業/付費廠商對應單位別功能
    ''' 刪除金流廠商對應部門
    ''' </summary>
    ''' <param name="DeptSN"></param>
    ''' <param name="FactoryID"></param>
    ''' <param name="CreateUser"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_FactoryInDept_Delete(ByVal DeptSN As Integer, ByVal FactoryId As String, ByVal CreateUser As String) As FactoryInDeptRE
        ConnStr(12)
        Dim sqlcomm As String
        sqlcomm = "PointsBilling_DB.dbo.SPS_StorePoints_FactoryInDept_Delete"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        com.CommandType = CommandType.StoredProcedure

        com.Parameters.Add("@DeptSn", SqlDbType.Int).Value = DeptSN
        com.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        com.Parameters.Add("@CreateUser", SqlDbType.VarChar, 50).Value = CreateUser
        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New FactoryInDeptRE
        Try
            conn.Open()
            com.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("SPS_StorePoints_FactoryInDept_Delete|" & ex.Message)
        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value

        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台/Billing後台維護作業/付費廠商對應單位別功能
    ''' 取得後台單位別
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_BKPortalDept() As DataSet
        ConnStr(14)
        Dim sqlcomm As String
        sqlcomm = "select distinct Sn,DeptName from Mycard_backup.dbo.VIEW_BKPortal_Dept"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 300

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet()
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("Get_BKPortalDept|" & ex.Message)
        Finally
            conn.Close()
        End Try
        Return ds
    End Function

    ''' <summary>
    ''' 付費廠商對應單位別功能需求
    ''' 取得帳號所屬單位別
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function VIEW_BKPortal_SWUsers(ByVal User_UAccount As String) As DataSet
        ConnStr(14)
        Dim cmd As New SqlClient.SqlCommand()
        cmd.CommandType = CommandType.Text
        cmd.Connection = conn
        cmd.CommandTimeout = 300
        cmd.CommandText = "SELECT * FROM MyCard_Backup.dbo.VIEW_BKPortal_SWUsers WHERE User_UAccount =@User_UAccount and User_Status<>-1"
        cmd.Parameters.Add("@User_UAccount", SqlDbType.VarChar, 32).Value = User_UAccount
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("VIEW_BKPortal_SWUsers|" & ex.Message)
        Finally
            conn.Close()
        End Try
        Return DSssid
    End Function
    ''' <summary>
    ''' 差異報表處理功能
    ''' 差異報表補跑與取消功能
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_EC_GAME_TypeForCompany() As DataSet
        ConnStr(14)
        Dim myconn As New MyCardConn.Dbfuction
        'Dim Conn As New SqlConnection
        'Conn = lai.open14Db()
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        command.CommandText = "Select GameFacName  from MyCard_Backup.dbo.Admin_MyCardReport_FactoryGameType_View Group by GameFacName  order by GameFacName asc"
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("-1查詢Get_EC_GAME_TypeForCompany資料失敗：" & ex.Message)
        End Try
        Return DSssid
    End Function
    ''' <summary>
    ''' 差異報表處理功能
    ''' 
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_EC_GAME_TypeForGame(ByVal GameFacName As String) As DataSet
        ConnStr(14)
        Dim myconn As New MyCardConn.Dbfuction
        'Dim Conn As New SqlConnection
        'Conn = lai.open14Db()
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        If GameFacName = "" Then
            command.CommandText = "SELECT * FROM  MyCard_Backup.dbo.Admin_MyCardReport_FactoryGameType_View_A02  order by DSCR"
        Else
            command.CommandText = "Select * from MyCard_Backup.dbo.Admin_MyCardReport_FactoryGameType_View_A02  where GameFacName=@GameFacName order by DSCR"
            command.Parameters.Add("@GameFacName", SqlDbType.VarChar, 50).Value = GameFacName
        End If
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("-1_1查詢Get_EC_GAME_TypeForGame資料失敗：" & ex.Message)
        End Try
        Return DSssid
    End Function
    ''' <summary>
    ''' 差異報表處理功能
    ''' 
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetGameFacIdAndGAME_ID(ByVal DSCR As String) As DataSet
        ConnStr(14)
        Dim myconn As New MyCardConn.Dbfuction
        'Dim Conn As New SqlConnection
        'Conn = lai.open14Db()
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        command.CommandText = "Select * from MyCard_Backup.dbo.Admin_MyCardReport_FactoryGameType_View where DSCR=@DSCR order by GameFacName"
        command.Parameters.Add("@DSCR", SqlDbType.VarChar, 100).Value = DSCR
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("-1_2查詢GetGameFacIdAndGAME_ID資料失敗：" & ex.Message)
        End Try
        Return DSssid
    End Function
    ''' <summary>
    ''' 差異報表處理功能
    ''' 
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetVIEW_MyCP_ReasonType() As DataSet
        ConnStr(14)
        Dim myconn As New MyCardConn.Dbfuction
        'Dim Conn As New SqlConnection
        'Conn = lai.open14Db()
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        command.CommandText = "Select * from MyCard_Backup.dbo.VIEW_MyCP_ReasonType order by Priority,Name asc"
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("-2查詢VIEW_MyCP_ReasonType資料失敗：" & ex.Message)
        End Try
        Return DSssid
    End Function
    ''' <summary>
    ''' 差異報表處理功能
    ''' 
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function DiffReportProcessGetData(ByVal DateStart As String, ByVal DateEnd As String, ByVal GameFacName As String, ByVal GameServiceId As String, ByVal ProceStatus As String) As DataSet
        ConnStr(14)
        Dim ParaStr As String = ""
        'Dim Conn As New SqlConnection
        'Conn = lai.open14Db()
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = 300
        command.CommandText = "MyCard_Backup.dbo.MyCP_SaveDifferenceData_Query"
        command.Parameters.Add("@DateStart", SqlDbType.VarChar, 20).Value = DateStart
        command.Parameters.Add("@DateEnd", SqlDbType.VarChar, 20).Value = DateEnd
        command.Parameters.Add("@GameFacName", SqlDbType.VarChar, 50).Value = GameFacName
        command.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        command.Parameters.Add("@ProceStatus", SqlDbType.Int).Value = ProceStatus
        Dim ReturnNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnNo", SqlDbType.Int)
        ReturnNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 20)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("-3查詢差異報表處理資料失敗：" & ex.Message)
        End Try
        If ReturnNo.Value <> 1 Then
            Throw New Exception("-4查詢差異報表處理資料失敗：" & ReturnMsg.Value)
        End If
        Return DSssid
    End Function


    ''' <summary>
    ''' 差異報表補跑與取消功能
    ''' 
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function DiffReportSetGet_EC_GAME_TypeForGame(ByVal GameFacName As String) As DataSet
        ConnStr(14)
        Dim myconn As New MyCardConn.Dbfuction
        'Dim Conn As New SqlConnection
        'Conn = lai.open14Db()
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        'If GameFacId = "" Then
        '    command.CommandText = "Select GameFacName,GameFacId, GAME_ID, DSCR, EngDscr from MyCard_Backup.dbo.Admin_MyCardReport_FactoryGameType_View order by DSCR"
        'Else
        '    command.CommandText = "Select GameFacName,GameFacId, GAME_ID, DSCR, EngDscr from MyCard_Backup.dbo.Admin_MyCardReport_FactoryGameType_View where GameFacId=@GameFacId order by DSCR"
        '    command.Parameters.Add("@GameFacName", SqlDbType.VarChar, 50).Value = GameFacId
        'End If
        If GameFacName = "" Then
            command.CommandText = "SELECT GameFacName, GameFacId, GAME_ID, DSCR FROM Mycard_backup.dbo.Admin_MyCardReport_FactoryGameType_filtration_View Where GameFacName IS NOT NULL and DiffReportStatus ='1' order by DSCR"
        Else
            command.CommandText = "SELECT GameFacName, GameFacId, GAME_ID, DSCR FROM Mycard_backup.dbo.Admin_MyCardReport_FactoryGameType_filtration_View Where GameFacName IS NOT NULL and DiffReportStatus ='1' and GameFacName=@GameFacName"
            command.Parameters.Add("@GameFacName", SqlDbType.VarChar, 50).Value = GameFacName
        End If
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("-1_1查詢Get_EC_GAME_TypeForGame資料失敗：" & ex.Message)
        End Try
        Return DSssid
    End Function
    '''' <summary>
    '''' DiffReportSet
    '''' 差異報表補跑與取消功能
    '''' </summary>
    '''' <returns>DataSet</returns>
    '''' <remarks>
    '''' 回傳DataSet 
    '''' </remarks>
    '<WebMethod()> _
    'Public Function GetGameFacIdAndGAME_ID(ByVal DSCR As String) As DataSet
    '    ConnStr(14)
    '    Dim myconn As New MyCardConn.Dbfuction
    '    'Dim Conn As New SqlConnection
    '    'Conn = lai.open14Db()
    '    Dim command As New SqlClient.SqlCommand
    '    command.Parameters.Clear()
    '    command.Connection = conn
    '    command.CommandType = CommandType.Text
    '    command.CommandTimeout = 300
    '    command.CommandText = "Select GameFacName,GameFacId, GAME_ID, DSCR, EngDscr from MyCard_Backup.dbo.Admin_MyCardReport_FactoryGameType_View where DSCR=@DSCR order by GameFacName"
    '    command.Parameters.Add("@DSCR", SqlDbType.VarChar, 100).Value = DSCR
    '    Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
    '    Dim DSssid As New DataSet
    '    Try
    '        SDAssid.Fill(DSssid)
    '    Catch ex As Exception
    '        Throw New Exception("-2查詢GetGameFacIdAndGAME_ID資料失敗：" & ex.Message)
    '    End Try
    '    Return DSssid
    'End Function

    ''' <summary>
    ''' 差異報表補跑與取消功能
    ''' 
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function DiffReportSetGetData(ByVal Choice As Integer, ByVal DateStart As String, ByVal DateEnd As String, _
                            ByVal GameFacName As String, ByVal GameServiceId As String, ByVal UserId As String) As DataSet
        ConnStr(14)
        Dim ParaStr As String = ""
        'Dim Conn As New SqlConnection
        'Conn = lai.open14Db()
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = 300
        command.CommandText = "MyCard_Backup.dbo.MyCP_QueueProcess_ProcessData"
        command.Parameters.Add("@Choice", SqlDbType.Int).Value = Choice
        command.Parameters.Add("@DateStart", SqlDbType.VarChar, 20).Value = DateStart
        command.Parameters.Add("@DateEnd", SqlDbType.VarChar, 20).Value = DateEnd
        command.Parameters.Add("@GameFacName", SqlDbType.VarChar, 50).Value = GameFacName
        command.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        command.Parameters.Add("@UserId", SqlDbType.VarChar, 30).Value = UserId
        Dim ReturnNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnNo", SqlDbType.Int)
        ReturnNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 20)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("-3查詢差異報表補跑與取消功能資料失敗：" & ex.Message)
        End Try
        If ReturnNo.Value <> 1 Then
            Throw New Exception("-4查詢差異報表補跑與取消功能資料失敗：" & ReturnMsg.Value)
        End If
        Return DSssid
    End Function

    ''' <summary>
    ''' 差異報表補跑與取消功能
    ''' 
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function MyCP_QueueProcess_ProcessData(ByVal Choice As Integer, ByVal SData As String, ByVal EData As String, ByVal GameFacName As String, ByVal Game_Id As String, ByVal UserID As String) As ReturnRE
        ConnStr(14)
        'Dim Conn As New SqlConnection
        'Conn = lai.open14Db()
        ' 建立命令(預存程序名,使用連線) 
        Dim Command As New SqlClient.SqlCommand("MyCard_Backup.dbo.MyCP_QueueProcess_ProcessData", conn)
        ' 設置命令型態為預存程序 
        Command.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位 
        Command.Parameters.Add("@Choice", SqlDbType.Int).Value = Choice
        Command.Parameters.Add("@DateStart", SqlDbType.VarChar, 20).Value = SData
        Command.Parameters.Add("@DateEnd", SqlDbType.VarChar, 20).Value = EData
        Command.Parameters.Add("@GameFacName", SqlDbType.VarChar, 50).Value = GameFacName
        Command.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = Game_Id
        Command.Parameters.Add("@UserId", SqlDbType.VarChar, 30).Value = UserID
        Dim ReturnNo As SqlClient.SqlParameter = Command.Parameters.Add("@ReturnNo", SqlDbType.Int)
        ReturnNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 20)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnRE As New ReturnRE
        Try
            conn.Open()
            Command.ExecuteNonQuery()
            ReturnRE.ReturnMsgNo = ReturnNo.Value
            ReturnRE.ReturnMsg = ReturnMsg.Value
        Catch ex As Exception
            ReturnRE.ReturnMsgNo = -99
            ReturnRE.ReturnMsg = "-7Insert差異報表補跑與取消功能功能時：" & ex.Message

            Return ReturnRE
        Finally
            conn.Close()
        End Try
        Return ReturnRE
    End Function
    ''' <summary>
    ''' 差異報表補跑與取消功能
    ''' 
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function MyCP_QueueProcess_Cancel(ByVal Sn As Integer, ByVal UserID As String) As ReturnRE
        ConnStr(14)
        'Dim Conn As New SqlConnection
        'Conn = lai.open14Db()
        ' 建立命令(預存程序名,使用連線) 
        Dim Command As New SqlClient.SqlCommand("MyCard_Backup.dbo.MyCP_QueueProcess_Cancel", conn)
        ' 設置命令型態為預存程序 
        Command.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位 
        Command.Parameters.Add("@Sn", SqlDbType.Int).Value = Sn
        Command.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserID
        Dim ReturnNo As SqlClient.SqlParameter = Command.Parameters.Add("@ReturnNo", SqlDbType.Int)
        ReturnNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 20)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnRE As New ReturnRE
        Try
            conn.Open()
            Command.ExecuteNonQuery()
            ReturnRE.ReturnMsgNo = ReturnNo.Value
            ReturnRE.ReturnMsg = ReturnMsg.Value
        Catch ex As Exception
            ReturnRE.ReturnMsgNo = -99
            ReturnRE.ReturnMsg = "-9Cancel差異報表補跑與取消功能時：" & ex.Message
            Return ReturnRE
        Finally
            conn.Close()
        End Try
        Return ReturnRE
    End Function
    ''' <summary>
    ''' 付費廠商品項管理
    ''' 
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_BK_StorePoints_FactoryService_View(ByVal FactoryId As String, ByVal Factory As String, ByVal Price As Integer, ByVal DeptSn As Integer, ByVal RiskManagement As Integer) As DataSet 'View_BK_StorePoints_FactoryService
        ConnStr(12)
        Dim cmd As New SqlClient.SqlCommand
        cmd.Connection = conn
        cmd.CommandType = CommandType.Text
        cmd.CommandTimeout = 300
        Dim Adp As SqlClient.SqlDataAdapter
        If FactoryId = "0" And Factory = "" And Price = -1 Then
            Dim NullDATA As DataTable = New DataTable("NullDATA")
            Dim DBViewDS As New DataSet
            DBViewDS.Tables.Add(NullDATA)
            Return DBViewDS
        Else
            cmd.CommandText = "Select distinct  Sn, FactoryId, FactoryServiceId,FactoryServiceName_All, FactoryServiceName, FactoryServiceType, ProductNo, Price , ServiceFee, Status, ItemOrder, MinimumPrice, RiskManagement From PointsBilling_DB.dbo.View_BK_StorePoints_FactoryService_dept where 1=1 and  DeptSn=@DeptSn  "
            cmd.Parameters.Add("@DeptSN", SqlDbType.Int).Value = DeptSn
            If FactoryId <> "" Then
                cmd.CommandText = cmd.CommandText & " and FactoryId=@FactoryId "
                cmd.Parameters.Add("@FactoryId", SqlDbType.VarChar, 100).Value = FactoryId

            End If

            If Factory <> "" Then
                cmd.CommandText = cmd.CommandText & " and  FactoryServiceName LIKE '%" & Factory & "%' "

            End If
            If Price <> -1 Then
                cmd.CommandText = cmd.CommandText & " and  Price=@Price "
                cmd.Parameters.Add("@Price", SqlDbType.Int, 10).Value = Price

            End If
            If RiskManagement <> -1 Then
                cmd.CommandText = cmd.CommandText & " and  RiskManagement=@RiskManagement "
                cmd.Parameters.Add("@RiskManagement", SqlDbType.Int).Value = RiskManagement

            End If

            cmd.CommandText = cmd.CommandText & " Order by ItemOrder "


            Adp = New SqlClient.SqlDataAdapter(cmd)
            'Adp = New SqlClient.SqlDataAdapter
            Dim DBViewDS As New DataSet
            Try
                Adp.Fill(DBViewDS)
            Catch ex As Exception
                Throw New Exception("Service的View條件資料失敗", ex)
            End Try
            Return DBViewDS
        End If
    End Function 'DB連線 View
    ''' <summary>
    ''' 付費廠商品項管理
    ''' 
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_BK_StorePoints_Factory_View(ByVal DeptSn As Integer) As DataSet 'View_BK_StorePoints_Factory
        ConnStr(12)
        Dim cmd As New SqlClient.SqlCommand
        cmd.Connection = conn
        cmd.CommandType = CommandType.Text
        cmd.CommandTimeout = 300
        cmd.CommandText = "Select FactoryId, FactoryName From PointsBilling_DB.dbo.View_BK_StorePoints_Factory_dept where DeptSn=@DeptSn Order by FactoryName"
        cmd.Parameters.Add("@DeptSN", SqlDbType.Int).Value = DeptSn
        Dim Adp As SqlClient.SqlDataAdapter
        Adp = New SqlClient.SqlDataAdapter(cmd)

        Dim DBViewDS As New DataSet
        Try
            Adp.Fill(DBViewDS)
        Catch ex As Exception
            Throw New Exception("Service的View條件資料失敗", ex)
        End Try
        Return DBViewDS
    End Function  'DB連線 View  DDL
    ''' <summary>
    ''' 付費廠商品項管理
    ''' 
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_SPS_StorePoints_FactoryService_Insert(ByVal FactoryId As String, ByVal FactoryServiceId As String, ByVal FactoryServiceName As String, ByVal FactoryServiceType As String, ByVal ProductNo As String, ByVal Price As Integer,
                                                                                                             ByVal MinimumPrice As Integer, ByVal ItemOrder As Integer, ByVal ServiceFee As Double, ByVal CreateUser As String, ByVal RiskManagement As Integer) As GetSPSStorePointsFactoryServiceValue
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線) 
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_FactoryService_Insert", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位 
        cmd.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        cmd.Parameters.Add("@FactoryServiceId", SqlDbType.VarChar, 20).Value = FactoryServiceId
        cmd.Parameters.Add("@FactoryServiceName", SqlDbType.VarChar, 30).Value = FactoryServiceName
        cmd.Parameters.Add("@FactoryServiceType", SqlDbType.TinyInt).Value = FactoryServiceType
        cmd.Parameters.Add("@ProductNo", SqlDbType.VarChar, 20).Value = ProductNo
        cmd.Parameters.Add("@Price", SqlDbType.Int).Value = Price
        cmd.Parameters.Add("@MinimumPrice", SqlDbType.Int).Value = MinimumPrice
        cmd.Parameters.Add("@ItemOrder", SqlDbType.Int).Value = ItemOrder
        cmd.Parameters.Add("@ServiceFee", SqlDbType.Decimal).Value = ServiceFee
        cmd.Parameters.Add("@CreateUser", SqlDbType.VarChar, 30).Value = CreateUser
        '20140710 qq 新增是否有風控
        cmd.Parameters.Add("@RiskManagement", SqlDbType.Int).Value = RiskManagement
        cmd.CommandTimeout = 300
        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnValue As New GetSPSStorePointsFactoryServiceValue
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnMsg = ReturnMsg.Value
        Catch ex As Exception
            Throw New Exception("Service的付費廠商Insert條件資料失敗", ex)
        Finally
            conn.Close()
        End Try
        Return ReturnValue
    End Function 'Insert
    ''' <summary>
    ''' 付費廠商品項管理
    ''' 
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function Get_SPS_StorePoints_FactoryService_Update(ByVal Sn As Integer, ByVal FactoryId As String, ByVal FactoryServiceId As String, ByVal FactoryServiceName As String, ByVal FactoryServiceType As Integer, ByVal ProductNo As String,
                                                                                                               ByVal Price As Integer, ByVal MinimumPrice As Integer, ByVal Status As Integer, ByVal ItemOrder As Integer, ByVal ServiceFee As Double, ByVal CreateUser As String, ByVal RiskManagement As Integer) As GetSPSStorePointsFactoryServiceValue
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線) 
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_FactoryService_Update", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位 
        cmd.Parameters.Add("@Sn", SqlDbType.Int).Value = Sn
        cmd.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        cmd.Parameters.Add("@FactoryServiceId", SqlDbType.VarChar, 30).Value = FactoryServiceId
        cmd.Parameters.Add("@FactoryServiceName", SqlDbType.VarChar, 30).Value = FactoryServiceName
        cmd.Parameters.Add("@FactoryServiceType", SqlDbType.TinyInt).Value = FactoryServiceType
        cmd.Parameters.Add("@ProductNo", SqlDbType.VarChar, 20).Value = ProductNo
        cmd.Parameters.Add("@Price", SqlDbType.Int).Value = Price
        cmd.Parameters.Add("@MinimumPrice", SqlDbType.Int).Value = MinimumPrice
        cmd.Parameters.Add("@Status", SqlDbType.TinyInt).Value = Status
        cmd.Parameters.Add("@ItemOrder", SqlDbType.Int).Value = ItemOrder
        cmd.Parameters.Add("@ServiceFee", SqlDbType.Decimal).Value = ServiceFee
        cmd.Parameters.Add("@CreateUser", SqlDbType.VarChar, 30).Value = CreateUser
        '20140710 qq 新增是否有風控
        cmd.Parameters.Add("@RiskManagement", SqlDbType.Int).Value = RiskManagement
        cmd.CommandTimeout = 300
        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnValue As New GetSPSStorePointsFactoryServiceValue
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnMsg = ReturnMsg.Value
        Catch ex As Exception
            Throw New Exception("Service的付費廠商Update條件資料失敗", ex)
        Finally
            conn.Close()
        End Try
        Return ReturnValue
    End Function 'Update
    ''' <summary>
    ''' 付費廠商品項管理
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_FactoryService_Check(ByVal FactoryId As String) As GetSPSStorePointsFactoryServiceValue
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線) 
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_FactoryService_Check", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位 

        cmd.Parameters.Add("@FactoryServiceId", SqlDbType.VarChar, 30).Value = FactoryId

        cmd.CommandTimeout = 300
        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnValue As New GetSPSStorePointsFactoryServiceValue
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnMsg = ReturnMsg.Value
        Catch ex As Exception
            Throw New Exception("Service的付費廠商SPS_StorePoints_FactoryService_Check條件資料失敗", ex)
        Finally
            conn.Close()
        End Try
        Return ReturnValue
    End Function
    ''' <summary>
    ''' Billing後台/CP對應付費方式批次修改
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function View_BK_StorePoints_GameFactory(ByVal DeptSn As Integer) As DataSet
        ConnStr(12)
        Dim SQlstring As String = ""
        Dim cmd As New SqlClient.SqlCommand("Select GameFactoryName, GameFactoryId from PointsBilling_db.dbo.View_BK_StorePoints_GameFactory_dept where DeptSn=@DeptSn ", conn)
        cmd.CommandType = CommandType.Text
        cmd.Parameters.Add("@DeptSN", SqlDbType.Int).Value = DeptSn
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Try
            conn.Open()
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            conn.Close()
        End Try
        Return DSssid
    End Function
    ''' <summary>
    ''' Billing後台/CP對應付費方式批次修改
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function View_BK_StorePoints_GameService(ByVal GameFactoryId As String) As DataSet
        ConnStr(12)
        Dim SQlstring As String = ""
        Dim cmd As New SqlClient.SqlCommand("Select GameServiceId, GameName from PointsBilling_db.dbo.View_BK_StorePoints_GameService where GameFactoryId=@GameFactoryId ", conn)
        cmd.CommandType = CommandType.Text
        cmd.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 30).Value = GameFactoryId

        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Try
            conn.Open()
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            conn.Close()
        End Try
        Return DSssid
    End Function
    ''' <summary>
    ''' Billing後台/CP對應付費方式批次修改
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function View_StorePoints_CPAndFactoryLimit(ByVal GameServiceId As String, ByVal DeptSn As Integer) As DataSet
        ConnStr(12)
        Dim SQlstring As String = ""
        Dim cmd As New SqlClient.SqlCommand("Select DISTINCT FactoryId, FactoryName from PointsBilling_db.dbo.View_StorePoints_CPAndFactoryLimit where GameServiceId=@GameServiceId and DeptSn=@DeptSn", conn)
        cmd.CommandType = CommandType.Text
        cmd.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        cmd.Parameters.Add("@DeptSN", SqlDbType.Int).Value = DeptSn
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Try
            conn.Open()
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            conn.Close()
        End Try
        Return DSssid
    End Function
    ''' <summary>
    ''' Billing後台/CP對應付費方式批次修改
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function View_StorePoints_CPAndFactoryLimitSN(ByVal FactoryId As String, ByVal GameServiceId As String, ByVal DeptSn As Integer) As DataSet
        ConnStr(12)
        Dim SQlstring As String = ""
        Dim cmd As New SqlClient.SqlCommand("Select sn,IncludeBadAccountStatus from PointsBilling_db.dbo.View_StorePoints_CPAndFactoryLimit where FactoryId=@FactoryId and GameServiceId=@GameServiceId and DeptSn=@DeptSn", conn)
        cmd.CommandType = CommandType.Text
        cmd.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        cmd.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        cmd.Parameters.Add("@DeptSN", SqlDbType.Int).Value = DeptSn
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Try
            conn.Open()
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            conn.Close()
        End Try
        Return DSssid
    End Function
    ''' <summary>
    ''' Billing後台/CP對應付費方式批次修改
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' 
    ''' </remarks>
    <WebMethod()> _
    Public Function MyCard_ADSell_ProgramSchedule_Insert(ByVal Sn As Integer, ByVal FactoryId As String, ByVal GameFactoryId As String, ByVal GameServiceId As String, ByVal UpperFristBound As Integer, ByVal DayUpperMobileBound As Integer, ByVal MonthUpperMobileBound As Integer, ByVal DayUpperMailBound As Integer, ByVal MonthUpperMailBound As Integer, ByVal DayUpperCustIdBound As Integer, ByVal MonthUpperCustIdBound As Integer, ByVal TempCodeLimit As Integer, ByVal IPTradeLimit As Integer, ByVal IncludeBadAccountStatus As Integer, ByVal UserStamp As String) As ReturnValue
        Dim ReturnValue As New ReturnValue
        ConnStr(12)
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_db.dbo.SPS_StorePoints_CPAndFactoryLimit_Update", conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add("@Sn", SqlDbType.Int).Value = Sn
        cmd.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        cmd.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = GameFactoryId
        cmd.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        cmd.Parameters.Add("@UpperFristBound", SqlDbType.Int).Value = UpperFristBound
        cmd.Parameters.Add("@DayUpperMobileBound", SqlDbType.Int).Value = DayUpperMobileBound
        cmd.Parameters.Add("@MonthUpperMobileBound", SqlDbType.Int).Value = MonthUpperMobileBound
        cmd.Parameters.Add("@DayUpperMailBound", SqlDbType.Int).Value = DayUpperMailBound
        cmd.Parameters.Add("@MonthUpperMailBound", SqlDbType.Int).Value = MonthUpperMailBound
        cmd.Parameters.Add("@DayUpperCustIdBound", SqlDbType.Int).Value = DayUpperCustIdBound
        cmd.Parameters.Add("@MonthUpperCustIdBound", SqlDbType.Int).Value = MonthUpperCustIdBound
        cmd.Parameters.Add("@TempCodeLimit", SqlDbType.TinyInt).Value = TempCodeLimit
        cmd.Parameters.Add("@IPTradeLimit", SqlDbType.TinyInt).Value = IPTradeLimit
        cmd.Parameters.Add("@IncludeBadAccountStatus", SqlDbType.TinyInt).Value = IncludeBadAccountStatus
        cmd.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserStamp
        cmd.CommandTimeout = 300

        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsg = ReturnMsg.Value
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        Return ReturnValue
    End Function

    ''' <summary>
    ''' Billing後台/付費方式對應CP批次修改
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetFactoryForCPAndFactoryLimit(ByVal DeptSn As Integer) As DataSet
        ConnStr(12)
        Dim command As New SqlClient.SqlCommand
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        command.CommandText = "SELECT FactoryId, FactoryName FROM [PointsBilling_DB].[dbo].[View_StorePoints_CPAndFactoryLimit] where DeptSn=@DeptSn Group by FactoryId, FactoryName order by FactoryId, FactoryName asc"
        command.Parameters.Add("@DeptSN", SqlDbType.Int).Value = DeptSn

        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("查詢遊戲廠商遊戲介接資料失敗：" & ex.Message)
        End Try
        Return DSssid
    End Function
    ''' <summary>
    ''' Billing後台/付費方式對應CP批次修改
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetFactoryByFactoryIdForCPAndFactoryLimit(ByVal FactoryId As String, ByVal DeptSn As Integer) As DataSet
        ConnStr(12)
        Dim Cmd As New SqlClient.SqlCommand
        Cmd.Connection = conn
        Cmd.CommandType = CommandType.Text
        Cmd.CommandTimeout = 300
        Cmd.CommandText = "SELECT GameServiceId, GameName FROM [PointsBilling_DB].[dbo].[View_StorePoints_CPAndFactoryLimit] where FactoryId=@FactoryId and DeptSn=@DeptSn Group by GameServiceId, GameName order by GameServiceId, GameName asc"
        Cmd.Parameters.Add("@FactoryId", SqlDbType.VarChar).Value = FactoryId
        Cmd.Parameters.Add("@DeptSN", SqlDbType.Int).Value = DeptSn
        Dim Adp As SqlDataAdapter = New SqlDataAdapter(Cmd)
        Dim Ds As DataSet = New DataSet()
        Try
            conn.Open()
            Adp.Fill(Ds)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            conn.Close()
        End Try
        Return Ds
    End Function
    ''' <summary>
    ''' Billing後台/付費方式對應CP批次修改
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' 
    ''' </remarks>
    <WebMethod()> _
    Public Function FactorInfoQuery(ByVal factoryId As String, ByVal gameServiceId As String, ByVal DeptSn As Integer) As GetDataResult
        ConnStr(12)
        Dim Cmd As New SqlClient.SqlCommand
        Cmd.Connection = conn
        Cmd.CommandType = CommandType.Text
        Cmd.CommandTimeout = 300
        Cmd.CommandText = "Select Sn, FactoryId, GameFactoryId, GameServiceId From PointsBilling_DB.dbo.View_StorePoints_CPAndFactoryLimit Where FactoryId=@FactoryId and GameServiceId=@GameServiceId and DeptSn=@DeptSn order by Sn asc"
        Cmd.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = factoryId
        Cmd.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = gameServiceId
        Cmd.Parameters.Add("@DeptSN", SqlDbType.Int).Value = DeptSn
        Dim Adp As New System.Data.SqlClient.SqlDataAdapter(Cmd)
        Dim Ds As New DataSet
        Dim GetDataResult As New GetDataResult
        Try
            Adp.Fill(Ds)
            GetDataResult.ReturnMsgNo = 1
            GetDataResult.ReturnMsg = "查詢成功！"
            GetDataResult.ReturnDS = Ds
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return GetDataResult
    End Function
    ''' <summary>
    ''' Billing後台/付費方式對應CP批次修改
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' 
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_CPAndFactoryLimit_Update(ByVal IntSn As Integer, ByVal StrFactoryId As String, ByVal StrGameFactoryId As String, _
                                                             ByVal StrGameServiceId As String, ByVal IntUpperFristBound As Integer, ByVal IntDayUpperMobileBound As Integer, _
                                                             ByVal IntMonthUpperMobileBound As Integer, ByVal IntDayUpperMailBound As Integer, ByVal IntMonthUpperMailBound As Integer, _
                                                             ByVal IntDayUpperCustIdBound As Integer, ByVal IntMonthUpperCustIdBound As Integer, ByVal IntTempCodeLimit As Integer, _
                                                             ByVal IntIPTradeLimit As Integer, _
                                                             ByVal IntSn1 As Integer, ByVal StrFactoryId1 As String, ByVal StrGameFactoryId1 As String, _
                                                             ByVal StrGameServiceId1 As String, ByVal IntUpperFristBound1 As Integer, ByVal IntDayUpperMobileBound1 As Integer, _
                                                             ByVal IntMonthUpperMobileBound1 As Integer, ByVal IntDayUpperMailBound1 As Integer, ByVal IntMonthUpperMailBound1 As Integer, _
                                                             ByVal IntDayUpperCustIdBound1 As Integer, ByVal IntMonthUpperCustIdBound1 As Integer, ByVal IntTempCodeLimit1 As Integer, _
                                                             ByVal IntIPTradeLimit1 As Integer, _
                                                             ByVal StrUserName As String) As ReturnRE
        Dim Msg As String = ""
        Dim ReturnValue As New ReturnRE
        ConnStr(12)
        BeginTrans(Msg)

        ' 建立命令(預存程序名,使用連線) 
        Dim Cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_CPAndFactoryLimit_Update", conn)
        ' 設置命令型態為預存程序 
        Cmd.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位 
        Cmd.Parameters.Add("@Sn", SqlDbType.Int).Value = IntSn
        Cmd.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = StrFactoryId
        Cmd.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = StrGameFactoryId
        Cmd.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = StrGameServiceId
        Cmd.Parameters.Add("@UpperFristBound", SqlDbType.Int).Value = IntUpperFristBound
        Cmd.Parameters.Add("@DayUpperMobileBound", SqlDbType.Int).Value = IntDayUpperMobileBound
        Cmd.Parameters.Add("@MonthUpperMobileBound", SqlDbType.Int).Value = IntMonthUpperMobileBound
        Cmd.Parameters.Add("@DayUpperMailBound", SqlDbType.Int).Value = IntDayUpperMailBound
        Cmd.Parameters.Add("@MonthUpperMailBound", SqlDbType.Int).Value = IntMonthUpperMailBound
        Cmd.Parameters.Add("@DayUpperCustIdBound", SqlDbType.Int).Value = IntDayUpperCustIdBound
        Cmd.Parameters.Add("@MonthUpperCustIdBound", SqlDbType.Int).Value = IntMonthUpperCustIdBound
        Cmd.Parameters.Add("@TempCodeLimit", SqlDbType.Int).Value = IntTempCodeLimit
        Cmd.Parameters.Add("@IPTradeLimit", SqlDbType.Int).Value = IntIPTradeLimit
        Cmd.Parameters.Add("@IncludeBadAccountStatus", SqlDbType.Int).Value = 1
        Cmd.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = StrUserName
        Cmd.CommandTimeout = 300

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Try
            ConnOpen(Cmd)
            Cmd.ExecuteNonQuery()
        Catch ex As Exception

            ReturnMsgNo.Value = -99
            ReturnMsg.Value = "Update含呆額度時發生錯誤：" & ex.Message

        Finally
            ConnClose()
        End Try
        If ReturnMsgNo.Value <> 1 Then
            RollbackTrans(ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnMsg = ReturnMsg.Value
            Return ReturnValue
        End If

        ' 建立命令(預存程序名,使用連線) 
        Dim Cmd1 As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_CPAndFactoryLimit_Update", conn)
        ' 設置命令型態為預存程序 
        Cmd1.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位 
        Cmd1.Parameters.Add("@Sn", SqlDbType.Int).Value = IntSn1
        Cmd1.Parameters.Add("@FactoryId", SqlDbType.VarChar, 20).Value = StrFactoryId1
        Cmd1.Parameters.Add("@GameFactoryId", SqlDbType.VarChar, 20).Value = StrGameFactoryId1
        Cmd1.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = StrGameServiceId1
        Cmd1.Parameters.Add("@UpperFristBound", SqlDbType.Int).Value = IntUpperFristBound1
        Cmd1.Parameters.Add("@DayUpperMobileBound", SqlDbType.Int).Value = IntDayUpperMobileBound1
        Cmd1.Parameters.Add("@MonthUpperMobileBound", SqlDbType.Int).Value = IntMonthUpperMobileBound1
        Cmd1.Parameters.Add("@DayUpperMailBound", SqlDbType.Int).Value = IntDayUpperMailBound1
        Cmd1.Parameters.Add("@MonthUpperMailBound", SqlDbType.Int).Value = IntMonthUpperMailBound1
        Cmd1.Parameters.Add("@DayUpperCustIdBound", SqlDbType.Int).Value = IntDayUpperCustIdBound1
        Cmd1.Parameters.Add("@MonthUpperCustIdBound", SqlDbType.Int).Value = IntMonthUpperCustIdBound1
        Cmd1.Parameters.Add("@TempCodeLimit", SqlDbType.Int).Value = IntTempCodeLimit1
        Cmd1.Parameters.Add("@IPTradeLimit", SqlDbType.Int).Value = IntIPTradeLimit1
        Cmd1.Parameters.Add("@IncludeBadAccountStatus", SqlDbType.Int).Value = 0
        Cmd1.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = StrUserName
        Cmd1.CommandTimeout = 300

        Dim ReturnMsgNo1 As SqlClient.SqlParameter = Cmd1.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo1.Direction = ParameterDirection.Output
        Dim ReturnMsg1 As SqlClient.SqlParameter = Cmd1.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg1.Direction = ParameterDirection.Output
        Try
            ConnOpen(Cmd1)
            Cmd1.ExecuteNonQuery()
        Catch ex As Exception
            ReturnMsgNo1.Value = -99
            ReturnMsg1.Value = "Update不含呆額度時發生錯誤：" & ex.Message

        Finally
            ConnClose()
        End Try
        If ReturnMsgNo1.Value <> 1 Then
            RollbackTrans(ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = ReturnMsgNo1.Value
            ReturnValue.ReturnMsg = ReturnMsg1.Value
            Return ReturnValue
        End If
        CommitTrans(Msg)
        ReturnValue.ReturnMsgNo = ReturnMsgNo1.Value
        ReturnValue.ReturnMsg = ReturnMsg1.Value
        Return ReturnValue
    End Function
    ''' <summary>
    ''' MyCard帳務後台/差異報表處理功能修改
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' 
    ''' </remarks>
    <WebMethod()> _
    Public Function MyCP_SaveDifferenceData_Update(ByVal Sn As Integer, ByVal ReasonTypeSn As Integer, ByVal ProceStatus As Integer, ByVal UserID As String) As ReturnValue
        Dim ReturnValue As New ReturnValue
        ConnStr(14)
        ' 建立命令(預存程序名,使用連線) 
        Dim Command As New SqlClient.SqlCommand("MyCard_Backup.dbo.MyCP_SaveDifferenceData_Update", conn)
        ' 設置命令型態為預存程序 
        Command.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位 
        Command.Parameters.Add("@Sn", SqlDbType.Int).Value = Sn
        Command.Parameters.Add("@ReasonTypeSn", SqlDbType.Int).Value = ReasonTypeSn
        Command.Parameters.Add("@ProceStatus", SqlDbType.Int).Value = ProceStatus
        Command.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserID
        Dim ReturnNo As SqlClient.SqlParameter = Command.Parameters.Add("@ReturnNo", SqlDbType.Int)
        ReturnNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 20)
        ReturnMsg.Direction = ParameterDirection.Output
        Try
            conn.Open()
            Command.ExecuteNonQuery()
        Catch ex As Exception
            ReturnValue.ReturnMsg = "-7修改差異報表處理功能時：" & ex.Message
            ReturnValue.ReturnMsgNo = -99
            Return ReturnValue
        Finally
            conn.Close()
        End Try
        ReturnValue.ReturnMsg = ReturnMsg.Value
        ReturnValue.ReturnMsgNo = ReturnNo.Value
        Return ReturnValue
    End Function

    ''' <summary>
    ''' MyCard Billing後台/HappyGo/處理多筆兌點退貨結果檔
    ''' 資料查詢(4:產生匯入檔)
    ''' </summary>
    ''' <param name="CancelDate"></param>
    ''' <param name="CancelStatus">4:產生匯入檔</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    <WebMethod()> _
    Public Function VIEW_StorePoints_HomeReceivable_RollBackPoint(ByVal CancelDate As Date, ByVal CancelStatus As Integer) As DataSet
        ConnStr(12)
        Dim SQlstring As String = ""
        Dim cmd As New SqlClient.SqlCommand("Select SN,TradeSeq from PointsBilling_db.dbo.VIEW_HappyGoPo_TradeList_Query where CancelDate<=@CancelEndDate and CancelDate>@CancelStartDate and CancelStatus=@CancelStatus", conn)
        cmd.CommandType = CommandType.Text
        cmd.Parameters.Add("@CancelStartDate", SqlDbType.VarChar, 30).Value = CancelDate.ToString("yyyy/MM/dd") & " 00:00:00"
        cmd.Parameters.Add("@CancelEndDate", SqlDbType.VarChar, 30).Value = CancelDate.AddDays(1).AddMilliseconds(-1).ToString("yyyy/MM/dd HH:mm:sss")
        cmd.Parameters.Add("@CancelStatus", SqlDbType.Int).Value = CancelStatus
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Try
            conn.Open()
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            conn.Close()
        End Try
        Return DSssid
    End Function

    ''' <summary>
    ''' 2013/10/29 qq 新的SP
    ''' MyCard Billing後台/HappyGo/處理多筆兌點退貨結果檔
    ''' 資料查詢(4:產生匯入檔)
    ''' </summary>
    ''' <param name="CancelDate"></param>
    ''' <param name="CancelStatus">4:產生匯入檔</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_HappyGoPo_TradeListCancel_Query(ByVal CancelDate As Date, ByVal CancelStatus As Integer) As PublicReturnRE
        ConnStr(14)
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_HappyGoPo_TradeListCancel_Query", conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 16).Value = ""
        cmd.Parameters.Add("@CancelStartDate", SqlDbType.VarChar, 30).Value = CancelDate.ToString("yyyy/MM/dd") & " 00:00:00"
        cmd.Parameters.Add("@CancelEndDate", SqlDbType.VarChar, 30).Value = CancelDate.AddDays(1).ToString("yyyy/MM/dd")
        cmd.Parameters.Add("@Cancel_Status", SqlDbType.Int).Value = CancelStatus

        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 100)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim PublicReturnRE As New PublicReturnRE

        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Try
            conn.Open()
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            conn.Close()
        End Try
        PublicReturnRE.ReturnMsgNo = 1
        PublicReturnRE.ReturnMsg = "查詢成功"
        PublicReturnRE.ReturnDS = DSssid
        Return PublicReturnRE
    End Function

    ''' <summary>
    ''' MyCard Billing後台/HappyGo/處理多筆兌點退貨結果檔
    ''' 累積退點_取消狀態修改  3：取消確認
    ''' </summary>
    ''' <param name="SN"></param>
    ''' <param name="CancelStatus">取消標記：1：一般取消 2：批次取消 3：取消確認</param>
    ''' <param name="CancelUser"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_HappyGoPointData_TradeList_CancelStatus_A01(ByVal SN As Integer, ByVal CancelStatus As Integer, ByVal CancelUser As String) As ReturnRE
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線) 
        Dim Command As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_HappyGoPo_TradeList_CancelStatus", conn)
        ' 設置命令型態為預存程序 
        Command.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位 
        Command.Parameters.Add("@SN", SqlDbType.Int).Value = SN
        Command.Parameters.Add("@CancelStatus", SqlDbType.Int).Value = CancelStatus
        Command.Parameters.Add("@CancelUser", SqlDbType.VarChar, 30).Value = CancelUser

        Dim ReturnMsgNo As SqlClient.SqlParameter = Command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 50)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnRE As New ReturnRE
        Try
            conn.Open()
            Command.ExecuteNonQuery()
            ReturnRE.ReturnMsgNo = ReturnMsgNo.Value
            ReturnRE.ReturnMsg = ReturnMsg.Value
        Catch ex As Exception
            ReturnRE.ReturnMsgNo = -99
            ReturnRE.ReturnMsg = ex.Message
            Return ReturnRE
        Finally
            conn.Close()
        End Try
        Return ReturnRE
    End Function
    ''' <summary>
    ''' MyCard Billing後台/Billing後台維護作業/付費廠商對應單位別功能
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    <WebMethod()> _
    Public Function View_MessageServices_Group() As PublicReturnRE
        ConnStr(14)
        Dim SQlstring As String = ""
        Dim cmd As New SqlClient.SqlCommand("Select Sn, GroupName from Mycard_backup.dbo.View_MessageServices_Group  ORDER BY  GroupName ", conn)
        cmd.CommandTimeout = 300
        cmd.CommandType = CommandType.Text

        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Dim PublicReturnRE As New PublicReturnRE
        Try
            conn.Open()
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            'Throw New Exception(ex.Message)
            PublicReturnRE.ReturnMsgNo = -99
            PublicReturnRE.ReturnMsg = "執行View_MessageServices_Group時發生錯誤" & ex.ToString
        Finally
            conn.Close()
        End Try
        PublicReturnRE.ReturnMsgNo = 1
        PublicReturnRE.ReturnMsg = "查詢成功"
        PublicReturnRE.ReturnDS = DSssid
        Return PublicReturnRE
    End Function
    ''' <summary>
    ''' MyCard Billing後台/Billing後台維護作業/付費廠商對應單位別功能
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    <WebMethod()> _
    Public Function VIEW_MessageService_GROUPvsFACTORY(ByVal MessageGroupSn As String) As PublicReturnRE
        ConnStr(14)
        Dim SQlstring As String = ""
        Dim cmd As New SqlClient.SqlCommand
        cmd.CommandText = "Select SN, MessageGroupSn, FactoryID, FactoryName, Status from Mycard_backup.dbo.VIEW_MessageService_GROUPvsFACTORY Where 1=1 "
        If MessageGroupSn <> "" Then
            cmd.CommandText &= " And MessageGroupSn=@MessageGroupSn and Status=1 Order by FactoryName "
            cmd.Parameters.AddWithValue("@MessageGroupSn", MessageGroupSn)
        Else
            cmd.CommandText &= " And Status=1 Order by FactoryName "
        End If

        cmd.Connection = conn
        cmd.CommandTimeout = 300
        cmd.CommandType = CommandType.Text


        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Dim PublicReturnRE As New PublicReturnRE
        Try
            conn.Open()
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            'Throw New Exception(ex.Message)
            PublicReturnRE.ReturnMsgNo = -99
            PublicReturnRE.ReturnMsg = "執行VIEW_MessageService_GROUPvsFACTORY時發生錯誤" & ex.ToString
        Finally
            conn.Close()
        End Try
        PublicReturnRE.ReturnMsgNo = 1
        PublicReturnRE.ReturnMsg = "查詢成功"
        PublicReturnRE.ReturnDS = DSssid
        Return PublicReturnRE
    End Function
    ''' <summary>
    ''' MyCard Billing後台/Billing後台維護作業/付費廠商對應單位別功能
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    <WebMethod()> _
    Public Function VIEW_MessageService_GROUPvsFACTORYSearch(ByVal MessageGroupSn As String, ByVal FactoryID As String) As PublicReturnRE
        ConnStr(14)
        Dim SQlstring As String = ""
        Dim cmd As New SqlClient.SqlCommand
        cmd.CommandText = "Select SN, MessageGroupSn, FactoryID, FactoryName, Status from Mycard_backup.dbo.VIEW_MessageService_GROUPvsFACTORY Where 1=1 "
        If MessageGroupSn <> "" Then
            cmd.CommandText &= " And MessageGroupSn=@MessageGroupSn and FactoryID=@FactoryID Order by FactoryName "
            cmd.Parameters.AddWithValue("@MessageGroupSn", MessageGroupSn)
            cmd.Parameters.AddWithValue("@FactoryID", FactoryID)
        Else
            cmd.CommandText &= " Order by FactoryName "
        End If

        cmd.Connection = conn
        cmd.CommandTimeout = 300
        cmd.CommandType = CommandType.Text


        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Dim PublicReturnRE As New PublicReturnRE
        Try
            conn.Open()
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            'Throw New Exception(ex.Message)
            PublicReturnRE.ReturnMsgNo = -99
            PublicReturnRE.ReturnMsg = "執行VIEW_MessageService_GROUPvsFACTORY時發生錯誤" & ex.ToString
        Finally
            conn.Close()
        End Try
        PublicReturnRE.ReturnMsgNo = 1
        PublicReturnRE.ReturnMsg = "查詢成功"
        PublicReturnRE.ReturnDS = DSssid
        Return PublicReturnRE
    End Function
    ''' <summary>
    ''' MyCard Billing後台/Billing後台維護作業/付費廠商對應單位別功能
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    <WebMethod()> _
    Public Function MessageService_GROUPvsFACTORY_Setting(ByVal SN As Integer, ByVal MessageGroupSn As Integer, ByVal FactoryID As String, ByVal Status As Integer, ByVal CancelUser As String) As PublicReturnRE
        ConnStr(14)
        Dim PublicReturnRE As New PublicReturnRE

        If Status = 1 Then
            Dim SQlstring As String = ""
            Dim cmd As New SqlClient.SqlCommand("Select SN, MessageGroupSn, FactoryID, FactoryName, Status from Mycard_backup.dbo.VIEW_MessageService_GROUPvsFACTORY Where FactoryID=@FactoryID and Status=1 Order by FactoryName ", conn)
            cmd.CommandTimeout = 300
            cmd.CommandType = CommandType.Text
            cmd.Parameters.AddWithValue("@FactoryID", FactoryID)

            Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
            Dim DSssid As New DataSet
            'Dim PublicReturnRE As New PublicReturnRE
            Try
                conn.Open()
                SDAssid.Fill(DSssid)
            Catch ex As Exception
                'Throw New Exception(ex.Message)
                PublicReturnRE.ReturnMsgNo = -99
                PublicReturnRE.ReturnMsg = "執行VIEW_MessageService_GROUPvsFACTORY時發生錯誤" & ex.ToString
            Finally
                conn.Close()
            End Try
            If DSssid.Tables.Count <> 0 Then
                If DSssid.Tables(0).Rows.Count <> 0 Then
                    PublicReturnRE.ReturnMsgNo = -98
                    PublicReturnRE.ReturnMsg = "已被其他群組設定過"
                    PublicReturnRE.ReturnDS = DSssid
                    Return PublicReturnRE
                End If
            End If

        End If

        ' 建立命令(預存程序名,使用連線) 
        Dim Command As New SqlClient.SqlCommand("Mycard_backup.dbo.MessageService_GROUPvsFACTORY_Setting", conn)
        ' 設置命令型態為預存程序 
        Command.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位 
        Command.Parameters.Add("@SN", SqlDbType.Int).Value = SN
        Command.Parameters.Add("@MessageGroupSn", SqlDbType.Int).Value = MessageGroupSn
        Command.Parameters.Add("@FactoryID", SqlDbType.VarChar, 20).Value = FactoryID
        Command.Parameters.Add("@Status", SqlDbType.Int).Value = Status
        Command.Parameters.Add("@CreateUser", SqlDbType.VarChar, 50).Value = CancelUser

        Dim ReturnMsgNo As SqlClient.SqlParameter = Command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 50)
        ReturnMsg.Direction = ParameterDirection.Output
        'Dim PublicReturnRE As New PublicReturnRE
        Try
            conn.Open()
            Command.ExecuteNonQuery()

        Catch ex As Exception
            PublicReturnRE.ReturnMsgNo = -99
            PublicReturnRE.ReturnMsg = "執行MessageService_GROUPvsFACTORY_Setting時發生錯誤" & ex.ToString
            Return PublicReturnRE
        Finally
            conn.Close()
        End Try
        PublicReturnRE.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
        PublicReturnRE.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
        Return PublicReturnRE
    End Function
    ''' <summary>
    ''' MyCard Billing後台/Billing後台維護作業/付費廠商對應單位別功能
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function View_BK_StorePoints_Factory() As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim com As New SqlCommand()
        com.Connection = conn
        com.CommandTimeout = 300
        com.CommandText = "SELECT FactoryId, FactoryName FROM PointsBilling_DB.dbo.View_BK_StorePoints_Factory ORDER BY  FactoryName"


        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢失敗|View_BK_StorePoints_Factory|" & ex.ToString)
        Finally
            conn.Close()
        End Try

        Return ds

    End Function

    ''' <summary>
    ''' MyCard Billing後台/CP對應付費方式批次新增
    ''' 查詢報表寄送群組
    ''' </summary>
    ''' <param name="ReportId">報表ID</param>
    ''' <param name="DeptSn">單位別</param>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetVIEW_BKportal_Dept_ForReprot(ByVal ReportId As String, ByVal DeptSn As String) As DataSet
        ConnStr(14)
        Dim com As New SqlCommand()
        com.Connection = conn
        Dim sqlstr As String = "Select ReportSendGroupSn From Mycard_backup.dbo.VIEW_BKportal_Dept_ForReprot where 1=1 "
        If ReportId <> "" Then
            sqlstr = sqlstr & " and ReprotId=@ReprotId "
            com.Parameters.Add("@ReprotId", SqlDbType.VarChar, 20).Value = ReportId
        End If
        If DeptSn <> "" Then
            sqlstr = sqlstr & " and DeptSn=@DeptSn "
            com.Parameters.Add("@DeptSn", SqlDbType.Int).Value = CInt(DeptSn)
        End If

        sqlstr = sqlstr
        com.CommandText = sqlstr

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception("查詢報表寄送群組失敗|" & ex.Message)
        End Try

        If ds.Tables(0).Rows.Count <> 0 Then
            Dim GroupSn As Integer = CInt(ds.Tables(0).Rows(0).Item("ReportSendGroupSn").ToString())
            Dim MailDS As New DataSet
            Try
                MailDS = Mycard_Backup_Remind_VB_Mail_View(GroupSn)
            Catch ex As Exception
                Throw New Exception("查詢寄送人員失敗|" & ex.Message)
            End Try

            Return MailDS
        Else
            Return ds
        End If
    End Function
    <WebMethod()>
    Public Function ReturnEndCpData(ByVal GameServiceId As String) As DataTable
        ConnStr(12)
        Dim dt As New DataTable
        Dim SqlExpression As String = String.Empty
        SqlExpression = "SELECT *  FROM PointsBilling_DB.dbo.View_StorePoints_ContractDate_RemindStatus WHERE GameServiceId=@GameServiceId"

        Dim Cmd As New SqlClient.SqlCommand(SqlExpression, conn)
        Cmd.CommandType = CommandType.Text
        Cmd.CommandTimeout = 100
        Cmd.Parameters.AddWithValue("@GameServiceId", GameServiceId)
        Dim da As New SqlDataAdapter(Cmd)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            Return Nothing
        Finally
            Cmd.Connection.Close()
        End Try
        Return dt
    End Function
    ''' <summary>
    ''' 查詢付費方式列表
    ''' </summary>
    ''' <param name="GameServiceId">遊戲服務名稱</param>
    ''' <param name="ReturnMsgNo"></param>
    ''' <param name="ReturnMsg"></param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    <WebMethod()>
    Public Function GetPayWayList(ByVal GameServiceId As String, ByRef ReturnMsgNo As Integer, ByRef ReturnMsg As String) As DataTable
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線)
        Dim dt As New DataTable
        Dim Cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_CPFactorySharedRate_Query", conn)
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.Parameters.AddWithValue("@GameServiceId", GameServiceId)

        Dim pReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        pReturnMsgNo.Direction = ParameterDirection.Output
        Dim pReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NChar, 50)
        pReturnMsg.Direction = ParameterDirection.Output

        Dim da As New SqlDataAdapter(Cmd)
        Dim ds As New DataSet

        Try
            da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            Throw New Exception("查詢失敗|SPS_StorePoints_CPFactorySharedRate_Query|" & ex.ToString)
        Finally
            Cmd.Connection.Close()
        End Try
        Return dt
    End Function

    ''' <summary>
    ''' 更新廠商結束合約狀態
    ''' </summary>
    ''' <param name="CPGameId"></param>
    ''' <param name="RemindStatus"></param>
    ''' <param name="CreateUser"></param>
    ''' <param name="ReturnMsgNo"></param>
    ''' <param name="ReturnMsg"></param>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Sub Upd_RemindStautsByGameID_View(ByVal CPGameId As String, ByVal RemindStatus As Byte, ByVal CreateUser As String, ByRef ReturnMsgNo As Integer, ByRef ReturnMsg As String)
        ConnStr(11)
        ' 建立命令(預存程序名,使用連線) 
        Dim Cmd As New SqlClient.SqlCommand("Echannel_db.dbo.EC_SP_GameTypeRemindStatus_Upd", conn)
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 100

        Cmd.Parameters.AddWithValue("@P_cGAME_ID", CPGameId)
        Cmd.Parameters.AddWithValue("@P_iRemindStatus", RemindStatus)
        Cmd.Parameters.AddWithValue("@P_cCreateUser", CreateUser)

        Dim pReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@P_iReturnMsgNo", SqlDbType.Int)
        pReturnMsgNo.Direction = ParameterDirection.Output
        Dim pReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@P_cReturnMsg", SqlDbType.NChar, 50)
        pReturnMsg.Direction = ParameterDirection.Output
        Try
            Cmd.Connection.Open()
            Cmd.ExecuteNonQuery()
            Integer.TryParse(pReturnMsgNo.Value.ToString(), ReturnMsgNo)
            ReturnMsg = pReturnMsg.Value.ToString()
        Catch ex As Exception
            Throw New Exception("更新合約狀態失敗|Upd_RemindStautsByGameID_View|" & ex.Message)
        Finally
            Cmd.Connection.Close()
        End Try
    End Sub
    ''' <summary>
    ''' 更新廠商結束合約狀態(虛擬)
    ''' </summary>
    ''' <param name="FactoryID"></param>
    ''' <param name="GameFactoryID"></param>
    ''' <param name="GameServiceID"></param>
    ''' <param name="RemindStatus"></param>
    ''' <param name="CreateUser"></param>
    ''' <param name="ReturnMsgNo"></param>
    ''' <param name="ReturnMsg"></param>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Sub Upd_RemindStautsByGameID_Virtual_View(ByVal FactoryID As String, ByVal GameFactoryID As String, ByVal GameServiceID As String, ByVal RemindStatus As Byte, ByVal CreateUser As String, ByRef ReturnMsgNo As Integer, ByRef ReturnMsg As String)
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線) 
        Dim Cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_ContractDateRemindStatus", conn)
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 100

        Cmd.Parameters.AddWithValue("@P_cFactoryID", FactoryID)
        Cmd.Parameters.AddWithValue("@P_cGameFactoryID", GameFactoryID)
        Cmd.Parameters.AddWithValue("@P_cGameServiceID", GameServiceID)
        Cmd.Parameters.AddWithValue("@P_iRemindStatus", RemindStatus)
        Cmd.Parameters.AddWithValue("@P_cCreateUser", CreateUser)

        Dim pReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@P_iReturnMsgNo", SqlDbType.Int)
        pReturnMsgNo.Direction = ParameterDirection.Output
        Dim pReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@P_cReturnMsg", SqlDbType.NChar, 50)
        pReturnMsg.Direction = ParameterDirection.Output
        Try
            Cmd.Connection.Open()
            Cmd.ExecuteNonQuery()
            Integer.TryParse(pReturnMsgNo.Value.ToString(), ReturnMsgNo)
            ReturnMsg = pReturnMsg.Value.ToString()
        Catch ex As Exception
            Throw New Exception("更新合約狀態失敗|Upd_RemindStautsByGameID_View|" & ex.Message)
        Finally
            Cmd.Connection.Close()
        End Try
    End Sub


    ''' <summary>
    ''' Billing後台-付費廠商類別維護
    ''' 新增付費廠商類別
    ''' </summary>
    ''' <param name="FactoryTypeDesc"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_FactoryType_Insert_A01(ByVal FactoryTypeDesc As String, ByVal UserStamp As String, ByVal SDKImageURL As String, ByVal Priority As Integer) As InsertRE
        ConnStr(12)
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_FactoryType_Insert_A01", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位
        cmd.Parameters.Add("@FactoryTypeDesc", SqlDbType.NVarChar, 30).Value = FactoryTypeDesc
        cmd.Parameters.Add("@UserId", SqlDbType.VarChar, 30).Value = UserStamp
        cmd.Parameters.Add("@SDKImageURL", SqlDbType.VarChar, 200).Value = SDKImageURL
        cmd.Parameters.Add("@Priority", SqlDbType.Int).Value = Priority

        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnValue As New InsertRE
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("新增付費廠商類別失敗|SPS_StorePoints_FactoryType_Insert_A01|" & ex.Message)
        Finally
            conn.Close()
        End Try

        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        Return ReturnValue

    End Function


    ''' <summary>
    ''' Billing後台-付費廠商類別維護
    ''' 修改付費廠商類別
    ''' </summary>
    ''' <param name="FactoryTypeDesc"></param>
    ''' <param name="UserStamp"></param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 回傳ReturnMsgNo,ReturnMsg
    ''' </remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_FactoryType_Update_A01(ByVal Sn As Integer, ByVal FactoryTypeDesc As String, ByVal UserStamp As String, ByVal SDKImageURL As String, ByVal Priority As Integer) As UpdateRE
        ConnStr(12)
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_FactoryType_Update_A01", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        ' 增加參數欄位
        cmd.Parameters.Add("@SN", SqlDbType.Int, 4).Value = Sn
        cmd.Parameters.Add("@FactoryTypeDesc", SqlDbType.NVarChar, 30).Value = FactoryTypeDesc
        cmd.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserStamp
        cmd.Parameters.Add("@SDKImageURL", SqlDbType.VarChar, 200).Value = SDKImageURL
        cmd.Parameters.Add("@Priority", SqlDbType.Int).Value = Priority

        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ReturnValue As New UpdateRE
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("修改付費廠商類別失敗|SPS_StorePoints_FactoryType_Update_A01|" & ex.Message)
        Finally
            conn.Close()
        End Try

        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        Return ReturnValue

    End Function



    ''' <summary>
    ''' Billing後台-付費廠商類別維護
    ''' 查詢
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetView_BK_StorePoints_FactoryType_A01() As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線) 
        Dim sqlcomm As String
        sqlcomm = "select  Sn,FactoryTypeDesc,SDKImageURL , Priority  from PointsBilling_DB.dbo.View_BK_StorePoints_FactoryType_A01"
        Dim com As New SqlCommand(sqlcomm, conn)
        com.CommandTimeout = 100
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet()
        Try
            da.Fill(ds)

        Catch ex As Exception
            Throw New Exception("GetView_BK_StorePoints_FactoryType|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds
    End Function

    ''' <summary>
    ''' </summary>
    ''' <returns>ReturnNo 回傳代碼 ,ReturnMsg 回傳訊息,ReturnDs Dataset</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_EC_GAME_Type() As ReturnValue
        ConnStr(14)
        Dim Cmd As New SqlClient.SqlCommand
        Cmd.Connection = conn
        Cmd.CommandType = CommandType.Text
        Cmd.CommandTimeout = 300
        Dim Adp As SqlClient.SqlDataAdapter
        Cmd.CommandText = "Select  *  from MyCard_Backup.dbo.Admin_MyCardReport_FactoryGameType_View where 1=1 order by GameFacName  asc"

        Adp = New SqlClient.SqlDataAdapter(Cmd)
        Dim DBViewDS As New DataSet
        Dim ReturnValue As New ReturnValue
        Try
            Adp.Fill(DBViewDS)
        Catch ex As Exception
            ReturnValue.ReturnMsgNo = -99
            ReturnValue.ReturnMsg = "WebServicee的查詢Get_EC_GAME_Type-View條件資料失敗" & ex.ToString
            ReturnValue.ReturnDs = New DataSet
            Return ReturnValue
        End Try
        ReturnValue.ReturnMsgNo = 1
        ReturnValue.ReturnMsg = "查詢成功"
        ReturnValue.ReturnDs = DBViewDS
        Return ReturnValue
    End Function


    ''' <summary>
    '''  差異報表補跑與取消功能 SDK服務查詢
    ''' 查詢
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetView_View_StorePoints_CompanyLiaiseList_CPfile(ByVal GameFactoryName As String) As DataSet
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線) 
        Dim com As New SqlCommand()
        com.Connection = conn
        If GameFactoryName = "" Then
            com.CommandText = "Select GameFactoryId,GameFactoryName,GameName,GameNameEn,GameServiceId,GameServiceTypeSn,GameServiceTypeDesc from  PointsBilling_DB.dbo.View_StorePoints_CompanyLiaiseList_CPfile where 1=1  and SDKDifferencesReportStatus = 1    order by GameFactoryName  asc"
        Else
            com.CommandText = "Select GameFactoryId,GameFactoryName,GameName,GameNameEn,GameServiceId,GameServiceTypeSn,GameServiceTypeDesc from  PointsBilling_DB.dbo.View_StorePoints_CompanyLiaiseList_CPfile where 1=1   and  GameFactoryName=@GameFactoryName and SDKDifferencesReportStatus = 1  order by GameFactoryName asc"
            com.Parameters.Add("@GameFactoryName", SqlDbType.NVarChar, 30).Value = GameFactoryName
        End If

        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet()

        com.CommandTimeout = 100
        Try
            da.Fill(ds)

        Catch ex As Exception
            Throw New Exception("GetView_View_StorePoints_CompanyLiaiseList_CPfile|" & ex.Message)
        Finally
            conn.Close()
        End Try

        Return ds
    End Function

    <WebMethod()> _
    Public Function GetStorePoints_Currency() As DataSet
        ConnStr(12)

        ' 建立命令(預存程序名,使用連線) 
        Dim cmd As New SqlClient.SqlCommand
        cmd.CommandType = CommandType.Text
        cmd.CommandTimeout = 300
        cmd.Connection = conn
        cmd.CommandText = "SELECT Sn, CrrnyShotName, CrrnyName FROM [PointsBilling_DB].dbo.StorePoints_Currency ORDER BY Sn"

        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("Get_StorePoints_Currency資料失敗：" & ex.Message)
        End Try
        Return DSssid
    End Function

    '多語系設定
    <WebMethod()> _
    Public Function SPS_StorePoints_Global_FactoryService_MERGE(ByVal FactoryServiceId As String, ByVal FactoryServiceName As String, ByVal CurrencySn As Integer, ByVal Status As Integer, ByVal CreateUser As String) As ReturnValue
        ConnStr(12)
        ' 建立命令(預存程序名,使用連線) 
        Dim cmd As New SqlClient.SqlCommand("PointsBilling_DB.dbo.SPS_StorePoints_Global_FactoryService_MERGE", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandTimeout = 300

        cmd.Parameters.Add("@FactoryServiceId", System.Data.SqlDbType.NVarChar, 20).Value = FactoryServiceId
        cmd.Parameters.Add("@FactoryServiceName", System.Data.SqlDbType.NVarChar, 100).Value = FactoryServiceName
        cmd.Parameters.Add("@CurrencySn", System.Data.SqlDbType.Int).Value = CurrencySn
        cmd.Parameters.Add("@Status", System.Data.SqlDbType.Int).Value = Status
        cmd.Parameters.Add("@CreateUser", System.Data.SqlDbType.NVarChar, 30).Value = CreateUser

        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 50)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnValue
        Dim ADP As New SqlClient.SqlDataAdapter(cmd)
        Dim DS As New DataSet

        Dim Returnds As New ReturnValue
        Try
            ADP.Fill(DS)
        Catch ex As Exception
            Throw New Exception("SPS_StorePoints_Global_FactoryService_MERGE失敗：" & ex.Message)
        Finally
            conn.Close()
        End Try
        Returnds.ReturnMsgNo = ReturnMsgNo.Value
        Returnds.ReturnMsg = ReturnMsg.Value
        Return Returnds
    End Function

    '(明細)多語系查詢
    <WebMethod()> _
    Public Function VIEW_StorePoints_Global_FactoryService(ByVal FactoryServiceId As String) As DataSet
        Me.ConnStr(12)
        Dim command As New SqlCommand()
        command.Parameters.Clear()
        command.Connection = Me.conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300

        command.CommandText = " select * from [PointsBilling_DB].[dbo].[VIEW_StorePoints_Global_FactoryService] where 1=1 and Status = 1  and  FactoryServiceId = @FactoryServiceId "
        command.Parameters.Add("@FactoryServiceId", SqlDbType.NVarChar, 20).Value = FactoryServiceId

        Dim SDAssid As New SqlDataAdapter(command)
        Dim DSssid As New DataSet()
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("多語系資料查詢失敗：" & ex.Message)
        End Try
        Return DSssid
    End Function

    ''(明細)多語系查詢
    '<WebMethod()> _
    'Public Function View_BK_StorePoints_FactoryService_dept(ByVal FactoryServiceId As String) As DataSet
    '    Me.ConnStr(12)
    '    Dim command As New SqlCommand()
    '    command.Parameters.Clear()
    '    command.Connection = Me.conn
    '    command.CommandType = CommandType.Text
    '    command.CommandTimeout = 300

    '    command.CommandText = " Select distinct  Sn, FactoryId, FactoryServiceId,FactoryServiceName_All, FactoryServiceName, FactoryServiceType, ProductNo, Price , ServiceFee, Status, ItemOrder, MinimumPrice, RiskManagement  from [PointsBilling_DB].[dbo].[View_BK_StorePoints_FactoryService_dept] where 1=1 and Status = 1  and  FactoryServiceId = @FactoryServiceId "
    '    command.Parameters.Add("@FactoryServiceId", SqlDbType.NVarChar, 20).Value = FactoryServiceId

    '    Dim SDAssid As New SqlDataAdapter(command)
    '    Dim DSssid As New DataSet()
    '    Try
    '        SDAssid.Fill(DSssid)
    '    Catch ex As Exception
    '        Throw New Exception("View_BK_StorePoints_FactoryService_dept資料查詢失敗：" & ex.Message)
    '    End Try
    '    Return DSssid
    'End Function

#Region "Begintrans"
    Private TransManager As SqlClient.SqlTransaction
    Private BeinTransCommad As Boolean '設定這條連線要不要做Begintrans
    Public CommandTimeout As Integer = 30

    ''' <summary>
    ''' 初始交易確認機制,每個DBService 物件都有自己的begintrans
    ''' </summary>
    ''' <param name="oErrMsg">錯誤訊息</param>
    ''' <returns>執行結果,true就是成功</returns>
    ''' <remarks></remarks>
    Public Function BeginTrans(ByRef oErrMsg As String) As Boolean
        BeinTransCommad = True
        Try
            conn.Open()
            TransManager = conn.BeginTransaction()
        Catch ex As Exception
            BeginTrans = False
            oErrMsg = ex.Message
            Return BeginTrans
        End Try

        '祇要有一條 connection 開啟交易成功
        '就算已有開啟之交易
        '就需要 rollback


        BeginTrans = True
        oErrMsg = "ok"

        Return BeginTrans
    End Function
    ''' <summary>
    ''' 恢復更改的交易,begintrans之後,當有錯誤交易的時候執行這個
    ''' </summary>
    ''' <param name="oErrMsg">執行錯誤訊息</param>
    ''' <returns>執行成功回傳true</returns>
    ''' <remarks></remarks>
    Public Function RollbackTrans(ByRef oErrMsg As String) As Boolean
        Try
            TransManager.Rollback()

        Catch ex As Exception
            RollbackTrans = False
            oErrMsg = ex.Message
            Return RollbackTrans
        Finally
            conn.Close()

            BeinTransCommad = False
        End Try

        RollbackTrans = True
        oErrMsg = "ok"
        Return RollbackTrans
    End Function
    ''' <summary>
    ''' 確認交易,交易成功之後,執行這個確認交易
    ''' </summary>
    ''' <param name="oErrMsg">執行錯誤訊息</param>
    ''' <returns>執行結果,成功為true</returns>
    ''' <remarks></remarks>
    Public Function CommitTrans(ByRef oErrMsg As String) As Boolean
        Try
            TransManager.Commit()
        Catch ex As Exception
            CommitTrans = False
            oErrMsg = ex.Message
            Return CommitTrans
        Finally
            conn.Close()

            BeinTransCommad = False
        End Try

        CommitTrans = True
        oErrMsg = "ok"
        Return CommitTrans
    End Function
    Private Sub ConnOpen(ByRef Command As SqlClient.SqlCommand)
        '假如有藥beintrans不open
        Command.CommandTimeout = CommandTimeout
        If BeinTransCommad Then
            Command.Transaction = TransManager
        Else
            conn.Open()
        End If
    End Sub
    Private Sub ConnClose()
        '有要begintrans 時 不能關掉連線
        If BeinTransCommad = False Then
            conn.Close()
        End If
    End Sub
#End Region
End Class

Public Class PublicReturnRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnDS As DataSet
    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
        ReturnDS = New DataSet
    End Sub
End Class

Public Class InsertRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
End Class
Public Class UpdateRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Sub New()
        ReturnMsg = ""
    End Sub
End Class
Public Class ProductServiceUpdateRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
End Class
Public Class BillingOnlineSellReportRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnDS As DataSet
    Sub New()
        ReturnMsg = ""
    End Sub
End Class
Public Class ReturnRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public GSContractDateSn As Integer
End Class
Public Class TradeFailRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnDS As DataSet
    Sub New()
        ReturnMsg = ""
    End Sub
End Class
Public Class FactoryInDeptRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnDS As DataSet
    Sub New()
        ReturnMsg = ""
    End Sub
End Class
Public Class GetSPSStorePointsFactoryServiceValue
    'Insert,Update 
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
End Class
Public Class GetDataResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnDS As DataSet
    Sub New()
        ReturnMsg = ""
    End Sub
End Class

Public Class ReportRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnDS As DataSet
    Sub New()
        ReturnMsg = ""
    End Sub
End Class

Public Class ReturnResult
    Public ReturnMsgNo As Integer
    Public ReturnValue As Integer
    Public Sn As Integer
    Public ReturnMsg As String
    Public ReturnDS As DataSet
    Sub New()
        ReturnMsg = ""
    End Sub
End Class

Public Class ReturnValue
    'Insert,Update
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnDs As DataSet
    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
        ReturnDs = New DataSet
    End Sub
End Class