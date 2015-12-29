Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

' 若要允許使用 ASP.NET AJAX 從指令碼呼叫此 Web 服務，請取消註解下一行。
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class BLBACK
    Inherits System.Web.Services.WebService
    Dim myDbfnc As New MyCardConn.Dbfuction
    Dim conn As New SqlClient.SqlConnection()
    Dim conn1 As New SqlClient.SqlConnection()
    Dim MyCardDbService As New MyCardDbServices.DbWebService
    Dim ErrorLog As New wsError.wsError
    Dim Intreturnno As Integer
    ';
    Public Sub ConnStr(ByVal ConnNum As Integer)
        '正式環境用連19 , 測試用連100 
        Select Case My.Settings.DebugMode
            Case "1"
                conn = myDbfnc.open100Db()
                conn1 = myDbfnc.open100Db()
            Case "2"
                conn = myDbfnc.openTestDb()
                conn1 = myDbfnc.openTestDb()
            Case "3"
                Select Case ConnNum
                    Case 11
                        conn = myDbfnc.open11Db
                        conn1 = myDbfnc.open11Db
                    Case 12
                        conn = myDbfnc.open12Db
                        conn1 = myDbfnc.open12Db
                    Case 13
                        conn = myDbfnc.open13Db
                        conn1 = myDbfnc.open13Db
                    Case 14
                        conn = myDbfnc.open14Db
                        conn1 = myDbfnc.open14Db
                    Case 15
                        conn = myDbfnc.open15Db
                        conn1 = myDbfnc.open15Db
                    Case 16
                        conn = myDbfnc.open16Db
                        conn1 = myDbfnc.open16Db
                    Case 17
                        conn = myDbfnc.open17Db
                        conn1 = myDbfnc.open17Db
                    Case 19
                        conn = myDbfnc.open19Db
                        conn1 = myDbfnc.open19Db
                    Case 20
                        conn = myDbfnc.open20Db
                        conn1 = myDbfnc.open20Db
                    Case 21
                        conn = myDbfnc.open21Db
                        conn1 = myDbfnc.open21Db
                End Select
            Case Else
                conn = myDbfnc.open100Db()
                conn1 = myDbfnc.open100Db()
        End Select
    End Sub
    ''' <summary>
    ''' Billing後台/CP廠商引進單位管理。
    ''' 二、輸出參數BLBACKReturnResult，欄位值如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' 3、ReturnDS     DataSet。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function GetIntroductionDept(ByVal DeptName As String) As BLBACKReturnResult
        ConnStr(12)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        If DeptName = "" Then
            command.CommandText = "SELECT Sn , DeptName FROM [PointsBilling_DB].[dbo].[View_BK_StorePoints_IntroductionDept]"
        Else
            command.CommandText = "SELECT Sn , DeptName FROM [PointsBilling_DB].[dbo].[View_BK_StorePoints_IntroductionDept] where DeptName like '%" & Trim(DeptName) & "%'"
        End If
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "查詢廠商引進單位失敗：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = 1
        BLBACKReturnResult.ReturnMsg = "查詢成功！"
        BLBACKReturnResult.ReturnDS = DSssid
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    '''Billing後台/CP廠商引進單位管理/Insert。
    ''' 一、輸入參數。
    ''' 1、DeptName。    
    ''' 2、UserID。
    ''' 二、輸出參數BLBACKReturnResult，定義如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_IntroductionDept_Insert(ByVal DeptName As String, ByVal UserID As String) As BLBACKReturnResult
        ConnStr(12)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "PointsBilling_DB.dbo.SPS_StorePoints_IntroductionDept_Insert"
        command.Parameters.Add("@DeptName", SqlDbType.NVarChar, 30).Value = DeptName
        command.Parameters.Add("@UserId", SqlDbType.VarChar, 30).Value = UserID
        command.CommandTimeout = 300
        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output
        Try
            conn.Open()
            command.ExecuteNonQuery()
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "InsertCP廠商引進單位管理時：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = ReturnMsgNo.Value
        BLBACKReturnResult.ReturnMsg = ReturnMsg.Value
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    '''Billing後台/CP廠商引進單位管理/Update。
    ''' 一、輸入參數。
    ''' 1、Sn
    ''' 2、DeptName。    
    ''' 3、UserID。
    ''' 二、輸出參數BLBACKReturnResult，定義如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_IntroductionDept_Update(ByVal Sn As Integer, ByVal DeptName As String, ByVal UserID As String) As BLBACKReturnResult
        ConnStr(12)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "PointsBilling_DB.dbo.SPS_StorePoints_IntroductionDept_Update"
        command.Parameters.Add("@Sn", SqlDbType.Int).Value = Sn
        command.Parameters.Add("@DeptName", SqlDbType.NVarChar, 30).Value = DeptName
        command.Parameters.Add("@UserId", SqlDbType.VarChar, 30).Value = UserID
        command.CommandTimeout = 300
        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output
        Try
            conn.Open()
            command.ExecuteNonQuery()
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "UpdateCP廠商引進單位管理時：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = ReturnMsgNo.Value
        BLBACKReturnResult.ReturnMsg = ReturnMsg.Value
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    ''' Billing後台/金流廠商額度。
    ''' 二、輸出參數BLBACKReturnResult，欄位值如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' 3、ReturnDS     DataSet。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function GetFactory(ByVal FactoryName As String) As BLBACKReturnResult
        ConnStr(12)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        If FactoryName = "" Then
            command.CommandText = "SELECT Sn, FactoryId, FactoryName, UpperFristBound,DayUpperMobileBound,MonthUpperMobileBound,DayUpperMailBound,MonthUpperMailBound,DayUpperCustIdBound,MonthUpperCustIdBound,TempCodeLimit,IPTradeLimit, Status FROM [PointsBilling_DB].[dbo].[View_BK_StorePoints_Factory] where Status=1"
        Else
            command.CommandText = "SELECT Sn, FactoryId, FactoryName, UpperFristBound,DayUpperMobileBound,MonthUpperMobileBound,DayUpperMailBound,MonthUpperMailBound,DayUpperCustIdBound,MonthUpperCustIdBound,TempCodeLimit,IPTradeLimit, Status FROM [PointsBilling_DB].[dbo].[View_BK_StorePoints_Factory] where Status=1 and FactoryName  like '%" & Trim(FactoryName) & "%'"
        End If
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "查詢金流廠商額度失敗：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = 1
        BLBACKReturnResult.ReturnMsg = "查詢成功！"
        BLBACKReturnResult.ReturnDS = DSssid
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    '''Billing後台/金流廠商額度/Update。
    ''' 一、輸入參數。
    ''' 1、Sn
    ''' 2、DeptName。    
    ''' 3、UserID。
    ''' 二、輸出參數BLBACKReturnResult，定義如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_Factory_UpperBound_Update(ByVal Sn As Integer, ByVal UpperFristBound As Integer, ByVal DayUpperMobileBound As Integer, _
                                                            ByVal MonthUpperMobileBound As Integer, ByVal DayUpperMailBound As Integer, ByVal MonthUpperMailBound As Integer, _
                                                            ByVal DayUpperCustIdBound As Integer, ByVal MonthUpperCustIdBound As Integer, ByVal TempCodeLimit As Integer, _
                                                            ByVal IPTradeLimit As Integer, ByVal UserID As String) As BLBACKReturnResult
        ConnStr(12)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "PointsBilling_DB.dbo.SPS_StorePoints_Factory_UpperBound_Update"
        command.Parameters.Add("@Sn", SqlDbType.Int).Value = Sn
        command.Parameters.Add("@UpperFristBound", SqlDbType.Int).Value = UpperFristBound
        command.Parameters.Add("@DayUpperMobileBound", SqlDbType.Int).Value = DayUpperMobileBound
        command.Parameters.Add("@MonthUpperMobileBound", SqlDbType.Int).Value = MonthUpperMobileBound
        command.Parameters.Add("@DayUpperMailBound", SqlDbType.Int).Value = DayUpperMailBound
        command.Parameters.Add("@MonthUpperMailBound", SqlDbType.Int).Value = MonthUpperMailBound
        command.Parameters.Add("@DayUpperCustIdBound", SqlDbType.Int).Value = DayUpperCustIdBound
        command.Parameters.Add("@MonthUpperCustIdBound", SqlDbType.Int).Value = MonthUpperCustIdBound
        command.Parameters.Add("@TempCodeLimit", SqlDbType.TinyInt).Value = TempCodeLimit
        command.Parameters.Add("@IPTradeLimit", SqlDbType.TinyInt).Value = IPTradeLimit
        command.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserID
        command.CommandTimeout = 300
        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 15)
        ReturnMsg.Direction = ParameterDirection.Output
        Try
            conn.Open()
            command.ExecuteNonQuery()
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "Update金流廠商額度時：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = ReturnMsgNo.Value
        BLBACKReturnResult.ReturnMsg = ReturnMsg.Value
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    ''' Billing後台/信用卡黑名單。
    ''' 一、輸入參數
    ''' 1、CreditCardNum
    ''' 2、Status
    ''' 二、輸出參數BLBACKReturnResult，欄位值如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' 3、ReturnDS     DataSet。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function GetBlackCreditCardList(ByVal CreditCardNum As String, ByVal Status As String) As BLBACKReturnResult
        ConnStr(12)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        If CreditCardNum = "" Then
            If Status = "" Then
                command.CommandText = "SELECT distinct Sn,CreditCardNum,Status FROM [PointsBilling_DB].[dbo].[View_BK_StorePoints_BlackCreditCardList]"
            Else
                command.CommandText = "SELECT distinct Sn,CreditCardNum,Status FROM [PointsBilling_DB].[dbo].[View_BK_StorePoints_BlackCreditCardList] where Status=@Status"
                command.Parameters.Add("@Status", SqlDbType.TinyInt).Value = CInt(Status)
            End If
        Else
            If Status = "" Then
                command.CommandText = "SELECT distinct Sn,CreditCardNum,Status FROM [PointsBilling_DB].[dbo].[View_BK_StorePoints_BlackCreditCardList] where CreditCardNum like '%" & Trim(CreditCardNum) & "%'"
            Else
                command.CommandText = "SELECT distinct Sn,CreditCardNum,Status FROM [PointsBilling_DB].[dbo].[View_BK_StorePoints_BlackCreditCardList] where Status=@Status and CreditCardNum like '%" & Trim(CreditCardNum) & "%'"
                command.Parameters.Add("@Status", SqlDbType.TinyInt).Value = CInt(Status)
            End If
        End If
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "查詢信用卡黑名單失敗：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = 1
        BLBACKReturnResult.ReturnMsg = "查詢成功！"
        BLBACKReturnResult.ReturnDS = DSssid
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    ''' Billing後台/信用卡黑名單/信用卡黑名單歷史資料。
    ''' 一、輸入參數
    ''' 1、Sn
    ''' 二、輸出參數BLBACKReturnResult，欄位值如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' 3、ReturnDS     DataSet。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function GetBlackCreditCardListHistory(ByVal Sn As String) As BLBACKReturnResult
        ConnStr(12)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        command.CommandText = "SELECT Sn,CreditCardNum,LogStatus,ProcDesc,CreateUser,CreateDate FROM [PointsBilling_DB].[dbo].[View_BK_StorePoints_BlackCreditCardList] where Sn=@Sn order by CreateDate Desc"
        command.Parameters.Add("@Sn", SqlDbType.Int).Value = CInt(Trim(Sn))
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "查詢信用卡黑名單歷史資料失敗：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = 1
        BLBACKReturnResult.ReturnMsg = "查詢成功！"
        BLBACKReturnResult.ReturnDS = DSssid
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    ''' Billing後台/信用卡黑名單/信用卡黑名單原因。
    ''' 一、輸入參數
    ''' 1、Sn
    ''' 二、輸出參數BLBACKReturnResult，欄位值如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' 3、ReturnDS     DataSet。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function GetProcDesc(ByVal Sn As String) As BLBACKReturnResult
        ConnStr(12)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        command.CommandText = "select top 1 ProcDesc from  [PointsBilling_DB].[dbo].[View_BK_StorePoints_BlackCreditCardList] where Sn=@Sn order by createdate desc"
        command.Parameters.Add("@Sn", SqlDbType.Int).Value = Sn
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "查詢信用卡黑名單原因失敗：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = 1
        BLBACKReturnResult.ReturnMsg = "查詢成功！"
        BLBACKReturnResult.ReturnDS = DSssid
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    '''Billing後台/信用卡黑名單/Insert。
    ''' 一、輸入參數。
    ''' 1、CreditCardNum。
    ''' 2、ProcDesc
    ''' 3、UserID
    ''' 二、輸出參數BLBACKReturnResult，定義如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_BlackCreditCardList_Insert(ByVal CreditCardNum As String, ByVal ProcDesc As String, ByVal UserID As String) As BLBACKReturnResult
        ConnStr(12)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "PointsBilling_DB.dbo.SPS_StorePoints_BlackCreditCardList_Insert"
        command.Parameters.Add("@CreditCardNum", SqlDbType.VarChar, 6).Value = CreditCardNum
        command.Parameters.Add("@ProcDesc", SqlDbType.NVarChar, 30).Value = ProcDesc
        command.Parameters.Add("@CreateUser ", SqlDbType.VarChar, 30).Value = UserID
        command.CommandTimeout = 300
        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Try
            conn.Open()
            command.ExecuteNonQuery()
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "Insert信用卡黑名單時：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = ReturnMsgNo.Value
        BLBACKReturnResult.ReturnMsg = ReturnMsg.Value
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    '''Billing後台/信用卡黑名單/Update。
    ''' 一、輸入參數。
    ''' 1、SN。    
    ''' 2、CreditCardNum。
    ''' 3、ProcDesc
    ''' 4、UserID
    ''' 二、輸出參數BLBACKReturnResult，定義如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_BlackCreditCardList_Update(ByVal SN As Integer, ByVal ProcDesc As String, ByVal Status As Integer, ByVal UserID As String) As BLBACKReturnResult
        ConnStr(12)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "PointsBilling_DB.dbo.SPS_StorePoints_BlackCreditCardList_Update"
        command.Parameters.Add("@Sn", SqlDbType.Int).Value = SN
        command.Parameters.Add("@ProcDesc", SqlDbType.NVarChar, 30).Value = ProcDesc
        command.Parameters.Add("@Status", SqlDbType.TinyInt).Value = Status
        command.Parameters.Add("@UserStamp ", SqlDbType.VarChar, 30).Value = UserID
        command.CommandTimeout = 300
        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Try
            conn.Open()
            command.ExecuteNonQuery()
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "Update信用卡黑名單時：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = ReturnMsgNo.Value
        BLBACKReturnResult.ReturnMsg = ReturnMsg.Value
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    ''' Billing後台/儲值交易數量日報表查詢/遊戲廠商。
    ''' 一、輸入參數
    ''' 1、SelectType
    ''' 2、TaxIDNo
    ''' 二、輸出參數BLBACKReturnResult，欄位值如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' 3、ReturnDS     DataSet。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_FactoryForMyCardBillingTradeDayReport(ByVal SelectType As String, ByVal TaxIDNo As String) As BLBACKReturnResult
        ConnStr(14)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        If SelectType = "0" Then
            command.CommandText = "Select * from MyCard_BackUp.dbo.View_PointsBilling_GameService  order by GameName asc"
        Else
            command.CommandText = "Select * from MyCard_BackUp.dbo.View_PointsBilling_GameService Where TaxIDNo=@TaxIDNo order by GameName asc"
            command.Parameters.Add("@TaxIDNo", SqlDbType.VarChar, 15).Value = TaxIDNo
        End If
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "查詢遊戲廠商失敗：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = 1
        BLBACKReturnResult.ReturnMsg = "查詢成功！"
        BLBACKReturnResult.ReturnDS = DSssid
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    ''' Billing後台/儲值交易數量日報表查詢/產品名稱。
    ''' 一、輸入參數
    ''' 1、SelectType
    ''' 2、TaxIDNo
    ''' 二、輸出參數BLBACKReturnResult，欄位值如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' 3、ReturnDS     DataSet。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_ProductsForMyCardBillingTradeDayReport(ByVal GameServiceId As String) As BLBACKReturnResult
        ConnStr(14)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        command.CommandText = "Select * from MyCard_BackUp.dbo.View_PointsBilling_GameProduct Where GameServiceId=@GameServiceId order by GameProductName asc"
        command.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 20).Value = GameServiceId
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "查詢產品名稱失敗：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = 1
        BLBACKReturnResult.ReturnMsg = "查詢成功！"
        BLBACKReturnResult.ReturnDS = DSssid
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    ''' Billing後台/付費廠商流程維護/廠商。
    ''' 二、輸出參數BLBACKReturnResult，欄位值如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' 3、ReturnDS     DataSet。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_FactoryForFlowManage() As BLBACKReturnResult
        ConnStr(12)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        command.CommandText = "SELECT  FactoryId, FactoryName FROM PointsBilling_DB.dbo.View_BK_StorePoints_Factory order by Sn asc"
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "查詢廠商失敗：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = 1
        BLBACKReturnResult.ReturnMsg = "查詢成功！"
        BLBACKReturnResult.ReturnDS = DSssid
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    ''' Billing後台/付費廠商流程維護/查詢。
    ''' 一、輸入參數
    ''' 1、TradeType
    ''' 2、TradeSeq
    ''' 二、輸出參數BLBACKReturnResult，欄位值如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' 3、ReturnDS     DataSet。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function GetFlowManageDetail(ByVal FactoryId As String) As BLBACKReturnResult
        ConnStr(14)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = 300
        command.CommandText = "PointsBilling_DB.dbo.SPS_StorePoints_TradeProceduresDescr_Query"
        command.Parameters.Add("@BankFactoryId", SqlDbType.VarChar, 20).Value = FactoryId
        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "查詢SPS_StorePoints_TradeProceduresDescr_Query失敗：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = ReturnMsgNo.Value
        BLBACKReturnResult.ReturnMsg = ReturnMsg.Value
        BLBACKReturnResult.ReturnDS = DSssid
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    '''Billing後台/付費廠商流程維護/Insert。
    ''' 一、輸入參數。
    ''' 1、BankFactoryId。
    ''' 2、TPId
    ''' 3、TPDescr
    ''' 4、UserID
    ''' 二、輸出參數BLBACKReturnResult，定義如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_TradeProceduresDescr_Insert(ByVal BankFactoryId As String, ByVal TPId As String, ByVal TPDescr As String, ByVal UserID As String) As BLBACKReturnResult
        ConnStr(12)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "PointsBilling_DB.dbo.SPS_StorePoints_TradeProceduresDescr_Insert"
        command.Parameters.Add("@BankFactoryId", SqlDbType.VarChar, 20).Value = BankFactoryId
        command.Parameters.Add("@TPId", SqlDbType.Int).Value = TPId
        command.Parameters.Add("@TPDescr", SqlDbType.NVarChar, 50).Value = TPDescr
        command.Parameters.Add("@CreateUser", SqlDbType.VarChar, 30).Value = UserID
        command.CommandTimeout = 300
        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Try
            conn.Open()
            command.ExecuteNonQuery()
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "Insert付費廠商流程維護時：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = ReturnMsgNo.Value
        BLBACKReturnResult.ReturnMsg = ReturnMsg.Value
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    '''Billing後台/付費廠商流程維護/Update。
    ''' 一、輸入參數。
    ''' 1、Sn。
    ''' 2、TPId
    ''' 3、TPDescr
    ''' 4、UserID
    ''' 二、輸出參數BLBACKReturnResult，定義如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function SPS_StorePoints_TradeProceduresDescr_Update(ByVal Sn As Integer, ByVal BankFactoryId As String, ByVal TPId As String, ByVal TPDescr As String, ByVal UserID As String) As BLBACKReturnResult
        ConnStr(12)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "PointsBilling_DB.dbo.SPS_StorePoints_TradeProceduresDescr_Update"
        command.Parameters.Add("@Sn", SqlDbType.Int).Value = Sn
        command.Parameters.Add("@TPId", SqlDbType.Int).Value = TPId
        command.Parameters.Add("@TPDescr", SqlDbType.NVarChar, 50).Value = TPDescr
        command.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserID
        command.CommandTimeout = 300
        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Try
            conn.Open()
            command.ExecuteNonQuery()
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "Update付費廠商流程維護時：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = ReturnMsgNo.Value
        BLBACKReturnResult.ReturnMsg = ReturnMsg.Value
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    ''' Billing後台/連續扣款查詢報表/付費機制。
    ''' 二、輸出參數BLBACKReturnResult，欄位值如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' 3、ReturnDS     DataSet。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function ViewBKStorePointsFactory() As BLBACKReturnResult
        ConnStr(14)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        command.CommandText = "Select * from PointsBilling_DB.dbo.View_BK_StorePoints_Factory "
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "查詢付費機制失敗：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = 1
        BLBACKReturnResult.ReturnMsg = "查詢成功！"
        BLBACKReturnResult.ReturnDS = DSssid
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    ''' Billing後台/連續扣款查詢報表/服務品項。
    ''' 二、輸出參數BLBACKReturnResult，欄位值如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' 3、ReturnDS     DataSet。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function ViewBKStorePointsGameProducts() As BLBACKReturnResult
        ConnStr(14)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        command.CommandText = "Select * from PointsBilling_DB.dbo.View_BK_StorePoints_GameProducts "
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "查詢服務品項失敗：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = 1
        BLBACKReturnResult.ReturnMsg = "查詢成功！"
        BLBACKReturnResult.ReturnDS = DSssid
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    ''' Billing後台/連續扣款查詢報表/服務。
    ''' 二、輸出參數BLBACKReturnResult，欄位值如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' 3、ReturnDS     DataSet。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function ViewBKStorePointsGameService() As BLBACKReturnResult
        ConnStr(14)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.Text
        command.CommandTimeout = 300
        command.CommandText = "Select * from PointsBilling_DB.dbo.View_BK_StorePoints_GameService "
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "查詢服務失敗：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = 1
        BLBACKReturnResult.ReturnMsg = "查詢成功！"
        BLBACKReturnResult.ReturnDS = DSssid
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    ''' Billing後台/付費廠商流程維護/查詢彙總。
    ''' 一、輸入參數
    ''' 1、TradeType
    ''' 2、TradeSeq
    ''' 二、輸出參數BLBACKReturnResult，欄位值如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' 3、ReturnDS     DataSet。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function StorePointsSeriaPayListTotalReport(ByVal PayDateStart As String, ByVal PayDateEnd As String, ByVal CancelDateStart As String, ByVal CancelDateEnd As String, ByVal FactorySn As Integer, ByVal PayStatus As Integer, ByVal GameProductId As String, ByVal GameServiceId As String, ByVal ActiveStatus As Integer, ByVal SerialId As String) As BLBACKReturnResult
        ConnStr(14)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "PointsBilling_DB.dbo.SPS_StorePoints_SeriaPayList_TotalReport"

        If PayDateStart.Trim = "" Then
            command.Parameters.AddWithValue("@PayDateStart", DBNull.Value)
        Else
            command.Parameters.Add("@PayDateStart", SqlDbType.DateTime).Value = PayDateStart
        End If

        If PayDateEnd.Trim = "" Then
            command.Parameters.AddWithValue("@PayDateEnd", DBNull.Value)
        Else
            command.Parameters.Add("@PayDateEnd", SqlDbType.DateTime).Value = PayDateEnd
        End If

        If CancelDateStart = "" Then
            command.Parameters.AddWithValue("@CancelDateStart", DBNull.Value)
        Else
            command.Parameters.AddWithValue("@CancelDateStart", DateTime.Parse(CancelDateStart))
        End If

        If CancelDateEnd = "" Then
            command.Parameters.AddWithValue("@CancelDateEnd", DBNull.Value)
        Else
            command.Parameters.AddWithValue("@CancelDateEnd", DateTime.Parse(CancelDateEnd))
        End If

        command.Parameters.Add("@FactorySn", SqlDbType.Int).Value = FactorySn
        command.Parameters.Add("@PayStatus", SqlDbType.SmallInt).Value = PayStatus
        command.Parameters.Add("@GameProductSn", SqlDbType.Int).Value = IIf(GameProductId = "", 0, GameProductId)
        command.Parameters.Add("@GameServiceId", SqlDbType.VarChar, 30).Value = GameServiceId
        command.Parameters.Add("@ActiveStatus", SqlDbType.SmallInt).Value = ActiveStatus
        command.Parameters.Add("@SerialId", SqlDbType.VarChar, 16).Value = SerialId
        command.CommandTimeout = 900
        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "查詢StorePointsSeriaPayListTotalReport失敗：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = ReturnMsgNo.Value
        BLBACKReturnResult.ReturnMsg = ReturnMsg.Value
        BLBACKReturnResult.ReturnDS = DSssid
        Return BLBACKReturnResult
    End Function
    ''' <summary>
    ''' Billing後台/付費廠商流程維護/查詢明細。
    ''' 一、輸入參數
    ''' 1、TradeType
    ''' 2、TradeSeq
    ''' 二、輸出參數BLBACKReturnResult，欄位值如下：
    ''' 1、ReturnMsgNo	錯誤訊息代碼。
    ''' 2、ReturnMsg	錯誤訊息。
    ''' 3、ReturnDS     DataSet。
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function StorePointsSeriaPayListReport(ByVal dtDate_Start As String, ByVal dtDate_End As String, ByVal dCancelDate_Start As String, ByVal dCancelDate_END As String, ByVal factory As String, ByVal CustId As String, ByVal PayStatus As String, ByVal GameProductsSn As String, ByVal FactoryServiceName As String, ByVal ActiveStatus As Integer, ByVal Status As Integer, ByVal TradeSeq As String, ByVal QueryType As Integer, ByVal SerialId As String) As BLBACKReturnResult
        ConnStr(14)
        Dim BLBACKReturnResult As New BLBACKReturnResult
        Dim command As New SqlClient.SqlCommand
        command.Parameters.Clear()
        command.Connection = conn
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = 900
        command.CommandText = "PointsBilling_DB.dbo.SPS_StorePoints_SeriaPayListReport"

        If dtDate_Start = "" Then
            command.Parameters.AddWithValue("@dtDate_Start", DBNull.Value)
        Else
            command.Parameters.Add("@dtDate_Start", SqlDbType.VarChar, 50).Value = dtDate_Start
        End If

        If dtDate_End = "" Then
            command.Parameters.AddWithValue("@dtDate_End", DBNull.Value)
        Else
            command.Parameters.Add("@dtDate_End", SqlDbType.VarChar, 50).Value = dtDate_End
        End If

        If dCancelDate_Start = "" Then
            command.Parameters.AddWithValue("@dCancelDate_Start", DBNull.Value)
        Else
            command.Parameters.AddWithValue("@dCancelDate_Start", DateTime.Parse(dCancelDate_Start))
        End If

        If dCancelDate_END = "" Then
            command.Parameters.AddWithValue("@dCancelDate_END", DBNull.Value)
        Else
            command.Parameters.AddWithValue("@dCancelDate_END", DateTime.Parse(dCancelDate_END))
        End If
        command.Parameters.Add("@factory", SqlDbType.VarChar, 20).Value = factory
        command.Parameters.Add("@CustId", SqlDbType.VarChar, 100).Value = CustId
        command.Parameters.Add("@PayStatus", SqlDbType.VarChar, 1).Value = PayStatus
        command.Parameters.Add("@GameProductsSn", SqlDbType.VarChar, 10).Value = GameProductsSn
        command.Parameters.Add("@GameCustId", SqlDbType.VarChar, 100).Value = FactoryServiceName.Trim
        command.Parameters.Add("@ActiveStatus", SqlDbType.VarChar, 1).Value = ActiveStatus
        command.Parameters.Add("@Status", SqlDbType.VarChar, 1).Value = Status
        command.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 16).Value = TradeSeq
        command.Parameters.Add("@QueryType", SqlDbType.Int).Value = QueryType
        command.Parameters.Add("@SerialId", SqlDbType.VarChar, 16).Value = SerialId

        command.CommandTimeout = 300
        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnValue", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(command)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            BLBACKReturnResult.ReturnMsgNo = -100
            BLBACKReturnResult.ReturnMsg = "查詢StorePointsSeriaPayListReport失敗：" & ex.Message
            Return BLBACKReturnResult
        Finally
            conn.Close()
        End Try
        BLBACKReturnResult.ReturnMsgNo = ReturnMsgNo.Value
        BLBACKReturnResult.ReturnMsg = "查詢成功"
        BLBACKReturnResult.ReturnDS = DSssid
        Return BLBACKReturnResult
    End Function

End Class
Public Class BLBACKReturnResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnDS As DataSet
    Sub New()
        ReturnMsg = ""
    End Sub
End Class