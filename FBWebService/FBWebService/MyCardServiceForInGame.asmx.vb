Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class MyCardServiceForInGame
    Inherits System.Web.Services.WebService
    Dim myDbfnc As New MyCardConn.Dbfuction
    Dim Con As New SqlClient.SqlConnection()
    Dim MyCardDbwebservice As New dbwebservice.DbWebService
    Public Sub ConnStr(ByVal ConnNum As Integer)
        Select Case My.Settings.IsTest
            Case UCase("True")
                Con = myDbfnc.open100Db
            Case UCase("Test")
                Con = myDbfnc.openTestDb
            Case UCase("False")
                Select Case ConnNum
                    'Case 11
                    '    Con = myDbfnc.open11Db
                    'Case 12
                    '    Con = myDbfnc.open12Db
                    'Case 13
                    '    Con = myDbfnc.open13Db
                    'Case 14
                    '    Con = myDbfnc.open14Db
                    Case 15
                        Con = myDbfnc.open15Db
                        'Case 16
                        '    Con = myDbfnc.open16Db
                        'Case 17
                        '    Con = myDbfnc.open17Db
                    Case 19
                        Con = myDbfnc.open19Db
                    Case 20
                        Con = myDbfnc.open20Db
                End Select
            Case Else
                Con = myDbfnc.open100Db
        End Select
 
    End Sub
    ''' <summary>
    '''   MyCardInGame
    ''' </summary>
    ''' <param name="GameFacID">遊戲廠商代號,格式String</param>
    ''' <param name="MyCardCardId">MyCard卡號,格式String</param>
    ''' <param name="GameCustId">遊戲帳號,格式String</param>
    ''' <param name="CreateIP">新增者IP,格式String</param>
    ''' <returns>ReturnNo 回傳代碼 ,ReturnMsg 回傳訊息,MyCardPoint 回傳廠商交易序號,ErrorCode 回傳ErrorCode,LogSn 回傳LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function MyCardCPSSInGameSaveTradeListInsert(ByVal GameFacID As String, ByVal MyCardCardId As String, ByVal GameCustId As String, ByVal CreateIP As String, ByVal TradeType As Integer) As ReturnValueRE
        ConnStr(15)
        ' 建立命令(預存程序名,使用連線) 
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_InGameSaveTradeList_Insert_A01", Con)
        ' 設置命令型態為預存程序 
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        ' 增加參數欄位

        Cmd.Parameters.Add("@GameFacID", SqlDbType.VarChar, 30).Value = GameFacID
        Cmd.Parameters.Add("@MyCardCardId", SqlDbType.VarChar, 20).Value = MyCardCardId
        Cmd.Parameters.Add("@GameCustId", SqlDbType.VarChar, 100).Value = GameCustId
        Cmd.Parameters.Add("@CreateIP", SqlDbType.VarChar, 30).Value = CreateIP
        Cmd.Parameters.Add("@TradeType ", SqlDbType.TinyInt).Value = TradeType

        Dim TradeSeq As SqlClient.SqlParameter = Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20)
        TradeSeq.Direction = ParameterDirection.Output
        Dim GamePrice As SqlClient.SqlParameter = Cmd.Parameters.Add("@GamePrice", SqlDbType.Int)
        GamePrice.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output


        Dim ReturnValue As New ReturnValueRE
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = ReturnMsg.Value
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnTradeSeq = IIf(IsDBNull(TradeSeq.Value), "", TradeSeq.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.ReturnGamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)

        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|MyCardCPSSInGameSaveTradeListInsert|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2001"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.ReturnGamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)

        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function
    ''' <summary>
    '''   MyCardInGame
    ''' </summary>
    ''' <param name="MyCardID">MyCard卡號,格式String</param>
    ''' <param name="MyCardPW">MyCard密碼,格式String</param>
    ''' <param name="GameFacID">遊戲廠商代號,格式String</param>
    ''' <param name="GameUser">遊戲帳號,格式String</param>
    ''' <param name="Game_No">交易序號,格式String</param>
    ''' <param name="GameCard_ID">廠商本身儲值的卡,格式String</param>
    ''' <returns>ReturnNo 回傳代碼 ,ReturnMsg 回傳訊息,MyCardPoint 點數,MyCardProjetNo活動代號,MyCardType 卡別</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function MyCardInGame(ByVal MyCardID As String, ByVal MyCardPW As String, ByVal GameFacID As String, ByVal GameUser As String, ByVal Game_No As String, ByVal GameCard_ID As String) As Mycardservice.ReturnSaveResult
        Dim InGameWS As New Mycardservice.MyCardService
        Dim InGameWSReValue As New Mycardservice.ReturnSaveResult
        Try
            InGameWSReValue = InGameWS.MyCardRender(MyCardID, MyCardPW, GameFacID, GameUser, Game_No, GameCard_ID)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|MyCardRender|" & ex.Message, "127.0.0.3")
            InGameWSReValue.ReturnNo = "-999"
            InGameWSReValue.ReturnMsg = "系統發生錯誤" & ex.Message
            Return InGameWSReValue
        End Try
        Return InGameWSReValue
    End Function
    ''' <summary>
    '''   MyCardInGame
    ''' </summary>
    ''' <param name="TradeSeq">交易序號,格式String</param>
    ''' <param name="CardPoint">點數,格式Integer</param>
    ''' <param name="oProjNo">活動代號,格式String</param>
    ''' <param name="CardKind">卡通路別(2實體,1虛擬),格式Integer</param>
    ''' <returns>ReturnNo 回傳代碼 ,ReturnMsg 回傳訊息,ErrorCode 回傳ErrorCode,LogSn 回傳LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function MyCardCPSSInGameSaveGetInfo(ByVal TradeSeq As String, ByVal CardPoint As Integer, ByVal oProjNo As String, ByVal CardKind As Integer) As MyCardCPSSInGameSaveGetInfoRE
        ConnStr(15)
        ' 建立命令(預存程序名,使用連線)
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_InGameSave_GetInfo_A02", Con)
        ' 設置命令型態為預存程序 
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        ' 增加參數欄位

        Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20).Value = TradeSeq
        Cmd.Parameters.Add("@CardPoint", SqlDbType.Int).Value = CardPoint
        Cmd.Parameters.Add("@oProjNo", SqlDbType.VarChar, 20).Value = oProjNo
        Cmd.Parameters.Add("@CardKind", SqlDbType.Int).Value = CardKind
        Dim GamePrice As SqlClient.SqlParameter = Cmd.Parameters.Add("@GamePrice", SqlDbType.Int)
        GamePrice.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output

        Dim CrrnyShotName As SqlClient.SqlParameter = Cmd.Parameters.Add("@CrrnyShotName", SqlDbType.VarChar, 6)
        CrrnyShotName.Direction = ParameterDirection.Output

        Dim Discount As SqlClient.SqlParameter = Cmd.Parameters.Add("@Discount", SqlDbType.VarChar, 7)
        Discount.Direction = ParameterDirection.Output


        Dim ReturnValue As New MyCardCPSSInGameSaveGetInfoRE
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = ReturnMsg.Value
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.ReturnGamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)
            ReturnValue.ReturnCrrnyShotName = IIf(IsDBNull(CrrnyShotName.Value), "", CrrnyShotName.Value)
            ReturnValue.ReturnDiscount = IIf(IsDBNull(Discount.Value), 0, Discount.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|MyCardCPSSInGameSaveGetInfo|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2002"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.ReturnGamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)
            ReturnValue.ReturnCrrnyShotName = IIf(IsDBNull(CrrnyShotName.Value), "", CrrnyShotName.Value)
            ReturnValue.ReturnDiscount = IIf(IsDBNull(Discount.Value), 0, Discount.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function
    ''' <summary>
    '''   MyCardInGame
    ''' </summary>
    ''' <param name="TradeSeq">交易序號,格式String</param>
    ''' <param name="OTPCode">OTP交易,格式String</param>
    ''' <returns>InnerOTPCode OTP ReturnNo 回傳代碼 ,ReturnMsg 回傳訊息,ErrorCode 回傳ErrorCode,LogSn 回傳LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function MyCardCPSSMemberAddGetOTPCode(ByVal TradeSeq As String, ByVal OTPCode As String) As ReturnValueRE
        ConnStr(15)
        ' 建立命令(預存程序名,使用連線) 
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_MemberAdd_GetOTPCode_A01", Con)
        ' 設置命令型態為預存程序 
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        ' 增加參數欄位

        Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20).Value = TradeSeq
        Cmd.Parameters.Add("@OTPCode", SqlDbType.VarChar, 50).Value = OTPCode


        Dim InnerOTPCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@InnerOTPCode", SqlDbType.VarChar, 50)
        InnerOTPCode.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnValueRE
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = ReturnMsg.Value
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.ReturnInnerOTPCode = IIf(IsDBNull(InnerOTPCode.Value), "", InnerOTPCode.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|MyCardCPSSMemberAddGetOTPCode|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2003"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.ReturnInnerOTPCode = InnerOTPCode.Value
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function
    ''' <summary>
    '''   MyCardInGame
    ''' </summary>
    ''' <param name="GameCustId">遊戲帳號,格式String</param>
    ''' <param name="GamePoint">實際轉入的遊戲點數,格式Integer</param>
    ''' <param name="TradeSeq">交易序號,格式String</param>
    ''' <param name="GameNo">交易序號,格式String</param>
    ''' <returns>ReturnNo 回傳代碼 ,ReturnMsg 回傳訊息,ErrorCode 回傳ErrorCode,LogSn 回傳LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function MyCardCPSSInGameSaveGetGamePoint(ByVal GameCustId As String, ByVal GamePoint As Integer, ByVal TradeSeq As String, ByVal GameNo As String) As ReturnValueRE
        ConnStr(15)
        ' 建立命令(預存程序名,使用連線)
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_InGameSave_GetGamePoint_A01", Con)
        ' 設置命令型態為預存程序 
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        ' 增加參數欄位

        Cmd.Parameters.Add("@GameCustId", SqlDbType.VarChar, 100).Value = GameCustId
        Cmd.Parameters.Add("@GamePoint", SqlDbType.Int).Value = GamePoint
        Cmd.Parameters.Add("@GameNo", SqlDbType.VarChar, 50).Value = GameNo
        Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20).Value = TradeSeq

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnValueRE
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = ReturnMsg.Value
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|MyCardCPSSInGameSaveGetGamePoint|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2004"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function

    ''' <summary>
    ''' MyCardInGame
    ''' NewAPI
    ''' </summary>
    ''' <param name="GameCustId">遊戲帳號,格式String</param>
    ''' <param name="GamePoint">實際轉入的遊戲點數,格式Integer</param>
    ''' <param name="TradeSeq">交易序號,格式String</param>
    ''' <param name="GameNo">交易序號,格式String</param>
    ''' <returns>ReturnNo 回傳代碼 ,ReturnMsg 回傳訊息,ErrorCode 回傳ErrorCode,LogSn 回傳LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function MyCardCPSSInGameSaveGetGamePoint_NewAPI(ByVal GameCustId As String, ByVal GamePoint As Decimal, ByVal TradeSeq As String, ByVal GameNo As String) As ReturnValueRE
        ConnStr(15)
        ' 建立命令(預存程序名,使用連線)
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_InGameSave_GetGamePoint_A02", Con)
        ' 設置命令型態為預存程序 
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        ' 增加參數欄位

        Cmd.Parameters.Add("@GameCustId", SqlDbType.VarChar, 100).Value = GameCustId
        Cmd.Parameters.Add("@GamePoint", SqlDbType.Money).Value = GamePoint
        Cmd.Parameters.Add("@GameNo", SqlDbType.VarChar, 50).Value = GameNo
        Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20).Value = TradeSeq

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnValueRE
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = ReturnMsg.Value
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|MyCardCPSSInGameSaveGetGamePoint|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2004"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function

    ''' <summary>
    ''' MyCardInGame - BONUS
    ''' 如果有Bonus交易序號，就會更新
    ''' </summary>
    ''' <param name="GameCustId">遊戲帳號,格式String</param>
    ''' <param name="GamePoint">實際轉入的遊戲點數,格式Integer</param>
    ''' <param name="TradeSeq">交易序號,格式String</param>
    ''' <param name="GameNo">交易序號,格式String</param>
    ''' <returns>ReturnNo 回傳代碼 ,ReturnMsg 回傳訊息,ErrorCode 回傳ErrorCode,LogSn 回傳LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function MyCardCPSSInGameSaveGetGamePoint_NewAPI_Bonus(ByVal GameCustId As String, ByVal GamePoint As Decimal, ByVal TradeSeq As String, ByVal GameNo As String, ByVal GameNo_Bonus As String) As ReturnValueRE
        ConnStr(15)
        Dim ReturnValue As New ReturnValueRE
        Dim Times As Integer = IIf(GameNo_Bonus = "", 0, 1)

        For i As Integer = 0 To Times
            Dim InputValue As String = "GameCustId|" & GameCustId & "|GamePoint|" & GamePoint.ToString() & "|TradeSeq|" & TradeSeq & "|GameNo|" & GameNo & "|GameNo_Bonus|" & GameNo_Bonus
            ' 建立命令(預存程序名,使用連線)
            '20110614 qq 修改連新的SP
            Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_InGameSave_GetGamePoint_A02", Con)
            ' 設置命令型態為預存程序 
            Cmd.CommandType = CommandType.StoredProcedure
            Cmd.CommandTimeout = 300
            ' 增加參數欄位
            Cmd.Parameters.Add("@GameCustId", SqlDbType.VarChar, 100).Value = GameCustId
            Cmd.Parameters.Add("@GamePoint", SqlDbType.Money).Value = GamePoint
            If i = 0 Then
                Cmd.Parameters.Add("@GameNo", SqlDbType.VarChar, 50).Value = GameNo
                Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20).Value = TradeSeq
            Else
                Cmd.Parameters.Add("@GameNo", SqlDbType.VarChar, 50).Value = GameNo_Bonus
                Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20).Value = TradeSeq & "-B"
            End If

            Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
            ReturnMsgNo.Direction = ParameterDirection.Output
            Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
            ReturnMsg.Direction = ParameterDirection.Output
            Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
            ErrorCode.Direction = ParameterDirection.Output
            Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
            LogSn.Direction = ParameterDirection.Output

            Try
                Con.Open()
                Cmd.ExecuteNonQuery()
                ReturnValue.ReturnMsg = ReturnMsg.Value
                ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
                ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
                ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            Catch ex As Exception
                MyCardDbwebservice.MyCardErrorLog("MyCardServiceForInGame.asmx", "999|MyCardCPSSInGameSaveGetGamePoint|" & InputValue & "|" & ex.Message, "127.0.0.3")
                ReturnValue.ReturnMsgNo = "-999"
                ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
                ReturnValue.ReturnErrorCode = "FBWS2004"
                ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            Finally
                Con.Close()
            End Try

            If CInt(ReturnMsgNo.Value) <> 1 Then
                MyCardDbwebservice.MyCardErrorLog("MyCardServiceForInGame.asmx", "MyCardCPSSInGameSaveGetGamePoint|" & InputValue & "|" & ReturnMsgNo.Value.ToString() & "|" & ReturnMsg.Value.ToString() & "|" & ErrorCode.Value.ToString(), "127.0.0.3")
            End If
        Next

        Return ReturnValue
    End Function

    ''' <summary>
    '''   MyCardInGame差報
    ''' </summary>
    ''' <param name="StartDate">起始,格式String</param>
    ''' <param name="EndDate">結束,格式Integer</param>
    ''' <param name="MyCardID">卡號,格式Integer</param>
    ''' <returns>DataSet回傳MyCardCardId, GameCustId, TradeSeq, SaveDate</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function MyCardCPSaveInGameSaveTradeListDiffReport(ByVal StartDate As String, ByVal EndDate As String, ByVal MyCardID As String) As DataSet
        ConnStr(15)
        ' 建立命令(預存程序名,使用連線) 
        Dim Cmd As New SqlClient.SqlCommand
        ' 設置命令型態為預存程序 
        Cmd.Connection = Con
        Cmd.CommandType = CommandType.Text
        Cmd.CommandTimeout = 300
        ' 增加參數欄位

        If StartDate <> "" And EndDate <> "" Then
            Cmd.CommandText = "Select MyCardCardId, GameCustId, TradeSeq, SaveDate From MyCard_CPSaveService.dbo.CPSave_InGame_SaveTradeList_DiffReport WHERE SaveDate BETWEEN @StartDate AND @EndDate  "
            Cmd.Parameters.AddWithValue("@StartDate", StartDate & " 00:00:00")
            Cmd.Parameters.AddWithValue("@EndDate", EndDate & " 23:59:59")
        ElseIf MyCardID <> "" Then
            Cmd.CommandText = "Select MyCardCardId, GameCustId, TradeSeq, SaveDate From MyCard_CPSaveService.dbo.CPSave_InGame_SaveTradeList_DiffReport WHERE MyCardCardId=@MyCardCardId  "
            Cmd.Parameters.AddWithValue("@MyCardCardId", MyCardID)
        Else
            Cmd.CommandText = "Select MyCardCardId, GameCustId, TradeSeq, SaveDate From MyCard_CPSaveService.dbo.CPSave_InGame_SaveTradeList_DiffReport WHERE 1=2  "
        End If

        Dim Adp As SqlClient.SqlDataAdapter
        Adp = New SqlClient.SqlDataAdapter(Cmd)

        Dim DBViewDS As New DataSet
        Try
            Adp.Fill(DBViewDS)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|MyCardCPSaveInGameSaveTradeListDiffReport|" & ex.Message, "127.0.0.3")
            Throw New Exception("Service的View條件資料失敗", ex)
        End Try
        Return DBViewDS

    End Function
    ''' <summary>
    ''' Billing初始
    ''' </summary>
    ''' <param name="ProductServiceId">產品代碼</param>
    ''' <param name="TradePrice">產品金額</param>
    ''' <param name="UserId">使用者帳號</param>
    ''' <param name="UserIp">使用者IP</param>
    ''' <param name="TradeType">1:FACEBOOK 2會員加點IP</param>
    ''' <returns>TradeSeq:儲值交易序號;GamePrice:facebook點數;ReturnMsgNo;ReturnMsg;ErrorCode;LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function CPSS_BillingSave_GetProductInfo(ByVal ProductServiceId As String, ByVal TradePrice As Integer, ByVal UserId As String, ByVal UserIp As String, ByVal TradeType As Integer) As CPSS_BillingSave_GetProductInfoResult
        ConnStr(15)
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_BillingSave_GetProductInfo_A01", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@ProductServiceId", SqlDbType.VarChar, 10).Value = ProductServiceId '--產品代碼
        Cmd.Parameters.Add("@TradePrice", SqlDbType.Int).Value = TradePrice '--產品金額
        Cmd.Parameters.Add("@UserId", SqlDbType.VarChar, 100).Value = UserId ' --使用者帳號
        Cmd.Parameters.Add("@UserIp", SqlDbType.VarChar, 25).Value = UserIp '--使用者IP
        Cmd.Parameters.Add("@TradeType", SqlDbType.Int).Value = TradeType '--1:FACEBOOK 2會員加點IP

        Dim TradeSeq As SqlClient.SqlParameter = Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20) '--儲值交易序號
        TradeSeq.Direction = ParameterDirection.Output
        Dim GamePrice As SqlClient.SqlParameter = Cmd.Parameters.Add("@GamePrice", SqlDbType.Int)
        GamePrice.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output

        Dim ReturnValue As New CPSS_BillingSave_GetProductInfoResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
            ReturnValue.ReturnTradeSeq = IIf(IsDBNull(TradeSeq.Value), "", TradeSeq.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.GamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|CPSS_BillingSave_GetProductInfo|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2005"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.GamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function
    ''' <summary>
    ''' Billing記錄Authcode
    ''' </summary>
    ''' <param name="TradeSeq">儲值交易序號</param>
    ''' <param name="BillingAuthCode">billing驗證碼</param>
    ''' <returns>ReturnMsgNo;ReturnMsg;ErrorCode;LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function CPSS_BillingSave_GetBillingAuthCode(ByVal TradeSeq As String, ByVal BillingAuthCode As String) As CPSS_BillingSave_GetBillingAuthCodeResult
        ConnStr(15)
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_BillingSave_GetBillingAuthCode_A01", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20).Value = TradeSeq '--儲值交易序號
        Cmd.Parameters.Add("@BillingAuthCode", SqlDbType.VarChar, 100).Value = BillingAuthCode '--billing驗證碼

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output
        Dim ReturnValue As New CPSS_BillingSave_GetBillingAuthCodeResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|CPSS_BillingSave_GetBillingAuthCode|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2006"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function
    ''' <summary>
    ''' Billing儲值成功，但要改CPSaveService的table狀態
    ''' </summary>
    ''' <param name="BillingAuthCode">billing驗證碼</param>
    ''' <returns>TradeSeq:儲值交易序號;TradePrice:價格;GamePrice:fb點數;ReturnMsgNo;ReturnMsg;ErrorCode;LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function CPSS_BillingSave_RequestSaveGame(ByVal BillingAuthCode As String) As CPSS_BillingSave_RequestSaveGameResult
        ConnStr(15)
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_BillingSave_RequestSaveGame_A02", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@BillingAuthCode", SqlDbType.VarChar, 100).Value = BillingAuthCode '--billing驗證碼

        Dim TradeSeq As SqlClient.SqlParameter = Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20) '--儲值交易序號
        TradeSeq.Direction = ParameterDirection.Output
        Dim TradePrice As SqlClient.SqlParameter = Cmd.Parameters.Add("@TradePrice", SqlDbType.Int)
        TradePrice.Direction = ParameterDirection.Output
        Dim GamePrice As SqlClient.SqlParameter = Cmd.Parameters.Add("@GamePrice", SqlDbType.Int)
        GamePrice.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output
        Dim Discount As SqlClient.SqlParameter = Cmd.Parameters.Add("@Discount", SqlDbType.VarChar, 7)
        Discount.Direction = ParameterDirection.Output
        Dim ReturnValue As New CPSS_BillingSave_RequestSaveGameResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
            ReturnValue.ReturnTradeSeq = IIf(IsDBNull(TradeSeq.Value), "", TradeSeq.Value)
            ReturnValue.TradePrice = IIf(IsDBNull(TradePrice.Value), 0, TradePrice.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.GamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)
            ReturnValue.CostPoint = IIf(IsDBNull(Discount.Value), 0, Discount.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|CPSS_BillingSave_RequestSaveGame|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2007"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.GamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)
            ReturnValue.CostPoint = IIf(IsDBNull(Discount.Value), 0, Discount.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function
    ''' <summary>
    ''' 加fb點數，改CPSaveService的table狀態
    ''' </summary>
    ''' <param name="TradeSeq">儲值交易序號</param>
    ''' <param name="GamePrice">轉換後的點數</param>
    ''' <param name="GameNo">fb交易序號</param>
    ''' <returns>BillingAuthCode:billing驗證碼;GameCustId:使用者帳號;ReturnMsgNo;ReturnMsg;ErrorCode;LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function CPSS_BillingSave_ResponseSaveGame(ByVal TradeSeq As String, ByVal GamePrice As Integer, ByVal GameNo As String) As CPSS_BillingSave_ResponseSaveGameResult
        ConnStr(15)
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_BillingSave_ResponseSaveGame_A01", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20).Value = TradeSeq '--儲值交易序號
        Cmd.Parameters.Add("@GamePrice", SqlDbType.Int).Value = GamePrice '--轉換後的點數
        Cmd.Parameters.Add("@GameNo ", SqlDbType.VarChar, 50).Value = GameNo  '--轉換後的點數

        Dim BillingAuthCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@BillingAuthCode", SqlDbType.VarChar, 100) '--儲值交易序號
        BillingAuthCode.Direction = ParameterDirection.Output
        Dim GameCustId As SqlClient.SqlParameter = Cmd.Parameters.Add("@GameCustId", SqlDbType.VarChar, 100) '--儲值交易序號
        GameCustId.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output
        Dim ReturnValue As New CPSS_BillingSave_ResponseSaveGameResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
            ReturnValue.BillingAuthCode = IIf(IsDBNull(BillingAuthCode.Value), "", BillingAuthCode.Value)
            ReturnValue.GameCustId = IIf(IsDBNull(GameCustId.Value), "", GameCustId.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|CPSS_BillingSave_ResponseSaveGame|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2008"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function
    ''' <summary>
    ''' 查詢交易是否成功
    ''' </summary>
    ''' <param name="TradeSeq">儲值交易序號</param>
    ''' <returns>GamePrice:fb點數;ReturnMsgNo;ReturnMsg;ErrorCode;LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function CPSS_BillingSave_CheckSaveTrade(ByVal TradeSeq As String) As CPSS_BillingSave_CheckSaveTradeResult
        ConnStr(15)
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_BillingSave_CheckSaveTrade_A01", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20).Value = TradeSeq '--儲值交易序號

        Dim TradePrice As SqlClient.SqlParameter = Cmd.Parameters.Add("@GamePrice", SqlDbType.Int) '--儲值交易序號
        TradePrice.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output
        Dim ReturnValue As New CPSS_BillingSave_CheckSaveTradeResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
            ReturnValue.TradePrice = IIf(IsDBNull(TradePrice.Value), 0, TradePrice.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|CPSS_BillingSave_CheckSaveTrade|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2009"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function
    ''' <summary>
    ''' 會員扣點初始
    ''' </summary>
    ''' <param name="GameFacId">廠商ID</param>
    ''' <param name="GameFacSerId">廠商服務ID</param>
    ''' <param name="OneTimePassport">登入機碼</param>
    ''' <param name="GameCustId">會員帳號</param>
    ''' <param name="GameUserId">UID帳號</param>
    ''' <param name="CostPoint">扣點點數</param>
    ''' <param name="CreateIP">IP</param>
    ''' <returns>ReturnTradeSeq:扣點交易序號;SavePoint:轉換儲值點數;ReturnMsgNo;ReturnMsg;ErrorCode;LogSn</returns>
    ''' <remarks></remarks>
    ''' 20110216 多加一個記錄uid的欄位 詹旻融
    <WebMethod()> _
    Public Function CPSS_MemberCostSaveTradeList_Insert(ByVal GameFacId As String, ByVal GameFacSerId As String, ByVal OneTimePassport As String, ByVal GameCustId As String, ByVal CostPoint As Integer, ByVal CreateIP As String, ByVal GameUserId As String) As CPSS_MemberCostSaveTradeList_InsertResult
        ConnStr(15)
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_MemberCostSaveTradeList_Insert_A01", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@GameFacId", SqlDbType.NVarChar, 50).Value = GameFacId
        Cmd.Parameters.Add("@GameFacSerId", SqlDbType.NVarChar, 50).Value = GameFacSerId
        Cmd.Parameters.Add("@OneTimePassport", SqlDbType.VarChar, 50).Value = OneTimePassport
        Cmd.Parameters.Add("@GameCustId", SqlDbType.VarChar, 100).Value = GameCustId
        Cmd.Parameters.Add("@GameUserId", SqlDbType.VarChar, 100).Value = GameUserId
        Cmd.Parameters.Add("@CostPoint", SqlDbType.Int).Value = CostPoint
        Cmd.Parameters.Add("@CreateIP", SqlDbType.VarChar, 30).Value = CreateIP

        Dim TradeSeq As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnTradeSeq", SqlDbType.VarChar, 20)
        TradeSeq.Direction = ParameterDirection.Output
        Dim SavePoint As SqlClient.SqlParameter = Cmd.Parameters.Add("@SavePoint", SqlDbType.Int)
        SavePoint.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output
        Dim ReturnValue As New CPSS_MemberCostSaveTradeList_InsertResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
            ReturnValue.TradeSeq = IIf(IsDBNull(TradeSeq.Value), "", TradeSeq.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.SavePoint = IIf(IsDBNull(SavePoint.Value), 0, SavePoint.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|CPSS_MemberCostSaveTradeList_Insert|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2010"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.TradeSeq = ""
            ReturnValue.SavePoint = 0
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function
    ''' <summary>
    ''' 記錄扣點AuthCode
    ''' </summary>
    ''' <param name="TradeSeq">扣點交易序號</param>
    ''' <param name="AuthCode">授權碼</param>
    ''' <returns>ReturnMsgNo;ReturnMsg;ErrorCode;LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function CPSS_MemberCostAuthCode_Insert(ByVal TradeSeq As String, ByVal AuthCode As String) As CPSS_MemberCostAuthCode_InsertResult
        ConnStr(15)
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_MemberCostAuthCode_Insert_A01", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20).Value = TradeSeq
        Cmd.Parameters.Add("@AuthCode", SqlDbType.VarChar, 100).Value = AuthCode

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output
        Dim ReturnValue As New CPSS_MemberCostAuthCode_InsertResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|CPSS_MemberCostAuthCode_Insert|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2011"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function
    ''' <summary>
    ''' 扣點成功,更新table狀態
    ''' </summary>
    ''' <param name="TradeSeq"></param>
    ''' <returns>SavePoint:儲值兌換點數;ReturnMsgNo;ReturnMsg;ErrorCode;LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function CPSS_MemberCost_RequestSaveGame(ByVal TradeSeq As String) As CPSS_MemberCost_RequestSaveGameResult
        ConnStr(15)
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_MemberCost_RequestSaveGame_A02", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20).Value = TradeSeq

        Dim SavePoint As SqlClient.SqlParameter = Cmd.Parameters.Add("@SavePoint", SqlDbType.Int)
        SavePoint.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output
        Dim Discount As SqlClient.SqlParameter = Cmd.Parameters.Add("@Discount", SqlDbType.VarChar, 7)
        Discount.Direction = ParameterDirection.Output
        Dim ReturnValue As New CPSS_MemberCost_RequestSaveGameResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.SavePoint = IIf(IsDBNull(SavePoint.Value), 0, SavePoint.Value)
            ReturnValue.CostPoint = IIf(IsDBNull(Discount.Value), 0, Discount.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|CPSS_MemberCost_RequestSaveGame|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2012"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.SavePoint = IIf(IsDBNull(SavePoint.Value), 0, SavePoint.Value)
            ReturnValue.CostPoint = IIf(IsDBNull(Discount.Value), 0, Discount.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function
    ''' <summary>
    ''' 扣點流程完成
    ''' </summary>
    ''' <param name="TradeSeq">扣點交易序號</param>
    ''' <param name="GameNo">遊戲方訂單代號</param>
    ''' <returns>ReturnMsgNo;ReturnMsg;ErrorCode;LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function CPSS_MemberCost_ResponseSaveGame(ByVal TradeSeq As String, ByVal GameNo As String) As CPSS_MemberCost_ResponseSaveGameResult
        ConnStr(15)
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_MemberCost_ResponseSaveGame_A01", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20).Value = TradeSeq
        Cmd.Parameters.Add("@GameNo", SqlDbType.VarChar, 50).Value = GameNo

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output
        Dim ReturnValue As New CPSS_MemberCost_ResponseSaveGameResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|CPSS_MemberCost_ResponseSaveGame|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2013"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function

    <WebMethod()> _
    Public Function FBCheckUID(ByVal UID As String, ByVal MyCardCustIp As String) As FBCheckUIDResult
        ConnStr(19)
        '20110611 qq 增加回傳ErrorCode,LogSn
        Dim Cmd As New SqlClient.SqlCommand("MyCard_Member.dbo.SPS_FB_MC_CheckUID_A01", Con)
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        Cmd.Parameters.Add("@UID", SqlDbType.VarChar, 100).Value = UID
        Cmd.Parameters.Add("@MyCardCustIp", SqlDbType.NVarChar, 30).Value = MyCardCustIp

        Dim uMyCardCustId As SqlClient.SqlParameter = Cmd.Parameters.Add("@uMyCardCustId", SqlDbType.NVarChar, 50)
        uMyCardCustId.Direction = ParameterDirection.Output

        Dim uPOINTS As SqlClient.SqlParameter = Cmd.Parameters.Add("@uPOINTS", SqlDbType.Int)
        uPOINTS.Direction = ParameterDirection.Output

        Dim uUSER_NAME As SqlClient.SqlParameter = Cmd.Parameters.Add("@uUSER_NAME", SqlDbType.NVarChar, 50)
        uUSER_NAME.Direction = ParameterDirection.Output

        Dim MyCardOTP As SqlClient.SqlParameter = Cmd.Parameters.Add("@MyCardOTP", SqlDbType.NVarChar, 50)
        MyCardOTP.Direction = ParameterDirection.Output

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 100)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ReturnErrorCode.Direction = ParameterDirection.Output

        Dim ReturnLogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        ReturnLogSn.Direction = ParameterDirection.Output

        Dim ReturnValue As New FBCheckUIDResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.uMyCardCustId = uMyCardCustId.Value
            ReturnValue.uPOINTS = uPOINTS.Value
            ReturnValue.uUSER_NAME = uUSER_NAME.Value
            ReturnValue.MyCardOTP = MyCardOTP.Value
            ReturnValue.ReturnMsg = ReturnMsg.Value
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ReturnErrorCode.Value), "", ReturnErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|FBCheckUID|" & ex.Message, "127.0.0.3")
            ReturnValue.uMyCardCustId = ""
            ReturnValue.uPOINTS = 0
            ReturnValue.uUSER_NAME = ""
            ReturnValue.MyCardOTP = ""
            ReturnValue.ReturnMsgNo = -999
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBMB0001"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function

    <WebMethod()> _
    Public Function FBMemberLoginCheck(ByVal MyCardCustId As String, ByVal MyCardCustPwd As String, ByVal MyCardCustIp As String) As FBMemberLoginCheckResult
        ConnStr(19)
        '20110611 qq 修改連新的SP 新增回傳ErrorCode,LogSn
        Dim Cmd As New SqlClient.SqlCommand("MyCard_Member.dbo.SPS_FB_MC_MemberLoginCheck_A01", Con)
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        Cmd.Parameters.Add("@MyCardCustId", SqlDbType.VarChar, 50).Value = MyCardCustId
        Cmd.Parameters.Add("@MyCardCustPwd", SqlDbType.VarChar, 100).Value = MyCardCustPwd
        Cmd.Parameters.Add("@MyCardCustIp", SqlDbType.VarChar, 30).Value = MyCardCustIp

        Dim MemberType As SqlClient.SqlParameter = Cmd.Parameters.Add("@MemberType", SqlDbType.TinyInt)
        MemberType.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 1024)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ReturnErrorCode.Direction = ParameterDirection.Output

        Dim ReturnLogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        ReturnLogSn.Direction = ParameterDirection.Output

        Dim ReturnValue As New FBMemberLoginCheckResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.MemberType = MemberType.Value
            ReturnValue.ReturnMsg = ReturnMsg.Value
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ReturnErrorCode.Value), "", ReturnErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|FBMemberLoginCheck|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = -999
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBMB0002"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function

    <WebMethod()> _
    Public Function FBSimpleMemberRegister(ByVal MemberEmail As String) As FBSimpleMemberRegisterResult
        ConnStr(19)
        '20110611 qq 修改連新的SP 新增回傳ErrorCode,LogSn
        Dim Cmd As New SqlClient.SqlCommand("MyCard_Member.dbo.SPS_FB_MC_SimpleMemberRegister_A01", Con)
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        Cmd.Parameters.Add("@MemberEmail", SqlDbType.VarChar, 50).Value = MemberEmail
        'Cmd.Parameters.Add("@UserSavePassword", SqlDbType.VarChar, 300).Value = UserSavePassword

        Dim ReturnSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@uSn", SqlDbType.VarChar, 250)
        ReturnSn.Direction = ParameterDirection.Output

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ReturnErrorCode.Direction = ParameterDirection.Output

        Dim ReturnLogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        ReturnLogSn.Direction = ParameterDirection.Output

        Dim ReturnValue As New FBSimpleMemberRegisterResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.Sn = ReturnSn.Value
            ReturnValue.ReturnMsg = ReturnMsg.Value
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ReturnErrorCode.Value), "", ReturnErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|FBSimpleMemberRegister|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = -999
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBMB0003"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function
   
    <WebMethod()> _
    Public Function FBSimpleMemberRegisterPWDUpdate(ByVal MemberEmail As String, ByVal UserSavePassword As String) As FBSimpleMemberRegisterPWDUpdateResult
        ConnStr(19)
        '20110611 qq 修改連新的SP 新增回傳ErrorCode,LogSn
        Dim Cmd As New SqlClient.SqlCommand("MyCard_Member.dbo.SPS_FB_MC_SimpleMemberRegisterPWD_Update_A01", Con)
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        Cmd.Parameters.Add("@MemberEmail", SqlDbType.VarChar, 50).Value = MemberEmail
        Cmd.Parameters.Add("@UserSavePassword", SqlDbType.VarChar, 300).Value = UserSavePassword

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ReturnErrorCode.Direction = ParameterDirection.Output

        Dim ReturnLogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        ReturnLogSn.Direction = ParameterDirection.Output

        Dim ReturnValue As New FBSimpleMemberRegisterPWDUpdateResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = ReturnMsg.Value
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ReturnErrorCode.Value), "", ReturnErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|FBSimpleMemberRegisterPWDUpdate|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = -999
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBMB0004"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function

    <WebMethod()> _
    Public Function FBSimpleMemberUpdate(ByVal MyCardCustId As String, ByVal UserSavePassword As String, ByVal PERSONAL_ID As String, ByVal UserCountry As Integer, ByVal UserName As String, ByVal UserNickname As String, ByVal UserBirthDay As String, ByVal UserGender As Integer, ByVal UserPhoneCode As String, ByVal UserPhoneNumber As String, ByVal UserAddress As String, ByVal UserMobile As String, ByVal UserEnews As Integer, ByVal UserIP As String) As FBSimpleMemberUpdateResult
        ConnStr(19)
        '20110611 qq 修改連新的SP 新增回傳ErrorCode,LogSn
        Dim Cmd As New SqlClient.SqlCommand("MyCard_Member.dbo.SPS_FB_MC_SimpleMemberUpdate_A01", Con)
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        Cmd.Parameters.Add("@MyCardCustId", SqlDbType.VarChar, 50).Value = MyCardCustId
        Cmd.Parameters.Add("@UserSavePassword", SqlDbType.VarChar, 300).Value = UserSavePassword
        Cmd.Parameters.Add("@PERSONAL_ID", SqlDbType.VarChar, 40).Value = PERSONAL_ID
        Cmd.Parameters.Add("@UserCountry", SqlDbType.Int).Value = UserCountry
        Cmd.Parameters.Add("@UserName", SqlDbType.NVarChar, 50).Value = UserName
        Cmd.Parameters.Add("@UserNickname", SqlDbType.NVarChar, 50).Value = UserNickname
        Cmd.Parameters.Add("@UserBirthDay", SqlDbType.DateTime).Value = UserBirthDay
        Cmd.Parameters.Add("@UserGender", SqlDbType.Char).Value = UserGender
        Cmd.Parameters.Add("@UserPhoneCode", SqlDbType.VarChar, 4).Value = UserPhoneCode
        Cmd.Parameters.Add("@UserPhoneNumber", SqlDbType.VarChar, 20).Value = UserPhoneNumber
        Cmd.Parameters.Add("@UserAddress", SqlDbType.VarChar, 50).Value = UserAddress
        Cmd.Parameters.Add("@UserMobile", SqlDbType.VarChar, 20).Value = UserMobile
        Cmd.Parameters.Add("@UserEnews", SqlDbType.Char).Value = UserEnews
        Cmd.Parameters.Add("@UserIP", SqlDbType.VarChar, 50).Value = UserIP

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 1024)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ReturnErrorCode.Direction = ParameterDirection.Output

        Dim ReturnLogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        ReturnLogSn.Direction = ParameterDirection.Output

        Dim ReturnValue As New FBSimpleMemberUpdateResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnMsg = ReturnMsg.Value
            
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ReturnErrorCode.Value), "", ReturnErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|FBSimpleMemberUpdate|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = -999
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBMB0005"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function
   
    <WebMethod()> _
    Public Function FBUserBindingInsert(ByVal Uid As String, ByVal MyCardCustId As String, ByVal CreateIP As String) As FBUserBindingInsertResult
        ConnStr(19)
        '20110611 qq 修改連新的SP 新增回傳ErrorCode,LogSn
        Dim Cmd As New SqlClient.SqlCommand("MyCard_Member.dbo.SPS_FB_MC_UserBinding_Insert_A01", Con)
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        Cmd.Parameters.Add("@Uid", SqlDbType.VarChar, 100).Value = Uid
        Cmd.Parameters.Add("@MyCardCustId", SqlDbType.VarChar, 50).Value = MyCardCustId
        Cmd.Parameters.Add("@CreateIP", SqlDbType.VarChar, 50).Value = CreateIP

        Dim uPOINTS As SqlClient.SqlParameter = Cmd.Parameters.Add("@uPOINTS", SqlDbType.Int)
        uPOINTS.Direction = ParameterDirection.Output

        Dim uUSER_NAME As SqlClient.SqlParameter = Cmd.Parameters.Add("@uUSER_NAME", SqlDbType.NVarChar, 50)
        uUSER_NAME.Direction = ParameterDirection.Output

        Dim MyCardOTP As SqlClient.SqlParameter = Cmd.Parameters.Add("@MyCardOTP", SqlDbType.NVarChar, 50)
        MyCardOTP.Direction = ParameterDirection.Output

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 50)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ReturnErrorCode.Direction = ParameterDirection.Output

        Dim ReturnLogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        ReturnLogSn.Direction = ParameterDirection.Output

        Dim ReturnValue As New FBUserBindingInsertResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.uPOINTS = uPOINTS.Value
            ReturnValue.uUSER_NAME = uUSER_NAME.Value
            ReturnValue.MyCardOTP = MyCardOTP.Value
            ReturnValue.ReturnMsg = ReturnMsg.Value
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ReturnErrorCode.Value), "", ReturnErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|FBUserBindingInsert|" & ex.Message, "127.0.0.3")
            ReturnValue.uPOINTS = 0
            ReturnValue.uUSER_NAME = ""
            ReturnValue.MyCardOTP = ""
            ReturnValue.ReturnMsgNo = -999
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBMB0006"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function

    <WebMethod()> _
    Public Function ViewFBMemberSn(ByVal USER_ID As String) As ViewFBMemberSnResult
        ConnStr(19)

        Dim Cmd As New SqlClient.SqlCommand("select * from MyCard_Member.dbo.View_FB_MC_MemberSn where USER_ID=@USER_ID", Con)
        Cmd.CommandType = CommandType.Text
        Cmd.CommandTimeout = 300
        Cmd.Parameters.Add("@USER_ID", SqlDbType.VarChar, 50).Value = USER_ID

        Dim ReturnValue As New ViewFBMemberSnResult
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(Cmd)
        Dim DSssid As New DataSet
        Try
            Con.Open()
            SDAssid.Fill(DSssid)

            If DSssid.Tables(0).Rows.Count = 0 Then
                ReturnValue.ReturnMsg = "查無此會員"
                ReturnValue.ReturnMsgNo = -1
            Else
                ReturnValue.Sn = DSssid.Tables(0).Rows(0).Item("Sn")
                ReturnValue.ReturnMsg = ""
                ReturnValue.ReturnMsgNo = 1
            End If
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|ViewFBMemberSn|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = -999
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
        Finally
            Con.Close()
        End Try
        Return ReturnValue

    End Function
    ''' <summary>
    ''' 扣點品項
    ''' </summary>
    ''' <param name="FacId">廠商id</param>
    ''' <returns>Dataset(FacId, FacSerId, SerId,ServiceName,Price);ReturnMsgNo;ReturnMsg</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function MyCard_MemberService_GetElectronicWallet_FactoryService(ByVal FacId As String) As MyCard_MemberService_GetElectronicWallet_FactoryServiceResult
        ConnStr(20)

        Dim Cmd As New SqlClient.SqlCommand("MyCard_MemberService.dbo.MyCard_MemberService_GetElectronicWallet_FactoryService", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        Cmd.Parameters.Add("@FacId", SqlDbType.NVarChar, 50).Value = FacId

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgno", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.VarChar, 50)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New MyCard_MemberService_GetElectronicWallet_FactoryServiceResult
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(Cmd)
        Dim DSssid As New DataSet
        Try
            Con.Open()
            SDAssid.Fill(DSssid)

            If DSssid.Tables(0).Rows.Count = 0 Then
                ReturnValue.ReturnMsg = "查無資料"
                ReturnValue.ReturnMsgNo = -1
            Else
                ReturnValue.ReturnDS = DSssid
                ReturnValue.ReturnMsg = ""
                ReturnValue.ReturnMsgNo = 1
            End If
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|MyCard_MemberService_GetElectronicWallet_FactoryService|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = -999
            ReturnValue.ReturnMsg = "系統發生錯誤-" & ex.Message
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function

    ''' <summary>
    '''   取得遊戲廠商資料，前台首頁配合遊戲區使用
    ''' </summary>
    ''' <returns>DataSet回傳欄位Sn, GameFactoryName, PICAddress, Status, Order</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_CPSave_FaceBookGameFactory() As DataSet
        ConnStr(15)
        ' 建立命令(預存程序名,使用連線) 
        Dim Cmd As New SqlClient.SqlCommand
        ' 設置命令型態為預存程序 
        Cmd.Connection = Con
        Cmd.CommandType = CommandType.Text
        Cmd.CommandTimeout = 300
        ' 增加參數欄位

        Cmd.CommandText = "Select Top 8 Sn, GameFactoryName, PICAddress, Status, [Order] From MyCard_CPSaveService.dbo.VIEW_CPSave_FaceBookGameFactory WHERE Status = 1 order by [Order]"
        
        Dim Adp As SqlClient.SqlDataAdapter
        Adp = New SqlClient.SqlDataAdapter(Cmd)

        Dim DBViewDS As New DataSet
        Try
            Adp.Fill(DBViewDS)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|Get_CPSave_FaceBookGameFactory|" & ex.Message, "127.0.0.3")
            Throw New Exception(ex.ToString)
        End Try
        Return DBViewDS

    End Function

    ''' <summary>
    '''   取得廠商遊戲資料，前台廠商遊戲總頁用
    ''' </summary>
    ''' <param name="GameFactorySn">遊戲廠商Sn</param>
    ''' <returns>DataSet回傳欄位Sn, GameName,GameFactorySn,GameFactoryName,FacPICAddress, PICAddress,GameLink, Status, [Order]</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_CPSave_FaceBookGame(ByVal GameFactorySn As Integer) As DataSet
        ConnStr(15)
        ' 建立命令(預存程序名,使用連線) 
        Dim Cmd As New SqlClient.SqlCommand
        ' 設置命令型態為預存程序 
        Cmd.Connection = Con
        Cmd.CommandType = CommandType.Text
        Cmd.CommandTimeout = 300
        ' 增加參數欄位

        Cmd.CommandText = "Select Sn, GameName,GameFactorySn,GameFactoryName,FacPICAddress, PICAddress,GameLink, Status, [Order] From MyCard_CPSaveService.dbo.VIEW_CPSave_FaceBookGame WHERE GameFactorySn=@GameFactorySn and Status = 1 order by [Order]"
        Cmd.Parameters.Add("@GameFactorySn", SqlDbType.Int).Value = GameFactorySn

        Dim Adp As SqlClient.SqlDataAdapter
        Adp = New SqlClient.SqlDataAdapter(Cmd)

        Dim DBViewDS As New DataSet
        Try
            Adp.Fill(DBViewDS)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|Get_CPSave_FaceBookGame|" & ex.Message, "127.0.0.3")
            Throw New Exception(ex.ToString)
        End Try
        Return DBViewDS

    End Function

    ''' <summary>
    '''   取得遊戲介紹資料，前台遊戲介紹頁用
    ''' </summary>
    ''' <param name="GameSn">遊戲Sn</param>
    ''' <param name="LanguageID">語系代碼</param>
    ''' <returns>DataSet回傳欄位Sn, LanguageSn, LanguageID, Language, GameSn, GameName,GameLink, GameInfo,PICAddress</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_CPSave_FaceBookGameInfo(ByVal GameSn As Integer, ByVal LanguageID As String) As DataSet
        ConnStr(15)
        ' 建立命令(預存程序名,使用連線) 
        Dim Cmd As New SqlClient.SqlCommand
        ' 設置命令型態為預存程序 
        Cmd.Connection = Con
        Cmd.CommandType = CommandType.Text
        Cmd.CommandTimeout = 300
        ' 增加參數欄位

        Cmd.CommandText = "Select Sn, LanguageSn, LanguageID, Language, GameSn, GameName,GameLink, GameInfo,PICAddress From MyCard_CPSaveService.dbo.VIEW_CPSave_FaceBookGameInfo WHERE GameSn=@GameSn and LanguageID=@LanguageID"
        Cmd.Parameters.Add("@GameSn", SqlDbType.Int).Value = GameSn
        Cmd.Parameters.Add("@LanguageID", SqlDbType.VarChar, 30).Value = LanguageID

        Dim Adp As SqlClient.SqlDataAdapter
        Adp = New SqlClient.SqlDataAdapter(Cmd)

        Dim DBViewDS As New DataSet
        Try
            Adp.Fill(DBViewDS)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|Get_CPSave_FaceBookGameInfo|" & ex.Message, "127.0.0.3")
            Throw New Exception(ex.ToString)
        End Try
        Return DBViewDS

    End Function

    ''' <summary>
    ''' 檢查會員狀態(暫時凍結,永久凍結,點數鎖定)
    ''' </summary>
    ''' <param name="MemberId">會員帳號</param>
    ''' <param name="MemberIp">IP</param>
    ''' <returns>ReturnMsgNo,ReturnMsg,ReturnErrorCode,ReturnLogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function MyCard_MemberCheckStatus(ByVal MemberId As String, ByVal MemberIp As String) As MemberCheckStatusResult
        ConnStr(19)

        Dim Cmd As New SqlClient.SqlCommand("MyCard_Member.dbo.MyCard_MemberCheckStatus", Con)
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        Cmd.Parameters.Add("@MemberId", SqlDbType.VarChar, 50).Value = MemberId
        Cmd.Parameters.Add("@MemberIp", SqlDbType.VarChar, 25).Value = MemberIp

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 1024)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ReturnErrorCode.Direction = ParameterDirection.Output

        Dim ReturnLogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        ReturnLogSn.Direction = ParameterDirection.Output

        Dim ReturnValue As New MemberCheckStatusResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = ReturnMsg.Value
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ReturnErrorCode.Value), "", ReturnErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|MyCard_MemberCheckStatus|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = -999
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBMB0007"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function

    ''' <summary>
    ''' 會員帳號限制功能(-10 超過會員上限-20需綁定晶片卡或Motp)
    ''' </summary>
    ''' <param name="MemberId">MyCard會員帳號</param>
    ''' <param name="TradePoint">點數</param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 
    ''' </remarks>
    <WebMethod()> _
    Public Function MyCard_MemberIdTradeLimitForFB_A01(ByVal MemberId As String, ByVal TradePoint As Integer, ByVal CreateIP As String) As MyCard_MemberIdTradeLimitForFB_A01Result
        ConnStr(19)
        Dim Cmd As New SqlClient.SqlCommand("MyCard_Member.dbo.MyCard_MemberIdTradeLimitForFB_A01", Con)
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@MemberId", SqlDbType.VarChar, 50).Value = MemberId
        Cmd.Parameters.Add("@TradePoint", SqlDbType.Int).Value = TradePoint

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 1024)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ReturnErrorCode.Direction = ParameterDirection.Output

        Dim ReturnLogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        ReturnLogSn.Direction = ParameterDirection.Output

        Dim ReturnValue As New MyCard_MemberIdTradeLimitForFB_A01Result
        Dim Str1 As String = MemberId & "," & TradePoint

        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = ReturnMsg.Value
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ReturnErrorCode.Value), "", ReturnErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|MyCard_MemberIdTradeLimitForFB_A01|" & Str1 & "|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = -999
            ReturnValue.ReturnMsg = ex.Message
            ReturnValue.ReturnErrorCode = "FBMB0008"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(ReturnLogSn.Value), 0, ReturnLogSn.Value)
            Return ReturnValue
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function


    ''' <summary>
    ''' 2013-9-9
    ''' FB新API 卡片儲值
    ''' </summary>
    ''' <param name="TradeSeq">交易序號,格式String</param>
    ''' <param name="CardPoint">點數,格式Integer</param>
    ''' <param name="oProjNo">活動代號,格式String</param>
    ''' <param name="CardKind">卡通路別(2實體,1虛擬),格式Integer</param>
    ''' <returns>ReturnNo 回傳代碼 ,ReturnMsg 回傳訊息,ErrorCode 回傳ErrorCode,LogSn 回傳LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function MyCardCPSSInGameSaveGetInfo_NewAPI(ByVal TradeSeq As String, ByVal CardPoint As Integer, ByVal oProjNo As String, ByVal CardKind As Integer) As MyCardCPSSInGameSaveGetInfo_NewAPIRE
        ConnStr(15)
        ' 建立命令(預存程序名,使用連線)
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_InGameSave_GetInfo_A03", Con)
        ' 設置命令型態為預存程序 
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        ' 增加參數欄位

        Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20).Value = TradeSeq
        Cmd.Parameters.Add("@CardPoint", SqlDbType.Int).Value = CardPoint
        Cmd.Parameters.Add("@oProjNo", SqlDbType.VarChar, 20).Value = oProjNo
        Cmd.Parameters.Add("@CardKind", SqlDbType.Int).Value = CardKind
        Dim GamePrice As SqlClient.SqlParameter = Cmd.Parameters.Add("@GamePrice", SqlDbType.Money)
        GamePrice.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output

        Dim CrrnyShotName As SqlClient.SqlParameter = Cmd.Parameters.Add("@CrrnyShotName", SqlDbType.VarChar, 6)
        CrrnyShotName.Direction = ParameterDirection.Output

        Dim ProductId As SqlClient.SqlParameter = Cmd.Parameters.Add("@ProductId", SqlDbType.VarChar, 50)
        ProductId.Direction = ParameterDirection.Output

        Dim ReturnValue As New MyCardCPSSInGameSaveGetInfo_NewAPIRE
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = ReturnMsg.Value
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.ReturnGamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)
            ReturnValue.ReturnCrrnyShotName = IIf(IsDBNull(CrrnyShotName.Value), "", CrrnyShotName.Value)
            'ReturnValue.ReturnDiscount = IIf(IsDBNull(Discount.Value), 0, Discount.Value)
            ReturnValue.ProductId = IIf(IsDBNull(ProductId.Value), "", ProductId.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|MyCardCPSSInGameSaveGetInfo|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2002"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.ReturnGamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)
            ReturnValue.ReturnCrrnyShotName = IIf(IsDBNull(CrrnyShotName.Value), "", CrrnyShotName.Value)
            'ReturnValue.ReturnDiscount = IIf(IsDBNull(Discount.Value), 0, Discount.Value)
            ReturnValue.ProductId = IIf(IsDBNull(ProductId.Value), "", ProductId.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function

    ''' <summary>
    ''' 2013-9-9
    ''' FB新API 扣點
    ''' </summary>
    ''' <param name="TradeSeq"></param>
    ''' <returns>SavePoint:儲值兌換點數;ReturnMsgNo;ReturnMsg;ErrorCode;LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function CPSS_MemberCost_RequestSaveGame_NewAPI(ByVal TradeSeq As String) As CPSS_MemberCost_RequestSaveGame_NewAPIResult
        ConnStr(15)
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_MemberCost_RequestSaveGame_A03", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20).Value = TradeSeq

        Dim SavePoint As SqlClient.SqlParameter = Cmd.Parameters.Add("@SavePoint", SqlDbType.Int)
        SavePoint.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output
        Dim Discount As SqlClient.SqlParameter = Cmd.Parameters.Add("@Discount", SqlDbType.VarChar, 7)
        Discount.Direction = ParameterDirection.Output
        Dim ProductId As SqlClient.SqlParameter = Cmd.Parameters.Add("@ProductId", SqlDbType.VarChar, 50)
        ProductId.Direction = ParameterDirection.Output
        Dim ReturnValue As New CPSS_MemberCost_RequestSaveGame_NewAPIResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.SavePoint = IIf(IsDBNull(SavePoint.Value), 0, SavePoint.Value)
            ReturnValue.CostPoint = IIf(IsDBNull(Discount.Value), 0, Discount.Value)
            ReturnValue.ProductId = IIf(IsDBNull(ProductId.Value), "", ProductId.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), TradeSeq & "|999|CPSS_MemberCost_RequestSaveGame_NewAPI|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2012"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.SavePoint = IIf(IsDBNull(SavePoint.Value), 0, SavePoint.Value)
            ReturnValue.CostPoint = IIf(IsDBNull(Discount.Value), 0, Discount.Value)
            ReturnValue.ProductId = IIf(IsDBNull(ProductId.Value), "", ProductId.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function

    ''' <summary>
    ''' 2013-9-9
    ''' FB新API Billing
    ''' Billing儲值成功，但要改CPSaveService的table狀態
    ''' </summary>
    ''' <param name="BillingAuthCode">billing驗證碼</param>
    ''' <returns>TradeSeq:儲值交易序號;TradePrice:價格;GamePrice:fb點數;ReturnMsgNo;ReturnMsg;ErrorCode;LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function CPSS_BillingSave_RequestSaveGame_NewAPI(ByVal BillingAuthCode As String) As CPSS_BillingSave_RequestSaveGame_NewAPIResult
        ConnStr(15)
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_BillingSave_RequestSaveGame_A03", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@BillingAuthCode", SqlDbType.VarChar, 100).Value = BillingAuthCode '--billing驗證碼

        Dim TradeSeq As SqlClient.SqlParameter = Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20) '--儲值交易序號
        TradeSeq.Direction = ParameterDirection.Output
        Dim TradePrice As SqlClient.SqlParameter = Cmd.Parameters.Add("@TradePrice", SqlDbType.Int)
        TradePrice.Direction = ParameterDirection.Output
        Dim GamePrice As SqlClient.SqlParameter = Cmd.Parameters.Add("@GamePrice", SqlDbType.Int)
        GamePrice.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output
        Dim Discount As SqlClient.SqlParameter = Cmd.Parameters.Add("@Discount", SqlDbType.VarChar, 7)
        Discount.Direction = ParameterDirection.Output
        Dim ProductId As SqlClient.SqlParameter = Cmd.Parameters.Add("@ProductId", SqlDbType.VarChar, 50)
        ProductId.Direction = ParameterDirection.Output
        Dim CrrnyShotName As SqlClient.SqlParameter = Cmd.Parameters.Add("@CrrnyShotName", SqlDbType.VarChar, 6)
        CrrnyShotName.Direction = ParameterDirection.Output
        Dim ReturnValue As New CPSS_BillingSave_RequestSaveGame_NewAPIResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
            ReturnValue.ReturnTradeSeq = IIf(IsDBNull(TradeSeq.Value), "", TradeSeq.Value)
            ReturnValue.TradePrice = IIf(IsDBNull(TradePrice.Value), 0, TradePrice.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.GamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)
            ReturnValue.CostPoint = IIf(IsDBNull(Discount.Value), 0, Discount.Value)
            ReturnValue.ProductId = IIf(IsDBNull(ProductId.Value), "", ProductId.Value)
            ReturnValue.CrrnyShotName = IIf(IsDBNull(CrrnyShotName.Value), "", CrrnyShotName.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|CPSS_BillingSave_RequestSaveGame|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2007"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.GamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)
            ReturnValue.CostPoint = IIf(IsDBNull(Discount.Value), 0, Discount.Value)
            ReturnValue.ProductId = IIf(IsDBNull(ProductId.Value), "", ProductId.Value)
            ReturnValue.CrrnyShotName = IIf(IsDBNull(CrrnyShotName.Value), "", CrrnyShotName.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function

    ''' <summary>
    ''' 2013-9-9 New API
    '''   MyCardInGame
    ''' </summary>
    ''' <param name="GameFacID">遊戲廠商代號,格式String</param>
    ''' <param name="MyCardCardId">MyCard卡號,格式String</param>
    ''' <param name="GameCustId">遊戲帳號,格式String</param>
    ''' <param name="CreateIP">新增者IP,格式String</param>
    ''' <returns>ReturnNo 回傳代碼 ,ReturnMsg 回傳訊息,MyCardPoint 回傳廠商交易序號,ErrorCode 回傳ErrorCode,LogSn 回傳LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function MyCardCPSSInGameSaveTradeListInsert_NewAPI(ByVal GameFacID As String, ByVal MyCardCardId As String, ByVal GameCustId As String, ByVal CreateIP As String, ByVal TradeType As Integer) As NewAPIReturnValueRE
        ConnStr(15)
        ' 建立命令(預存程序名,使用連線) 
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_InGameSaveTradeList_Insert_A02", Con)
        ' 設置命令型態為預存程序 
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        ' 增加參數欄位

        Cmd.Parameters.Add("@GameFacID", SqlDbType.VarChar, 30).Value = GameFacID
        Cmd.Parameters.Add("@MyCardCardId", SqlDbType.VarChar, 20).Value = MyCardCardId
        Cmd.Parameters.Add("@GameCustId", SqlDbType.VarChar, 100).Value = GameCustId
        Cmd.Parameters.Add("@CreateIP", SqlDbType.VarChar, 30).Value = CreateIP
        Cmd.Parameters.Add("@TradeType ", SqlDbType.TinyInt).Value = TradeType

        Dim TradeSeq As SqlClient.SqlParameter = Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20)
        TradeSeq.Direction = ParameterDirection.Output
        Dim GamePrice As SqlClient.SqlParameter = Cmd.Parameters.Add("@GamePrice", SqlDbType.Money)
        GamePrice.Direction = ParameterDirection.Output
        Dim CrrnyShotName As SqlClient.SqlParameter = Cmd.Parameters.Add("@CrrnyShotName", SqlDbType.VarChar, 6)
        CrrnyShotName.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 1024)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output


        Dim ReturnValue As New NewAPIReturnValueRE
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = ReturnMsg.Value
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnTradeSeq = IIf(IsDBNull(TradeSeq.Value), "", TradeSeq.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.ReturnGamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)
            ReturnValue.ReturnCrrnyShotName = IIf(IsDBNull(CrrnyShotName.Value), "", CrrnyShotName.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|MyCardCPSSInGameSaveTradeListInsert_NewAPI|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2001"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.ReturnGamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)
            ReturnValue.ReturnCrrnyShotName = IIf(IsDBNull(CrrnyShotName.Value), "", CrrnyShotName.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function
    ''' <summary>
    ''' 2013-9-9 New API
    ''' 會員扣點初始
    ''' </summary>
    ''' <param name="GameFacId">廠商ID</param>
    ''' <param name="GameFacSerId">廠商服務ID</param>
    ''' <param name="OneTimePassport">登入機碼</param>
    ''' <param name="GameCustId">會員帳號</param>
    ''' <param name="GameUserId">UID帳號</param>
    ''' <param name="CostPoint">扣點點數</param>
    ''' <param name="CreateIP">IP</param>
    ''' <returns>ReturnTradeSeq:扣點交易序號;SavePoint:轉換儲值點數;ReturnMsgNo;ReturnMsg;ErrorCode;LogSn</returns>
    ''' <remarks></remarks>
    ''' 20110216 多加一個記錄uid的欄位 詹旻融
    <WebMethod()> _
    Public Function CPSS_MemberCostSaveTradeList_Insert_NewAPI(ByVal GameFacId As String, ByVal GameFacSerId As String, ByVal OneTimePassport As String, ByVal GameCustId As String, ByVal CostPoint As Integer, ByVal CreateIP As String, ByVal GameUserId As String) As CPSS_MemberCostSaveTradeList_InsertResult
        ConnStr(15)
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_MemberCostSaveTradeList_Insert_A02", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@GameFacId", SqlDbType.NVarChar, 50).Value = GameFacId
        Cmd.Parameters.Add("@GameFacSerId", SqlDbType.NVarChar, 50).Value = GameFacSerId
        Cmd.Parameters.Add("@OneTimePassport", SqlDbType.VarChar, 50).Value = OneTimePassport
        Cmd.Parameters.Add("@GameCustId", SqlDbType.VarChar, 100).Value = GameCustId
        Cmd.Parameters.Add("@GameUserId", SqlDbType.VarChar, 100).Value = GameUserId
        Cmd.Parameters.Add("@CostPoint", SqlDbType.Int).Value = CostPoint
        Cmd.Parameters.Add("@CreateIP", SqlDbType.VarChar, 30).Value = CreateIP

        Dim TradeSeq As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnTradeSeq", SqlDbType.VarChar, 20)
        TradeSeq.Direction = ParameterDirection.Output
        Dim SavePoint As SqlClient.SqlParameter = Cmd.Parameters.Add("@SavePoint", SqlDbType.Int)
        SavePoint.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output
        Dim ReturnValue As New CPSS_MemberCostSaveTradeList_InsertResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
            ReturnValue.TradeSeq = IIf(IsDBNull(TradeSeq.Value), "", TradeSeq.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.SavePoint = IIf(IsDBNull(SavePoint.Value), 0, SavePoint.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|CPSS_MemberCostSaveTradeList_Insert_NewAPI|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2010"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.TradeSeq = ""
            ReturnValue.SavePoint = 0
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function
    ''' <summary>
    ''' 2013-9-9 New API
    ''' Billing初始
    ''' </summary>
    ''' <param name="ProductServiceId">產品代碼</param>
    ''' <param name="TradePrice">產品金額</param>
    ''' <param name="UserId">使用者帳號</param>
    ''' <param name="UserIp">使用者IP</param>
    ''' <param name="TradeType">1:FACEBOOK 2會員加點IP</param>
    ''' <returns>TradeSeq:儲值交易序號;GamePrice:facebook點數;ReturnMsgNo;ReturnMsg;ErrorCode;LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function CPSS_BillingSave_GetProductInfo_NewAPI(ByVal ProductServiceId As String, ByVal TradePrice As Integer, ByVal UserId As String, ByVal UserIp As String, ByVal TradeType As Integer) As CPSS_BillingSave_GetProductInfoResult
        ConnStr(15)
        '20110614 qq 修改連新的SP
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_BillingSave_GetProductInfo_A02", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@ProductServiceId", SqlDbType.VarChar, 10).Value = ProductServiceId '--產品代碼
        Cmd.Parameters.Add("@TradePrice", SqlDbType.Int).Value = TradePrice '--產品金額
        Cmd.Parameters.Add("@UserId", SqlDbType.VarChar, 100).Value = UserId ' --使用者帳號
        Cmd.Parameters.Add("@UserIp", SqlDbType.VarChar, 25).Value = UserIp '--使用者IP
        Cmd.Parameters.Add("@TradeType", SqlDbType.Int).Value = TradeType '--1:FACEBOOK 2會員加點IP

        Dim TradeSeq As SqlClient.SqlParameter = Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 20) '--儲值交易序號
        TradeSeq.Direction = ParameterDirection.Output
        Dim GamePrice As SqlClient.SqlParameter = Cmd.Parameters.Add("@GamePrice", SqlDbType.Int)
        GamePrice.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output

        Dim ReturnValue As New CPSS_BillingSave_GetProductInfoResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
            ReturnValue.ReturnTradeSeq = IIf(IsDBNull(TradeSeq.Value), "", TradeSeq.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.GamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|CPSS_BillingSave_GetProductInfo_NewAPI|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2005"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.GamePrice = IIf(IsDBNull(GamePrice.Value), 0, GamePrice.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function

    ''' <summary>
    ''' MyCard_CPSaveService.dbo.CPSS_MemberCostSaveTradeList_Insert_A02(已不使用)
    ''' 2014/07/17 會員扣點初始交易
    ''' </summary>
    ''' <param name="GameFacId">廠商ID</param>
    ''' <param name="GameFacSerId">廠商服務ID</param>
    ''' <param name="OneTimePassport">登入機碼</param>
    ''' <param name="GameCustId">會員帳號</param>
    ''' <param name="GameUserId">UID帳號</param>
    ''' <param name="CostPoint">扣點點數</param>
    ''' <param name="CreateIp">IP</param>
    ''' <returns>ReturnTradeSeq :扣點交易序號;SavePoint:轉換儲值點數;ServiceTypeSn:服務typesn;ReturnMsgNo;ReturnMsg;ErrorCode;LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function CPSS_MemberCostSaveTradeList_Insert_NewAPI03(ByVal GameFacId As String, ByVal GameFacSerId As String, ByVal OneTimePassport As String, ByVal GameCustId As String, ByVal CostPoint As Integer, ByVal CreateIP As String, ByVal GameUserId As String) As CPSS_MemberCostSaveTradeList_InsertResult
        ConnStr(15)
        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_MemberCostSaveTradeList_Insert_A03", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@GameFacId", SqlDbType.NVarChar, 50).Value = GameFacId
        Cmd.Parameters.Add("@GameFacSerId", SqlDbType.NVarChar, 50).Value = GameFacSerId
        Cmd.Parameters.Add("@OneTimePassport", SqlDbType.VarChar, 50).Value = OneTimePassport
        Cmd.Parameters.Add("@GameCustId", SqlDbType.VarChar, 100).Value = GameCustId
        Cmd.Parameters.Add("@GameUserId", SqlDbType.VarChar, 100).Value = GameUserId
        Cmd.Parameters.Add("@CostPoint", SqlDbType.Int).Value = CostPoint
        Cmd.Parameters.Add("@CreateIP", SqlDbType.VarChar, 25).Value = CreateIP

        Dim TradeSeq As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnTradeSeq", SqlDbType.VarChar, 20)
        TradeSeq.Direction = ParameterDirection.Output
        Dim SavePoint As SqlClient.SqlParameter = Cmd.Parameters.Add("@SavePoint", SqlDbType.Int)
        SavePoint.Direction = ParameterDirection.Output
        Dim ServiceTypeSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@ServiceTypeSn", SqlDbType.Int)
        ServiceTypeSn.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 1024)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output
        Dim ReturnValue As New CPSS_MemberCostSaveTradeList_InsertResult
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
            ReturnValue.TradeSeq = IIf(IsDBNull(TradeSeq.Value), "", TradeSeq.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.SavePoint = IIf(IsDBNull(SavePoint.Value), 0, SavePoint.Value)
            ReturnValue.ServiceTypeSn = IIf(IsDBNull(ServiceTypeSn.Value), 0, ServiceTypeSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|CPSS_MemberCostSaveTradeList_Insert_NewAPI03|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2010"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.TradeSeq = ""
            ReturnValue.SavePoint = 0
            ReturnValue.ServiceTypeSn = 0
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function



    ''' <summary>
    '''   會員EMAIL
    ''' </summary>
    ''' <param name="MyCardCustId">會員帳號</param>
    ''' <returns> ReturnNo 回傳代碼 ,ReturnMsg 回傳訊息,ErrorCode 回傳ErrorCode,LogSn 回傳LogSn</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function MyCard_MemberQueryUserMail_A01(ByVal MyCardCustId As String) As ReturnValueMe
        ConnStr(19)
        ' 建立命令(預存程序名,使用連線) 
        Dim Cmd As New SqlClient.SqlCommand("MyCard_Member.dbo.MyCard_MemberQueryUserMail_A01", Con)
        ' 設置命令型態為預存程序 
        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300
        ' 增加參數欄位

        Cmd.Parameters.Add("@MyCardCustId", SqlDbType.VarChar, 50).Value = MyCardCustId

        Dim User_Mail As SqlClient.SqlParameter = Cmd.Parameters.Add("@User_Mail", SqlDbType.VarChar, 50)
        User_Mail.Direction = ParameterDirection.Output
        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnValueMe
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = ReturnMsg.Value
            ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.ReturnUserMail = IIf(IsDBNull(User_Mail.Value), "", User_Mail.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardServiceForInGame.asmx"), "999|MyCard_MemberQueryUserMail_A01|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS2014"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            ReturnValue.ReturnUserMail = User_Mail.Value
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function


End Class


Public Class FBSimpleMemberRegisterPWDUpdateResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer
    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
    End Sub
End Class

Public Class ViewFBMemberSnResult
    Public Sn As String
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Sub New()
        Sn = ""
        ReturnMsgNo = 0
        ReturnMsg = ""
    End Sub
End Class

Public Class FBUserBindingInsertResult
    Public uPOINTS As String
    Public uUSER_NAME As String
    Public MyCardOTP As String
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer
    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
    End Sub
End Class

Public Class FBSimpleMemberUpdateResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer
    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
    End Sub
End Class

Public Class FBSimpleMemberRegisterResult
    Public Sn As String
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer
    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
    End Sub
End Class

Public Class FBMemberLoginCheckResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public MyCardOTP As String
    Public MemberType As Integer
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer
    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
    End Sub
End Class

Public Class FBCheckUIDResult
    Public uMyCardCustId As String
    Public uPOINTS As Integer
    Public uUSER_NAME As String
    Public MyCardOTP As String
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer
    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
    End Sub
End Class

Public Class ReturnValueMe
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnDataSet As DataSet
    Public ReturnUserMail As String
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer

    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
        ReturnDataSet = New DataSet
        ReturnUserMail = ""
        ReturnErrorCode = ""
        ReturnLogSn = 0
    End Sub
End Class

Public Class ReturnValueRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnDataSet As DataSet
    Public ReturnTradeSeq As String
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer
    Public ReturnGamePrice As Integer
    Public ReturnInnerOTPCode As String

    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
        ReturnDataSet = New DataSet
        ReturnTradeSeq = ""
        ReturnErrorCode = ""
        ReturnLogSn = 0
        ReturnGamePrice = 0
        ReturnInnerOTPCode = ""

    End Sub
End Class
Public Class NewAPIReturnValueRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnDataSet As DataSet
    Public ReturnTradeSeq As String
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer
    Public ReturnGamePrice As Decimal
    Public ReturnCrrnyShotName As String

    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
        ReturnDataSet = New DataSet
        ReturnTradeSeq = ""
        ReturnErrorCode = ""
        ReturnLogSn = 0
        ReturnGamePrice = 0
        ReturnCrrnyShotName = ""

    End Sub
End Class
Public Class MyCardCPSSInGameSaveGetInfoRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer
    Public ReturnGamePrice As Integer
    Public ReturnCrrnyShotName As String
    Public ReturnDiscount As String
    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
        
        ReturnErrorCode = ""
        ReturnLogSn = 0
        ReturnGamePrice = 0

        ReturnCrrnyShotName = ""
        ReturnDiscount = ""
    End Sub
End Class
Public Class MyCardCPSSInGameSaveGetInfo_NewAPIRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer
    Public ReturnGamePrice As Decimal
    Public ReturnCrrnyShotName As String
    Public ReturnDiscount As String
    Public ProductId As String

    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""

        ReturnErrorCode = ""
        ReturnLogSn = 0
        ReturnGamePrice = 0

        ReturnCrrnyShotName = ""
        ReturnDiscount = ""
        ProductId = ""
    End Sub
End Class
Public Class CPSS_BillingSave_GetProductInfoResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String = ""
    Public ReturnTradeSeq As String = ""
    Public ReturnErrorCode As String = ""
    Public ReturnLogSn As Integer
    Public GamePrice As Integer
End Class
Public Class CPSS_BillingSave_GetBillingAuthCodeResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String = ""
    Public ReturnErrorCode As String = ""
    Public ReturnLogSn As Integer
End Class
Public Class CPSS_BillingSave_RequestSaveGameResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String = ""
    Public TradePrice As Integer
    Public ReturnTradeSeq As String = ""
    Public ReturnErrorCode As String = ""
    Public ReturnLogSn As Integer
    Public GamePrice As Integer
    Public CostPoint As String = ""
End Class
Public Class CPSS_BillingSave_RequestSaveGame_NewAPIResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String = ""
    Public TradePrice As Integer
    Public ReturnTradeSeq As String = ""
    Public ReturnErrorCode As String = ""
    Public ReturnLogSn As Integer
    Public GamePrice As Integer
    Public CostPoint As String = ""
    Public ProductId As String = ""
    Public CrrnyShotName As String = ""
End Class
Public Class CPSS_BillingSave_ResponseSaveGameResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String = ""
    Public BillingAuthCode As String = ""
    Public GameCustId As String = ""
    Public ReturnErrorCode As String = ""
    Public ReturnLogSn As Integer
End Class
Public Class CPSS_BillingSave_CheckSaveTradeResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String = ""
    Public ReturnErrorCode As String = ""
    Public ReturnLogSn As Integer
    Public TradePrice As Integer
End Class
Public Class CPSS_MemberCostSaveTradeList_InsertResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String = ""
    Public ReturnErrorCode As String = ""
    Public ReturnLogSn As Integer
    Public TradeSeq As String = ""
    Public SavePoint As Integer
    Public ServiceTypeSn As Integer
End Class
Public Class CPSS_MemberCostAuthCode_InsertResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String = ""
    Public ReturnErrorCode As String = ""
    Public ReturnLogSn As Integer
End Class
Public Class CPSS_MemberCost_RequestSaveGameResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String = ""
    Public ReturnErrorCode As String = ""
    Public ReturnLogSn As Integer
    Public SavePoint As Integer
    Public CostPoint As String = ""
End Class
Public Class CPSS_MemberCost_RequestSaveGame_NewAPIResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String = ""
    Public ReturnErrorCode As String = ""
    Public ReturnLogSn As Integer
    Public SavePoint As Integer
    Public CostPoint As String = ""
    Public ProductId As String = ""
End Class
Public Class CPSS_MemberCost_ResponseSaveGameResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String = ""
    Public ReturnErrorCode As String = ""
    Public ReturnLogSn As Integer
End Class
Public Class MyCard_MemberService_GetElectronicWallet_FactoryServiceResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String = ""
    Public ReturnDS As New DataSet
End Class
Public Class MemberCheckStatusResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String = ""
    Public ReturnErrorCode As String = ""
    Public ReturnLogSn As Integer
End Class
Public Class MyCard_MemberIdTradeLimitForFB_A01Result
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnErrorCode As String = ""
    Public ReturnLogSn As Integer
End Class