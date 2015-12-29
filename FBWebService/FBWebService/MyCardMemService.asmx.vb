Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Net

<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class MyCardMemService
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

    <WebMethod()> _
    Public Function MyCardMemberServiceAuth(ByVal FactoryId As String, ByVal FactoryServiceId As String, ByVal FactoryIp As String, ByVal FactorySeq As String, ByVal PointPayment As Integer, ByVal BonusPayment As Integer, ByVal FactoryReturnUrl As String) As MyCardMemberServiceAuthResult
        ConnStr(20)
        Dim command As New SqlClient.SqlCommand
        command.Connection = Con
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "MyCard_MemberService.dbo.MyCard_MemberService_Auth_A01"

        command.Parameters.Add("@FactoryId", SqlDbType.NVarChar, 50).Value = FactoryId
        command.Parameters.Add("@FactoryServiceId", SqlDbType.NVarChar, 50).Value = FactoryServiceId
        command.Parameters.Add("@FactoryIp", SqlDbType.NVarChar, 50).Value = FactoryIp
        command.Parameters.Add("@FactorySeq", SqlDbType.NVarChar, 50).Value = FactorySeq
        command.Parameters.Add("@PointPayment", SqlDbType.Int).Value = PointPayment
        command.Parameters.Add("@BonusPayment", SqlDbType.Int).Value = BonusPayment
        command.Parameters.Add("@FactoryReturnUrl", SqlDbType.NVarChar, 250).Value = FactoryReturnUrl

        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 1024)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnTradeSeq As SqlClient.SqlParameter = command.Parameters.Add("@ReturnTradeSeq", SqlDbType.NVarChar, 50)
        ReturnTradeSeq.Direction = ParameterDirection.Output

        Dim ReturnAuthCode As SqlClient.SqlParameter = command.Parameters.Add("@ReturnAuthCode", SqlDbType.NVarChar, 50)
        ReturnAuthCode.Direction = ParameterDirection.Output

        Dim ErrorCode As SqlClient.SqlParameter = command.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = command.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output

        Dim ReturnResult As New MyCardMemberServiceAuthResult
        Dim InputValue As String = FactoryId & "," & FactoryServiceId & "," & FactoryIp & "," & FactorySeq & "," & PointPayment & "," & BonusPayment & "," & FactoryReturnUrl
        Try
            Con.Open()
            command.ExecuteNonQuery()
            ReturnResult.ReturnMsgNo = ReturnMsgNo.Value
            ReturnResult.ReturnMsg = ReturnMsg.Value
            ReturnResult.ReturnTradeSeq = ReturnTradeSeq.Value
            ReturnResult.ReturnAuthCode = ReturnAuthCode.Value
            ReturnResult.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnResult.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "999|MyCardMemberServiceAuth|" & InputValue & "|" & ex.Message, "127.0.0.3")
            ReturnResult.ReturnMsgNo = "-999"
            ReturnResult.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnResult.ReturnTradeSeq = ""
            ReturnResult.ReturnAuthCode = ""
            ReturnResult.ReturnErrorCode = "FBWS3001"
            ReturnResult.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Finally
            Con.Close()
        End Try

        Return ReturnResult
    End Function

    <WebMethod()> _
    Public Function MyCardMemberServiceAuthWithOtp(ByVal FactoryId As String, ByVal FactoryServiceId As String, ByVal MyCardCustId As String, ByVal OneTimePassword As String, ByVal FactoryIp As String, ByVal FactorySeq As String, ByVal PointPayment As Integer, ByVal BonusPayment As Integer, ByVal FactoryReturnUrl As String) As MyCardMemberServiceAuthResult
        ConnStr(20)

        Dim ReturnResult As New MyCardMemberServiceAuthResult

        Dim command As New SqlClient.SqlCommand
        command.Connection = Con
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "MyCard_MemberService.dbo.MyCard_MemberService_AuthWithOtp_A01"

        command.Parameters.Add("@FactoryId", SqlDbType.NVarChar, 50).Value = FactoryId
        command.Parameters.Add("@FactoryServiceId", SqlDbType.NVarChar, 50).Value = FactoryServiceId
        command.Parameters.Add("@MyCardCustId", SqlDbType.NVarChar, 50).Value = MyCardCustId
        command.Parameters.Add("@OneTimePassword", SqlDbType.NVarChar, 50).Value = OneTimePassword
        command.Parameters.Add("@FactoryIp", SqlDbType.NVarChar, 50).Value = FactoryIp
        command.Parameters.Add("@FactorySeq", SqlDbType.NVarChar, 50).Value = FactorySeq
        command.Parameters.Add("@PointPayment", SqlDbType.Int).Value = PointPayment
        command.Parameters.Add("@BonusPayment", SqlDbType.Int).Value = BonusPayment
        command.Parameters.Add("@FactoryReturnUrl", SqlDbType.NVarChar, 250).Value = FactoryReturnUrl

        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnTradeSeq As SqlClient.SqlParameter = command.Parameters.Add("@ReturnTradeSeq", SqlDbType.NVarChar, 50)
        ReturnTradeSeq.Direction = ParameterDirection.Output

        Dim ReturnAuthCode As SqlClient.SqlParameter = command.Parameters.Add("@ReturnAuthCode", SqlDbType.NVarChar, 50)
        ReturnAuthCode.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 1024)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = command.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = command.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output
        Dim InputValue As String = FactoryId & "," & FactoryServiceId & "," & MyCardCustId & "," & OneTimePassword & "," & FactoryIp & "," & FactorySeq & "," & PointPayment & "," & BonusPayment & "," & FactoryReturnUrl
        Try
            Con.Open()
            command.ExecuteNonQuery()
            ReturnResult.ReturnMsgNo = ReturnMsgNo.Value
            ReturnResult.ReturnMsg = ReturnMsg.Value
            ReturnResult.ReturnTradeSeq = ReturnTradeSeq.Value
            ReturnResult.ReturnAuthCode = ReturnAuthCode.Value
            ReturnResult.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnResult.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "999|MyCardMemberServiceAuthWithOtp|" & InputValue & "|" & ex.Message, "127.0.0.3")
            ReturnResult.ReturnMsgNo = "-999"
            ReturnResult.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnResult.ReturnTradeSeq = ""
            ReturnResult.ReturnAuthCode = ""
            ReturnResult.ReturnErrorCode = "FBWS3002"
            ReturnResult.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Finally
            Con.Close()
        End Try

        Return ReturnResult
    End Function


    <WebMethod()> _
    Public Function MyCardMemberServiceMemberAddListGetUserInfo(ByVal AuthCode As String, ByVal OneTimePassword As String) As MyCardMemberServiceGetUserInfoResult
        ConnStr(20)
        Dim command As New SqlClient.SqlCommand
        command.Connection = Con
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "MyCard_MemberService.dbo.MyCard_MemberService_AddListGetUserInfo_A01"

        command.Parameters.Add("@AuthCode", SqlDbType.NVarChar, 50).Value = AuthCode
        command.Parameters.Add("@OneTimePassword", SqlDbType.NVarChar, 50).Value = OneTimePassword

        Dim MyCardCustId As SqlClient.SqlParameter = command.Parameters.Add("@MyCardCustId", SqlDbType.NVarChar, 50)
        MyCardCustId.Direction = ParameterDirection.Output

        Dim MyCardPointServiceType As SqlClient.SqlParameter = command.Parameters.Add("@MyCardPointServiceType", SqlDbType.Int)
        MyCardPointServiceType.Direction = ParameterDirection.Output

        Dim MyCardServiceName As SqlClient.SqlParameter = command.Parameters.Add("@MyCardServiceName", SqlDbType.NVarChar, 50)
        MyCardServiceName.Direction = ParameterDirection.Output

        Dim MyCardServiceTradeListSn As SqlClient.SqlParameter = command.Parameters.Add("@MyCardServiceTradeListSn", SqlDbType.Int)
        MyCardServiceTradeListSn.Direction = ParameterDirection.Output

        Dim MyCardPoint As SqlClient.SqlParameter = command.Parameters.Add("@MyCardPoint", SqlDbType.Int)
        MyCardPoint.Direction = ParameterDirection.Output

        Dim MyCardBonus As SqlClient.SqlParameter = command.Parameters.Add("@MyCardBonus", SqlDbType.Int)
        MyCardBonus.Direction = ParameterDirection.Output

        Dim FacTradeSeq As SqlClient.SqlParameter = command.Parameters.Add("@FacTradeSeq", SqlDbType.NVarChar, 50)
        FacTradeSeq.Direction = ParameterDirection.Output

        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 1024)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ErrorCode As SqlClient.SqlParameter = command.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = command.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output

        Dim ReturnResult As New MyCardMemberServiceGetUserInfoResult

        Try
            Con.Open()
            command.ExecuteNonQuery()
            ReturnResult.ReturnMsgNo = ReturnMsgNo.Value
            ReturnResult.ReturnMyCardCustId = MyCardCustId.Value
            ReturnResult.ReturnMyCardPointServiceType = MyCardPointServiceType.Value
            ReturnResult.ReturnMyCardServiceName = MyCardServiceName.Value
            ReturnResult.MyCardServiceTradeListSn = MyCardServiceTradeListSn.Value
            ReturnResult.ReturnMyCardPoint = MyCardPoint.Value
            ReturnResult.ReturnMyCardBonus = MyCardBonus.Value
            ReturnResult.ReturnFacTradeSeq = FacTradeSeq.Value
            ReturnResult.ReturnMsg = ReturnMsg.Value
            ReturnResult.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnResult.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "999|MyCardMemberServiceMemberAddListGetUserInfo|" & AuthCode & "," & OneTimePassword & "|" & ex.Message, "127.0.0.3")
            ReturnResult.ReturnMsgNo = "-999"
            ReturnResult.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnResult.ReturnMyCardCustId = ""
            ReturnResult.ReturnMyCardPointServiceType = ""
            ReturnResult.ReturnMyCardServiceName = ""
            ReturnResult.MyCardServiceTradeListSn = ""
            ReturnResult.ReturnMyCardPoint = ""
            ReturnResult.ReturnMyCardBonus = ""
            ReturnResult.ReturnFacTradeSeq = ""
            ReturnResult.ReturnErrorCode = "FBWS3003"
            ReturnResult.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Finally
            Con.Close()
        End Try

        Return ReturnResult
    End Function

    <WebMethod()> _
    Public Function MyCardCustPointAddEx(ByVal MyCardCustId As String, ByVal MyCardServiceTypeSn As Integer, ByVal MyCardPoint As Integer, ByVal MyCardBonusPoint As Integer, ByVal MyCardServiceTradeListSn As Integer, ByVal MyCardNote As String, ByVal MyCardServiceId As String, ByVal MyCardTradeNo As String, ByVal MyCardIdentifyData As String) As MyCardPointAddExResult
        ConnStr(19)
        Dim command As New SqlClient.SqlCommand
        command.Connection = Con
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "MyCard_Member.dbo.MyCard_MemberAddPointEx_A01"

        command.Parameters.Add("@MyCardCustId", SqlDbType.NVarChar, 50).Value = MyCardCustId
        command.Parameters.Add("@MyCardServiceTypeSn", SqlDbType.Int).Value = MyCardServiceTypeSn
        command.Parameters.Add("@MyCardPoint", SqlDbType.Int).Value = MyCardPoint
        command.Parameters.Add("@MyCardBonusPoint", SqlDbType.Int).Value = MyCardBonusPoint
        command.Parameters.Add("@MyCardServiceTradeListSn", SqlDbType.Int).Value = MyCardServiceTradeListSn
        command.Parameters.Add("@MyCardNote", SqlDbType.NVarChar, 50).Value = MyCardNote
        command.Parameters.Add("@MyCardServiceId", SqlDbType.NVarChar, 10).Value = MyCardServiceId
        command.Parameters.Add("@MyCardServiceTradeNo", SqlDbType.NVarChar, 32).Value = MyCardTradeNo
        command.Parameters.Add("@MyCardServiceIData", SqlDbType.NVarChar, 50).Value = MyCardIdentifyData

        Dim ReturnBonusTradeSeq As SqlClient.SqlParameter = command.Parameters.Add("@ReturnBonusTradeSeq", SqlDbType.NVarChar, 20)
        ReturnBonusTradeSeq.Direction = ParameterDirection.Output

        Dim ReturnPointTradeSeq As SqlClient.SqlParameter = command.Parameters.Add("@ReturnPointTradeSeq", SqlDbType.NVarChar, 20)
        ReturnPointTradeSeq.Direction = ParameterDirection.Output

        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 1024)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ErrorCode As SqlClient.SqlParameter = command.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = command.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output

        Dim ReturnResult As New MyCardPointAddExResult
        Dim InputValue As String = MyCardCustId & "," & MyCardServiceTypeSn & "," & MyCardPoint & "," & MyCardBonusPoint & "," & MyCardServiceTradeListSn & "," & MyCardNote & "," & MyCardServiceId & "," & MyCardTradeNo & "," & MyCardIdentifyData
        Try
            Con.Open()
            command.ExecuteNonQuery()

            If ReturnMsgNo.Value = 1 Then
                ReturnResult.ReturnResult = True
                ReturnResult.ReturnBonusTradeSeq = ReturnBonusTradeSeq.Value
                ReturnResult.ReturnPointTradeSeq = ReturnPointTradeSeq.Value
                ReturnResult.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
                ReturnResult.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            Else
                ReturnResult.ReturnMsgNo = ReturnMsgNo.Value
                ReturnResult.ReturnResult = False
                ReturnResult.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
                ReturnResult.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
            End If

            ReturnResult.ReturnMsg = ReturnMsg.Value
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "999|MyCardCustPointAddEx|" & InputValue & "|" & ex.Message, "127.0.0.3")
            ReturnResult.ReturnMsgNo = "-999"
            ReturnResult.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnResult.ReturnBonusTradeSeq = ""
            ReturnResult.ReturnPointTradeSeq = ""
            ReturnResult.ReturnErrorCode = "FBWS3004"
            ReturnResult.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Finally
            Con.Close()
        End Try

        Return ReturnResult
    End Function

    <WebMethod()> _
    Public Function MyCardMemberServiceMemberAddListRender(ByVal AuthCode As String, ByVal OneTimePassword As String) As MemberAddListRenderResult
        ConnStr(20)
        Dim command As New SqlClient.SqlCommand
        command.Connection = Con
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "MyCard_MemberService.dbo.MyCard_MemberService_AddListRender"

        command.Parameters.Add("@AuthCode", SqlDbType.NVarChar, 50).Value = AuthCode
        command.Parameters.Add("@OneTimePassword", SqlDbType.NVarChar, 50).Value = OneTimePassword

        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnResult As New MemberAddListRenderResult

        Try
            Con.Open()
            command.ExecuteNonQuery()
            ReturnResult.ReturnMsgNo = ReturnMsgNo.Value
            ReturnResult.ReturnMsg = ""
            ReturnResult.ReturnErrorCode = ""
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "999|MyCardMemberServiceMemberAddListRender|" & AuthCode & "," & OneTimePassword & "|" & ex.Message, "127.0.0.3")
            ReturnResult.ReturnMsgNo = -999
            ReturnResult.ReturnMsg = "系統發生錯誤" & ex.Message
            ReturnResult.ReturnErrorCode = "FBWS3005"
        Finally
            Con.Close()
        End Try

        Return ReturnResult
    End Function


    <WebMethod()> _
    Public Function MemberAddListRender(ByVal AuthCode As String, ByVal OneTimePassword As String, ByVal MyCardProjectNo As String) As MemberAddListRenderResult
        'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "MemberAddListRender|" & AuthCode & "|" & OneTimePassword & "|" & MyCardProjectNo, "127.0.0.3")
        Dim GetUserInfoResult As New MyCardMemberServiceGetUserInfoResult

        Dim MyReturnResult As New MemberAddListRenderResult

        GetUserInfoResult = MyCardMemberServiceMemberAddListGetUserInfo(AuthCode, OneTimePassword)

        If GetUserInfoResult.ReturnMsgNo <> 1 Then
            MyReturnResult.ReturnMsgNo = -902
            MyReturnResult.ReturnMsg = GetUserInfoResult.ReturnMsg
            MyReturnResult.ReturnErrorCode = GetUserInfoResult.ReturnErrorCode
            Return MyReturnResult
        End If

        Dim MyCardPointAddResult As New MyCardPointAddExResult
        'log
        MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "1.MyCardCustPointAddEx|" & MyCardProjectNo & "|" & GetUserInfoResult.ReturnMyCardCustId & "|" & GetUserInfoResult.ReturnMyCardPointServiceType & "|" & GetUserInfoResult.ReturnMyCardPoint & "|" & GetUserInfoResult.ReturnMyCardBonus & "|" & GetUserInfoResult.MyCardServiceTradeListSn & "|" & GetUserInfoResult.ReturnMyCardServiceName & "|" & GetUserInfoResult.ReturnFacTradeSeq & "|" & AuthCode, "127.0.0.3")
        '加點
        MyCardPointAddResult = MyCardCustPointAddEx(GetUserInfoResult.ReturnMyCardCustId, GetUserInfoResult.ReturnMyCardPointServiceType, GetUserInfoResult.ReturnMyCardPoint, GetUserInfoResult.ReturnMyCardBonus, GetUserInfoResult.MyCardServiceTradeListSn, GetUserInfoResult.ReturnMyCardServiceName, "MYC_PA", GetUserInfoResult.ReturnFacTradeSeq, AuthCode)

        If MyCardPointAddResult.ReturnResult = False Then
            MyReturnResult.ReturnMsgNo = -904
            MyReturnResult.ReturnMsg = MyCardPointAddResult.ReturnMsg
            MyReturnResult.ReturnErrorCode = MyCardPointAddResult.ReturnErrorCode
            Return MyReturnResult
        End If


        MyReturnResult = MyCardMemberServiceMemberAddListRender(AuthCode, OneTimePassword)

        If MyCardProjectNo.ToUpper = "A0000" Or MyCardProjectNo = "" Then

            Return MyReturnResult
        End If

        '20120308 qq 新增帶入活動代號查出加點服務 MyCard_MemberService_CheckActivityNum
        Dim CheckActivityNumResult As New MyCardMemberCheckActivityNumResult
        Try
            CheckActivityNumResult = MyCard_MemberService_CheckActivityNum(MyCardProjectNo)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "MemberAddListRender|-905|" & MyCardProjectNo & "|" & ex.Message, "127.0.0.3")
            MyReturnResult.ReturnMsgNo = 1
            MyReturnResult.ReturnMsg = ex.Message
            Return MyReturnResult
        End Try

        If CheckActivityNumResult.ReturnMsgNo <> 1 Then
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "MemberAddListRender|-906|" & MyCardProjectNo & "|" & CheckActivityNumResult.ReturnMsgNo & "|" & CheckActivityNumResult.ReturnMsg & "|" & CheckActivityNumResult.ReturnErrorCode & "|" & CheckActivityNumResult.ReturnLogSn, "127.0.0.3")
            MyReturnResult.ReturnMsgNo = 1
            MyReturnResult.ReturnMsg = CheckActivityNumResult.ReturnMsgNo & "|" & CheckActivityNumResult.ReturnMsg & "|" & CheckActivityNumResult.ReturnErrorCode & "|" & CheckActivityNumResult.ReturnLogSn
            Return MyReturnResult
        End If

        'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "MyCard_MemberService_CheckActivityNum|" & CheckActivityNumResult.ReturnFactoryId & "|" & CheckActivityNumResult.ReturnFactoryServiceId & "|" & CheckActivityNumResult.ReturnPointPrice, "127.0.0.3")

        '紀錄活動代號
        Dim AddActivityNumResult As New MemberAddListRenderResult
        Try
            AddActivityNumResult = AddGetActivityNum(AuthCode, MyCardProjectNo)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "MemberAddListRender|-907|" & AuthCode & "," & MyCardProjectNo & "|" & ex.Message, "127.0.0.3")
            MyReturnResult.ReturnMsgNo = 1
            MyReturnResult.ReturnMsg = ex.Message
            Return MyReturnResult
        End Try

        If AddActivityNumResult.ReturnMsgNo <> 1 Then
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "MemberAddListRender|-908|" & AuthCode & "," & MyCardProjectNo & "|" & AddActivityNumResult.ReturnMsgNo & "|" & AddActivityNumResult.ReturnMsg, "127.0.0.3")
            MyReturnResult.ReturnMsgNo = 1
            MyReturnResult.ReturnMsg = AddActivityNumResult.ReturnMsg
            Return MyReturnResult
        End If


        '先授權
        Dim AuthResult As New MyCardMemberServiceAuthResult
        Dim BonusFacId As String = CheckActivityNumResult.ReturnFactoryId
        Dim BonusFacSrvId As String = CheckActivityNumResult.ReturnFactoryServiceId
        Dim BonusFacSrvAmt As Integer = CheckActivityNumResult.ReturnPointPrice

        Try
            AuthResult = MyCardMemberServiceAuthWithOtp(BonusFacId, BonusFacSrvId, GetUserInfoResult.ReturnMyCardCustId, OneTimePassword, "127.0.0.1", GetUserInfoResult.ReturnFacTradeSeq & "-Bonus", BonusFacSrvAmt, 0, "")
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "MemberAddListRender|-909|" & BonusFacId & "," & BonusFacSrvId & "," & GetUserInfoResult.ReturnMyCardCustId & "," & OneTimePassword & "," & "127.0.0.1" & "," & GetUserInfoResult.ReturnFacTradeSeq & "-Bonus" & "," & BonusFacSrvAmt & "|" & ex.Message, "127.0.0.3")
            MyReturnResult.ReturnMsgNo = 1
            MyReturnResult.ReturnMsg = ex.Message
            Return MyReturnResult
        End Try

        If AuthResult.ReturnMsgNo <> 1 Then
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "MemberAddListRender|-910|" & BonusFacId & "," & BonusFacSrvId & "," & GetUserInfoResult.ReturnMyCardCustId & "," & OneTimePassword & "," & "127.0.0.1" & "," & GetUserInfoResult.ReturnFacTradeSeq & "-Bonus" & "," & BonusFacSrvAmt & "|" & AuthResult.ReturnMsgNo, "127.0.0.3")
            MyReturnResult.ReturnMsgNo = 1
            MyReturnResult.ReturnMsg = AuthResult.ReturnMsgNo
            Return MyReturnResult
        End If

        '紀錄活動代號
        Dim AddActivityNumResult1 As New MemberAddListRenderResult
        Try
            AddActivityNumResult1 = AddGetActivityNum(AuthResult.ReturnAuthCode, MyCardProjectNo)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "MemberAddListRender|-911|" & AuthResult.ReturnAuthCode & "," & MyCardProjectNo & "|" & ex.Message, "127.0.0.3")
            MyReturnResult.ReturnMsgNo = 1
            MyReturnResult.ReturnMsg = ex.Message
            Return MyReturnResult
        End Try

        If AddActivityNumResult1.ReturnMsgNo <> 1 Then
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "MemberAddListRender|-912|" & AuthResult.ReturnAuthCode & "," & MyCardProjectNo & "|" & AddActivityNumResult.ReturnMsgNo & "|" & AddActivityNumResult.ReturnMsg, "127.0.0.3")
            MyReturnResult.ReturnMsgNo = 1
            MyReturnResult.ReturnMsg = AddActivityNumResult.ReturnMsg
            Return MyReturnResult
        End If

        GetUserInfoResult = MyCardMemberServiceMemberAddListGetUserInfo(AuthResult.ReturnAuthCode, OneTimePassword)

        If GetUserInfoResult.ReturnMsgNo <> 1 Then
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "MemberAddListRender|-913|" & AuthResult.ReturnAuthCode & "," & MyCardProjectNo & "|" & GetUserInfoResult.ReturnMsgNo & "|" & GetUserInfoResult.ReturnMsg & "|" & GetUserInfoResult.ReturnErrorCode, "127.0.0.3")
            MyReturnResult.ReturnMsgNo = 1
            MyReturnResult.ReturnMsg = GetUserInfoResult.ReturnMsg
            MyReturnResult.ReturnErrorCode = GetUserInfoResult.ReturnErrorCode
            Return MyReturnResult
        End If
        'log
        MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "2.MyCardCustPointAddEx|" & GetUserInfoResult.ReturnMyCardCustId & "|" & GetUserInfoResult.ReturnMyCardPointServiceType & "|" & GetUserInfoResult.ReturnMyCardPoint & "|" & GetUserInfoResult.ReturnMyCardBonus & "|" & GetUserInfoResult.MyCardServiceTradeListSn & "|" & GetUserInfoResult.ReturnMyCardServiceName & "|" & GetUserInfoResult.ReturnFacTradeSeq & "|" & AuthResult.ReturnAuthCode, "127.0.0.3")
        '直接加點
        MyCardPointAddResult = MyCardCustPointAddEx(GetUserInfoResult.ReturnMyCardCustId, GetUserInfoResult.ReturnMyCardPointServiceType, GetUserInfoResult.ReturnMyCardPoint, GetUserInfoResult.ReturnMyCardBonus, GetUserInfoResult.MyCardServiceTradeListSn, GetUserInfoResult.ReturnMyCardServiceName, "MYC_PA", GetUserInfoResult.ReturnFacTradeSeq, AuthResult.ReturnAuthCode)

        If MyCardPointAddResult.ReturnResult = False Then
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "MemberAddListRender|-913|" & AuthResult.ReturnAuthCode & "," & MyCardProjectNo & "|" & MyCardPointAddResult.ReturnMsgNo & "|" & MyCardPointAddResult.ReturnMsg & "|" & MyCardPointAddResult.ReturnErrorCode, "127.0.0.3")
            MyReturnResult.ReturnMsgNo = 1
            MyReturnResult.ReturnMsg = MyCardPointAddResult.ReturnMsg
            MyReturnResult.ReturnErrorCode = MyCardPointAddResult.ReturnErrorCode
            Return MyReturnResult
        End If


        MyReturnResult = MyCardMemberServiceMemberAddListRender(AuthResult.ReturnAuthCode, OneTimePassword)


        Return MyReturnResult
    End Function

    ''' <summary>
    ''' 記錄活動代號
    '''
    ''' </summary>
    ''' <param name="AuthCode">授權碼</param>
    ''' <param name="ActiveCode">活動代號</param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 
    ''' </remarks>
    Private Function AddGetActivityNum(ByVal AuthCode As String, ByVal ActiveCode As String) As MemberAddListRenderResult
        ConnStr(20)
        Dim ReturnResult As New MemberAddListRenderResult

        Dim command As New SqlClient.SqlCommand
        command.Connection = Con
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "MyCard_MemberService.dbo.MyCard_MemberService_AddGetActivityNum_A01"

        command.Parameters.Add("@AuthCode", SqlDbType.NVarChar, 50).Value = AuthCode
        command.Parameters.Add("@ActivityNum", SqlDbType.VarChar, 6).Value = ActiveCode

        Dim ReturnMsgNo As SqlClient.SqlParameter = Command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = Command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 1024)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ErrorCode As SqlClient.SqlParameter = Command.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output

        Dim LogSn As SqlClient.SqlParameter = Command.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output

        Try
            Con.Open()
            command.ExecuteNonQuery()
            ReturnResult.ReturnMsgNo = ReturnMsgNo.Value
            ReturnResult.ReturnMsg = ReturnMsg.Value
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "999|AddGetActivityNum|" & AuthCode & "|" & ActiveCode & "|" & ex.Message, "127.0.0.3")
            Throw New Exception("記錄活動代號時發生錯誤", ex)
        Finally
            Con.Close()
        End Try

        Return ReturnResult
    End Function

    ''' <summary>
    ''' 20120308 qq 帶入活動代號查詢加點服務資訊
    '''
    ''' </summary>
    ''' <param name="ActivityNum">活動代號</param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' 
    ''' </remarks>
    <WebMethod()> _
    Public Function MyCard_MemberService_CheckActivityNum(ByVal ActivityNum As String) As MyCardMemberCheckActivityNumResult
        ConnStr(20)
        Dim command As New SqlClient.SqlCommand
        command.Connection = Con
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "MyCard_MemberService.dbo.MyCard_MemberService_CheckActivityNum"

        command.Parameters.Add("@ActivityNum", SqlDbType.VarChar, 6).Value = ActivityNum

        Dim FactoryId As SqlClient.SqlParameter = command.Parameters.Add("@FactoryId", SqlDbType.NVarChar, 50)
        FactoryId.Direction = ParameterDirection.Output

        Dim FactoryServiceId As SqlClient.SqlParameter = command.Parameters.Add("@FactoryServiceId", SqlDbType.NVarChar, 50)
        FactoryServiceId.Direction = ParameterDirection.Output

        Dim PointPrice As SqlClient.SqlParameter = command.Parameters.Add("@PointPrice", SqlDbType.Int)
        PointPrice.Direction = ParameterDirection.Output

        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ErrorCode As SqlClient.SqlParameter = command.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 10)
        ErrorCode.Direction = ParameterDirection.Output

        Dim LogSn As SqlClient.SqlParameter = command.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output

        Dim ReturnResult As New MyCardMemberCheckActivityNumResult

        Try
            Con.Open()
            command.ExecuteNonQuery()
            ReturnResult.ReturnMsgNo = ReturnMsgNo.Value
            ReturnResult.ReturnMsg = ReturnMsg.Value
            ReturnResult.ReturnFactoryId = IIf(IsDBNull(FactoryId.Value), "", FactoryId.Value)
            ReturnResult.ReturnFactoryServiceId = IIf(IsDBNull(FactoryServiceId.Value), "", FactoryServiceId.Value)
            ReturnResult.ReturnPointPrice = IIf(IsDBNull(PointPrice.Value), 0, PointPrice.Value)
            ReturnResult.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnResult.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "999|MyCard_MemberService_CheckActivityNum|" & ActivityNum & "|" & ex.Message, "127.0.0.3")
            Throw New Exception("依活動代號查詢加點服務時發生錯誤", ex)
        Finally
            Con.Close()
        End Try

        Return ReturnResult
    End Function

    <WebMethod()> _
    Public Function MyCard_MemberService_AuthWithOtpByApp(ByVal FactoryId As String, ByVal FactoryServiceId As String, ByVal MyCardCustId As String, ByVal OneTimePassword As String, ByVal FactoryIp As String, ByVal FactorySeq As String, ByVal PointPayment As Integer, ByVal BonusPayment As Integer, ByVal FactoryReturnUrl As String) As MyCardMemberServiceAuthResult
        ConnStr(20)

        Dim ReturnResult As New MyCardMemberServiceAuthResult

        Dim command As New SqlClient.SqlCommand
        command.Connection = Con
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "MyCard_MemberService.dbo.MyCard_MemberService_AuthWithOtpByApp"

        command.Parameters.Add("@FactoryId", SqlDbType.NVarChar, 50).Value = FactoryId
        command.Parameters.Add("@FactoryServiceId", SqlDbType.NVarChar, 50).Value = FactoryServiceId
        command.Parameters.Add("@MyCardCustId", SqlDbType.NVarChar, 50).Value = MyCardCustId
        command.Parameters.Add("@OneTimePassword", SqlDbType.NVarChar, 50).Value = OneTimePassword
        command.Parameters.Add("@FactoryIp", SqlDbType.NVarChar, 50).Value = FactoryIp
        command.Parameters.Add("@FactorySeq", SqlDbType.NVarChar, 50).Value = FactorySeq
        command.Parameters.Add("@PointPayment", SqlDbType.Int).Value = PointPayment
        command.Parameters.Add("@BonusPayment", SqlDbType.Int).Value = BonusPayment
        command.Parameters.Add("@FactoryReturnUrl", SqlDbType.NVarChar, 250).Value = FactoryReturnUrl

        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnTradeSeq As SqlClient.SqlParameter = command.Parameters.Add("@ReturnTradeSeq", SqlDbType.NVarChar, 50)
        ReturnTradeSeq.Direction = ParameterDirection.Output

        Dim ReturnAuthCode As SqlClient.SqlParameter = command.Parameters.Add("@ReturnAuthCode", SqlDbType.NVarChar, 50)
        ReturnAuthCode.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ErrorCode As SqlClient.SqlParameter = command.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 10)
        ErrorCode.Direction = ParameterDirection.Output

        Dim LogSn As SqlClient.SqlParameter = command.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output

        Dim InputValue As String = FactoryId & "," & FactoryServiceId & "," & MyCardCustId & "," & OneTimePassword & "," & FactoryIp & "," & FactorySeq & "," & PointPayment & "," & BonusPayment & "," & FactoryReturnUrl
        Try
            Con.Open()
            command.ExecuteNonQuery()
            ReturnResult.ReturnMsgNo = ReturnMsgNo.Value
            ReturnResult.ReturnMsg = ReturnMsg.Value
            ReturnResult.ReturnTradeSeq = ReturnTradeSeq.Value
            ReturnResult.ReturnAuthCode = ReturnAuthCode.Value
            ReturnResult.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnResult.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "999|MyCard_MemberService_AuthWithOtpByApp|" & InputValue & "|" & ex.Message, "127.0.0.3")
            ReturnResult.ReturnMsgNo = "-999"
            ReturnResult.ReturnMsg = "系統發生錯誤|" & ex.Message
            ReturnResult.ReturnTradeSeq = ""
            ReturnResult.ReturnAuthCode = ""
            ReturnResult.ReturnErrorCode = "FBWS3006"
            ReturnResult.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Finally
            Con.Close()
        End Try

        Return ReturnResult
    End Function

    <WebMethod()> _
    Public Function MyCard_MemberOtpCheckForApp_GetOtp(ByVal MyCardCustId As String) As GetOTPResult
        ConnStr(19)

        Dim ReturnResult As New GetOTPResult

        Dim command As New SqlClient.SqlCommand
        command.Connection = Con
        command.CommandType = CommandType.StoredProcedure
        command.CommandText = "MyCard_Member.dbo.MyCard_MemberOtpCheckForApp_GetOtp"

        command.Parameters.Add("@MyCardCustId", SqlDbType.NVarChar, 50).Value = MyCardCustId

        Dim ReturnMsgNo As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = command.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 100)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim MyCardOTP As SqlClient.SqlParameter = command.Parameters.Add("@MyCardOTP", SqlDbType.NVarChar, 50)
        MyCardOTP.Direction = ParameterDirection.Output

        Dim InputValue As String = MyCardCustId
        Try
            Con.Open()
            command.ExecuteNonQuery()
            ReturnResult.ReturnMsgNo = ReturnMsgNo.Value
            ReturnResult.ReturnMsg = ReturnMsg.Value
            ReturnResult.MyCardOTP = IIf(IsDBNull(MyCardOTP.Value), "", MyCardOTP.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("MyCardMemService.asmx"), "999|MyCard_MemberOtpCheckForApp_GetOtp|" & InputValue & "|" & ex.Message, "127.0.0.3")
            ReturnResult.ReturnMsgNo = "-999"
            ReturnResult.ReturnMsg = "系統發生錯誤|" & ex.Message
            ReturnResult.MyCardOTP = IIf(IsDBNull(MyCardOTP.Value), "", MyCardOTP.Value)
        Finally
            Con.Close()
        End Try

        Return ReturnResult
    End Function
End Class


Public Class MyCardMemberServiceAuthResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnTradeSeq As String
    Public ReturnAuthCode As String
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer
    Public Sub New()
        ReturnTradeSeq = ""
        ReturnAuthCode = ""
    End Sub
End Class

Public Class MyCardMemberServiceGetUserInfoResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnMyCardCustId As String
    Public ReturnMyCardPointServiceType As String
    Public ReturnMyCardServiceName As String
    Public MyCardServiceTradeListSn As Integer
    Public ReturnMyCardPoint As Integer
    Public ReturnMyCardTPoint As Integer
    Public ReturnMyCardBonus As Integer
    Public ReturnFacTradeSeq As String
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer

    Public Sub New()
        ReturnMyCardCustId = ""
        ReturnMyCardPointServiceType = ""
        ReturnMyCardServiceName = ""
        ReturnFacTradeSeq = ""
    End Sub
End Class
Public Class MemberAddListRenderResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnErrorCode As String
    Sub New()
        ReturnMsg = ""
    End Sub
End Class
Public Class MyCardPointAddExResult
    Public ReturnResult As Boolean
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnBonusTradeSeq As String
    Public ReturnPointTradeSeq As String
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer
    Sub New()
        ReturnResult = False
        ReturnMsg = ""
        ReturnBonusTradeSeq = ""
        ReturnPointTradeSeq = ""
    End Sub
End Class
Public Class MyCardMemberCheckActivityNumResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnFactoryId As String
    Public ReturnFactoryServiceId As String
    Public ReturnPointPrice As Integer
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer
    Sub New()
        ReturnMsg = ""
        ReturnErrorCode = ""
        ReturnLogSn = 0
    End Sub
End Class
Public Class GetOTPResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public MyCardOTP As String
End Class