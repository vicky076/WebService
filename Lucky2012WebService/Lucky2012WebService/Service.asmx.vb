Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Data.SqlClient

<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class Service
    Inherits System.Web.Services.WebService

    Dim myDbfnc As New MyCardConn.Dbfuction
    Dim conn As New SqlClient.SqlConnection()

    Public Sub ConnStr(ByVal ConnNum As Integer)
        '正式環境用連19 , 測試用連100 
        If My.Settings.DebugMode.ToUpper = "FALSE" Then
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
                Case 16
                    conn = myDbfnc.open16Db
                Case 17
                    conn = myDbfnc.open17Db
                Case 19
                    conn = myDbfnc.open19Db
                Case 20
                    conn = myDbfnc.open20Db
                Case 21
                    conn = myDbfnc.open21Db
            End Select
        ElseIf My.Settings.DebugMode.ToUpper = "DEV" Then
            conn = myDbfnc.open100Db()
        ElseIf My.Settings.DebugMode.ToUpper = "TEST" Then
            conn = myDbfnc.openTestDb
        Else
            conn = myDbfnc.open100Db()
        End If
    End Sub

    ''' <summary>
    ''' 龍年好手氣活動參加條件，回傳ReturnMsg,ReturnMsgNo,ReturnJoinTime
    ''' </summary>
    ''' <param name="ActId">活動代號</param>
    ''' <param name="MyCardID">MyCard卡號</param>
    ''' <returns>ReturnMsg,ReturnMsgNo,ReturnJoinTime</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function MyCardMemberForSavCheck(ByVal ActId As String, ByVal MyCardID As String) As MyCardMemberForSavCheckREsult
        ConnStr(19)
        Dim cmd As New SqlClient.SqlCommand("MyActivity.dbo.MyActivity_Codition_NotMyCardMember_UsePrizeList_CheckCardType_1", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandTimeout = 60
        ' 增加參數欄位
        cmd.Parameters.Add("@ActId", SqlDbType.VarChar, 10).Value = ActId
        cmd.Parameters.Add("@MyCardID", SqlDbType.VarChar, 16).Value = MyCardID

        Dim ReturnJoinTime As SqlClient.SqlParameter = cmd.Parameters.Add("@JoinTime", SqlDbType.Int)
        ReturnJoinTime.Direction = ParameterDirection.Output

        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 20)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New MyCardMemberForSavCheckREsult
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        Finally
            conn.Close()
        End Try

        ReturnValue.ReturnJoinTime = ReturnJoinTime.Value
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value

        Return ReturnValue
    End Function

    ''' <summary>
    ''' 龍年好手氣活動Check資料後取卡然後寫入資料庫
    ''' </summary>
    ''' <param name="ChoiceType">查詢類型</param>
    ''' <param name="MyCardID">MyCard卡號</param>
    ''' <param name="TradeSeq">交易序號</param>
    ''' <param name="CreateIp">IP</param>
    ''' <param name="ActId">活動代號</param>
    ''' <returns>ReturnMsg,ReturnMsgNo,ReturnDS</returns>
    ''' <remarks>
    ''' ChoiceType：1是取虛寶,2是查詢卡密
    ''' </remarks>
    <WebMethod()> _
    Public Function ConductProceduresCheckInsert(ByVal ChoiceType As Integer, ByVal MyCardID As String, ByVal TradeSeq As String, ByVal CreateIp As String, ByVal ActId As String) As ConductProceduresCheckInsertREsult
        ConnStr(19)
        Dim cmd As New SqlClient.SqlCommand("MyActivity.dbo.MyActivity_ConductProcedures_CheckList_Insert_CheckCardType_1", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandTimeout = 60
        ' 增加參數欄位
        cmd.Parameters.Add("@ChoiceType", SqlDbType.Int).Value = ChoiceType
        cmd.Parameters.Add("@MyCardID", SqlDbType.VarChar, 50).Value = MyCardID
        cmd.Parameters.Add("@CreateIp", SqlDbType.VarChar, 50).Value = CreateIp
        cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 100).Value = TradeSeq
        cmd.Parameters.Add("@ActId", SqlDbType.VarChar, 10).Value = ActId

        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 100)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ConductProceduresCheckInsertREsult
        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        Try
            da.Fill(ds)
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try
        If ReturnMsgNo.Value = 1 Then
            ReturnValue.ReturnDS = ds
        End If
        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value

        Return ReturnValue
    End Function

End Class

Public Class MyCardMemberForSavCheckREsult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnJoinTime As String

    Sub New()
        ReturnMsg = ""
    End Sub
End Class

Public Class ConductProceduresCheckInsertREsult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnDS As New DataSet

    Sub New()
        ReturnMsg = ""
    End Sub
End Class