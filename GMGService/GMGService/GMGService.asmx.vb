Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Data.SqlClient
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class GMGService
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
                    Case 17
                        conn = myDbfnc.open17Db
                    Case 19
                        conn = myDbfnc.open19Db
                    Case 20
                        conn = myDbfnc.open20Db
                    Case 21
                        conn = myDbfnc.open21Db
                End Select
            Case Else
                conn = myDbfnc.open100Db
        End Select
    End Sub


    ''' <summary>
    ''' CPSaveSvice核對CP儲值差異- 查詢排程資料
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function View_MyCardCPSave_ScheduleCheck(ByVal GameFacId As String) As DataSet
        ConnStr(14)
        ' 建立命令(預存程序名,使用連線) 
        Dim cmd As New SqlClient.SqlCommand
        cmd.CommandType = CommandType.Text
        cmd.CommandTimeout = 300
        cmd.Connection = conn
        cmd.CommandText = "SELECT Sn,GameFacId,GameSerId,SaveDate,Status,CreateDate,CreateUser, DateStamp,UserStamp  from  Mycard_backup.dbo.View_MyCardCPSave_ScheduleCheck where 1=1 and  Status = 0 AND GameFacId = '" & GameFacId & "'"

        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("查詢失敗")
        End Try
        Return DSssid
    End Function



    ''' <summary>
    ''' CPSaveSvice核對CP儲值差異- 寫入TempTable
    ''' </summary>
    ''' <returns>ReturnValue</returns>
    ''' <remarks>
    ''' 回傳ReturnValue 
    ''' </remarks>
    <WebMethod()> _
    Public Function MyCardCPSave_CPTempData_Insert(ByVal PrcType As Integer, ByVal CP_TradeSeq As String, ByVal GameNo As String, ByVal ProductId As String, ByVal Item As String, ByVal Price As Integer, ByVal SaveDate As DateTime) As ReturnValue
        ConnStr(14)
        ' 建立命令(預存程序名,使用連線) 
        Dim cmd As New SqlClient.SqlCommand("Mycard_backup.dbo.MyCardCPSave_CPTempData_Insert", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandTimeout = 300

        cmd.Parameters.Add("@PrcType", SqlDbType.Int).Value = PrcType
        cmd.Parameters.Add("@CP_TradeSeq", SqlDbType.VarChar, 50).Value = CP_TradeSeq
        cmd.Parameters.Add("@GameNo", SqlDbType.VarChar, 50).Value = GameNo
        cmd.Parameters.Add("@ProductId", SqlDbType.VarChar, 50).Value = ProductId
        cmd.Parameters.Add("@Item", SqlDbType.VarChar, 50).Value = Item
        cmd.Parameters.Add("@Price", SqlDbType.Int).Value = Price
        cmd.Parameters.Add("@SaveDate", SqlDbType.DateTime).Value = SaveDate


        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg ", SqlDbType.VarChar, 50)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim Returnds As New ReturnValue
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("CPSaveSvice核對CP儲值差異- 寫入TempTable失敗：" & ex.Message)
        Finally
            conn.Close()
        End Try
        Returnds.ReturnMsgNo = ReturnMsgNo.Value
        Returnds.ReturnMsg = ReturnMsg.Value
        Return Returnds
    End Function

    ''' <summary>
    ''' CPSaveSvice核對CP儲值差異- 比對儲值差異
    ''' </summary>
    ''' <returns>DataSet</returns>
    ''' <remarks>
    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function MyCardCPSave_Difference_Proc(ByVal GameFacId As String, ByVal GameSerId As String, ByVal SaveStartDate As DateTime, ByVal SaveEndDate As DateTime, ByVal CreateUser As String) As ReturnValue
        ConnStr(14)
        ' 建立命令(預存程序名,使用連線) 
        Dim cmd As New SqlClient.SqlCommand("Mycard_backup.dbo.MyCardCPSave_Difference_Proc", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandTimeout = 300

        cmd.Parameters.Add("@GameFacId", SqlDbType.VarChar, 50).Value = GameFacId
        cmd.Parameters.Add("@GameSerId", SqlDbType.VarChar, 50).Value = GameSerId
        cmd.Parameters.Add("@SaveStartDate", SqlDbType.DateTime).Value = SaveStartDate
        cmd.Parameters.Add("@SaveEndDate", SqlDbType.DateTime).Value = SaveEndDate
        cmd.Parameters.Add("@CreateUser", SqlDbType.VarChar, 30).Value = CreateUser

        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg", SqlDbType.VarChar, 50)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim Returnds As New ReturnValue
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("比對儲值差異失敗：" & ex.Message)
        Finally
            conn.Close()
        End Try
        Returnds.ReturnMsgNo = ReturnMsgNo.Value
        Returnds.ReturnMsg = ReturnMsg.Value
        Return Returnds
    End Function

    ''' <summary>
    ''' CPSaveSvice核對CP儲值差異- 修改排程狀態
    ''' </summary>
    ''' <returns>ReturnValue</returns>
    ''' <remarks>
    ''' 回傳ReturnValue 
    ''' </remarks>
    <WebMethod()> _
    Public Function MyCardCPSave_ScheduleCheck_UpdateStatus(ByVal Sn As Integer, ByVal Status As Integer, ByVal UserStamp As String) As ReturnValue
        ConnStr(14)
        ' 建立命令(預存程序名,使用連線) 
        Dim cmd As New SqlClient.SqlCommand("Mycard_backup.dbo.MyCardCPSave_ScheduleCheck_UpdateStatus", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandTimeout = 300

        cmd.Parameters.Add("@Sn", SqlDbType.Int).Value = Sn
        cmd.Parameters.Add("@Status", SqlDbType.TinyInt).Value = Status
        cmd.Parameters.Add("@UserStamp", SqlDbType.VarChar, 30).Value = UserStamp

        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@ReturnMsg ", SqlDbType.VarChar, 50)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim Returnds As New ReturnValue
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("修改排程狀態失敗：" & ex.Message)
        Finally
            conn.Close()
        End Try
        Returnds.ReturnMsgNo = ReturnMsgNo.Value
        Returnds.ReturnMsg = ReturnMsg.Value
        Return Returnds
    End Function
End Class

Public Class ReturnValue
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnDS As DataSet
    Sub New()
        ReturnMsg = ""
        ReturnDS = New DataSet
    End Sub
End Class