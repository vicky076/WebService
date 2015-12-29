Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Data.SqlClient

' 若要允許使用 ASP.NET AJAX 從指令碼呼叫此 Web 服務，請取消註解下一行。
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class ePlayWebServiceMAS
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


    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetBulletin(ByVal Bul_Subject As String, ByVal Bul_IsVisible As String) As DataSet
        ConnStr(19)
        ' 建立命令(預存程序名,使用連線) 
        Dim cmd As New SqlClient.SqlCommand
        cmd.CommandTimeout = 300
        cmd.CommandType = CommandType.Text
        cmd.Connection = conn

        If Bul_IsVisible = "" Then
            If Bul_Subject = "" Then
                cmd.CommandText = "SELECT Bul_BulletinID,Bul_Subject,Bul_Content,Bul_StartDate,Bul_EndDate,Bul_CreateDate,Bul_CreateUID,Bul_IsVisible FROM [MyCard_CardService_MAS].[dbo].[swvw_GetAllBulletin] order by Bul_BulletinID desc"
            Else
                cmd.CommandText = "SELECT Bul_BulletinID,Bul_Subject,Bul_Content,Bul_StartDate,Bul_EndDate,Bul_CreateDate,Bul_CreateUID,Bul_IsVisible FROM [MyCard_CardService_MAS].[dbo].[swvw_GetAllBulletin] where Bul_Subject like '%" & Trim(Bul_Subject) & "%' order by Bul_BulletinID desc"
            End If
        Else
            If Bul_Subject = "" Then
                cmd.CommandText = "SELECT Bul_BulletinID,Bul_Subject,Bul_Content,Bul_StartDate,Bul_EndDate,Bul_CreateDate,Bul_CreateUID,Bul_IsVisible FROM [MyCard_CardService_MAS].[dbo].[swvw_GetAllBulletin] where Bul_IsVisible=@Bul_IsVisible order by Bul_BulletinID desc"
                If Bul_IsVisible = 1 Then
                    cmd.Parameters.Add("@Bul_IsVisible", SqlDbType.Bit).Value = True
                Else
                    If Bul_IsVisible = 0 Then
                        cmd.Parameters.Add("@Bul_IsVisible", SqlDbType.Bit).Value = False
                    End If
                End If

            Else
                cmd.CommandText = "SELECT Bul_BulletinID,Bul_Subject,Bul_Content,Bul_StartDate,Bul_EndDate,Bul_CreateDate,Bul_CreateUID,Bul_IsVisible FROM [MyCard_CardService_MAS].[dbo].[swvw_GetAllBulletin] where Bul_Subject like '%" & Trim(Bul_Subject) & "%' and Bul_IsVisible=@Bul_IsVisible order by Bul_BulletinID desc"
                If Bul_IsVisible = 1 Then
                    cmd.Parameters.Add("@Bul_IsVisible", SqlDbType.Bit).Value = True
                Else
                    If Bul_IsVisible = 0 Then
                        cmd.Parameters.Add("@Bul_IsVisible", SqlDbType.Bit).Value = False
                    End If
                End If
            End If
        End If

        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("查詢公告資料失敗：" & ex.Message)
        End Try
        Return DSssid
    End Function

    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function swsp_AddNewBulletin(ByVal Bul_Subject As String, ByVal Bul_Content As String, ByVal Bul_StartDate As DateTime,
                                                                       ByVal Bul_EndDate As DateTime, ByVal Bul_CreateUID As Integer) As ReturnValue
        ConnStr(19)

        Dim com As New SqlClient.SqlCommand("MyCard_CardService_MAS.dbo.swsp_AddNewBulletin", conn)
        com.CommandType = CommandType.StoredProcedure
        com.CommandTimeout = 300

        com.Parameters.Add("@Bul_Subject", SqlDbType.Char, 128).Value = Bul_Subject
        com.Parameters.Add("@Bul_Content", SqlDbType.NVarChar, 1024).Value = Bul_Content
        com.Parameters.Add("@Bul_StartDate", SqlDbType.DateTime).Value = Bul_StartDate
        com.Parameters.Add("@Bul_EndDate", SqlDbType.DateTime).Value = CDate(Bul_EndDate).AddDays(1)
        com.Parameters.Add("@Bul_CreateDate", SqlDbType.DateTime).Value = Now
        com.Parameters.Add("@Bul_CreateUID", SqlDbType.Int).Value = Bul_CreateUID
        com.Parameters.Add("@Bul_IsVisible", SqlDbType.Bit).Value = True

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim Returnds As New ReturnValue
        Try
            conn.Open()
            com.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("Insert公告時：" & ex.Message)
        Finally
            conn.Close()
        End Try
        Returnds.ReturnMsgNo = ReturnMsgNo.Value
        Return Returnds
    End Function


    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function swsp_UpdateBulletin(ByVal Bul_BulletinID As Integer, ByVal Bul_Subject As String, ByVal Bul_Content As String,
                                                                     ByVal Bul_StartDate As DateTime, ByVal Bul_EndDate As DateTime, ByVal Bul_CreateUID As Integer) As ReturnValue
        ConnStr(19)

        Dim com As New SqlClient.SqlCommand("MyCard_CardService_MAS.dbo.swsp_UpdateBulletin", conn)
        com.CommandType = CommandType.StoredProcedure
        com.CommandTimeout = 300

        com.Parameters.Add("@Bul_BulletinID", SqlDbType.Int).Value = Bul_BulletinID
        com.Parameters.Add("@Bul_Subject", SqlDbType.Char, 128).Value = Bul_Subject
        com.Parameters.Add("@Bul_Content", SqlDbType.NVarChar, 1024).Value = Bul_Content
        com.Parameters.Add("@Bul_StartDate", SqlDbType.DateTime).Value = Bul_StartDate
        com.Parameters.Add("@Bul_EndDate", SqlDbType.DateTime).Value = CDate(Bul_EndDate).AddDays(1)
        com.Parameters.Add("@Bul_CreateDate", SqlDbType.DateTime).Value = Now
        com.Parameters.Add("@Bul_CreateUID", SqlDbType.Int).Value = Bul_CreateUID

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim Returnds As New ReturnValue
        Try
            conn.Open()
            com.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("Update公告時：" & ex.Message)
        Finally
            conn.Close()
        End Try
        Returnds.ReturnMsgNo = ReturnMsgNo.Value
        Return Returnds
    End Function

    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function swsp_DeleteBulletin(ByVal Bul_BulletinID As Integer, ByVal Bul_CreateUID As Integer) As ReturnValue
        ConnStr(19)

        Dim com As New SqlClient.SqlCommand("MyCard_CardService_MAS.dbo.swsp_DeleteBulletin", conn)
        com.CommandType = CommandType.StoredProcedure
        com.CommandTimeout = 300

        com.Parameters.Add("@Bul_BulletinID", SqlDbType.Int).Value = Bul_BulletinID
        com.Parameters.Add("@Bul_CreateUID", SqlDbType.Int).Value = Bul_CreateUID

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output

        Dim Returnds As New ReturnValue
        Try
            conn.Open()
            com.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception("Delete公告時：" & ex.Message)
        Finally
            conn.Close()
        End Try
        Returnds.ReturnMsgNo = ReturnMsgNo.Value
        Return Returnds
    End Function

    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function EMS_CS_Ad_Group_Insert(ByVal GroupName As String, ByVal CreateUser As String) As ReturnValue
        ConnStr(19)

        Dim com As New SqlClient.SqlCommand("MyCard_CardService_MAS.dbo.EMS_CS_Ad_Group_Insert", conn)
        com.CommandType = CommandType.StoredProcedure
        com.CommandTimeout = 300

        com.Parameters.Add("@GroupName", SqlDbType.VarChar, 30).Value = GroupName
        com.Parameters.Add("@CreateUser", SqlDbType.VarChar, 20).Value = CreateUser

        Dim ReturnNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim Returnds As New ReturnValue
        Try
            conn.Open()
            com.ExecuteNonQuery()
            Returnds.ReturnMsgNo = ReturnNo.Value
            Returnds.ReturnMsg = ReturnMsg.Value
        Catch ex As Exception
            Returnds.ReturnMsgNo = -1
            Returnds.ReturnMsg = "insert廣宣群組時發生錯誤：" & ex.Message
            Return Returnds
        Finally
            conn.Close()
        End Try

        Return Returnds
    End Function

    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function EMS_CS_Ad_Group_Update(ByVal GroupSn As String, ByVal GroupName As String, ByVal UserStamp As String) As ReturnValue
        ConnStr(19)

        Dim com As New SqlClient.SqlCommand("MyCard_CardService_MAS.dbo.EMS_CS_Ad_Group_Update", conn)
        com.CommandType = CommandType.StoredProcedure
        com.CommandTimeout = 300
        com.Parameters.Add("@GroupNameSn", SqlDbType.Int).Value = GroupSn
        com.Parameters.Add("@GroupName", SqlDbType.VarChar, 30).Value = GroupName
        com.Parameters.Add("@UserStamp", SqlDbType.VarChar, 20).Value = UserStamp

        Dim ReturnNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim Returnds As New ReturnValue
        Try
            conn.Open()
            com.ExecuteNonQuery()
            Returnds.ReturnMsgNo = ReturnNo.Value
            Returnds.ReturnMsg = ReturnMsg.Value
        Catch ex As Exception
            Returnds.ReturnMsgNo = -1
            Returnds.ReturnMsg = "update廣宣群組時發生錯誤：" & ex.Message
            Return Returnds
        Finally
            conn.Close()
        End Try

        Return Returnds
    End Function


    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function EMS_Ad_Group_Query(ByVal Name As String) As DataSet
        ConnStr(19)
        ' 建立命令(預存程序名,使用連線) 
        Dim cmd As New SqlClient.SqlCommand
        cmd.CommandTimeout = 300
        cmd.CommandType = CommandType.Text
        cmd.Connection = conn

         Dim Sql As String
        If Name <> "" Then
            Sql = "SELECT Sn,Name, CreateUser , UserStamp , CreateDate ,DateStamp  FROM [MyCard_CardService_MAS].[dbo].[View_EMS_Ad_Group]  Where Name like '%" & Name & "%' ORDER BY Name,CreateDate"
        Else
            Sql = "SELECT Sn,Name, CreateUser , UserStamp , CreateDate ,DateStamp  FROM [MyCard_CardService_MAS].[dbo].[View_EMS_Ad_Group]   ORDER BY Name,CreateDate"
        End If
        cmd.CommandText = Sql

        'conn.Open()
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("WebService廣宣群組資料失敗", ex)
        End Try
        Return DSssid
    End Function

    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function EMS_CS_Remind_Insert(ByVal HeadStoreId As String, ByVal LimitRatio As Double, ByVal CreateUser As String) As ReturnValue
        ConnStr(19)

        Dim com As New SqlClient.SqlCommand("MyCard_CardService_MAS.dbo.EMS_CS_Remind_Insert", conn)
        com.CommandType = CommandType.StoredProcedure
        com.CommandTimeout = 300
        com.Parameters.Add("@HeadStoreId", SqlDbType.VarChar, 10).Value = HeadStoreId
        com.Parameters.Add("@LimitRatio", SqlDbType.Decimal).Value = LimitRatio
        com.Parameters.Add("@CreateUser", SqlDbType.VarChar, 20).Value = CreateUser

        Dim ReturnNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.VarChar, 50)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim Returnds As New ReturnValue
        Try
            conn.Open()
            com.ExecuteNonQuery()
            Returnds.ReturnMsgNo = ReturnNo.Value
            Returnds.ReturnMsg = ReturnMsg.Value
        Catch ex As Exception
            Returnds.ReturnMsgNo = -1
            Returnds.ReturnMsg = "insert提醒額度時發生錯誤：" & ex.Message
            Return Returnds
        Finally
            conn.Close()
        End Try

        Return Returnds
    End Function

    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function EMS_CS_Remind_Update(ByVal HeadStoreId As String, ByVal LimitRatio As Double, ByVal UserStamp As String) As ReturnValue
        ConnStr(19)

        Dim com As New SqlClient.SqlCommand("MyCard_CardService_MAS.dbo.EMS_CS_Remind_Update", conn)
        com.CommandType = CommandType.StoredProcedure
        com.CommandTimeout = 300
        com.Parameters.Add("@HeadStoreId", SqlDbType.VarChar, 10).Value = HeadStoreId
        com.Parameters.Add("@LimitRatio", SqlDbType.Decimal).Value = LimitRatio
        com.Parameters.Add("@UserStamp", SqlDbType.VarChar, 20).Value = UserStamp

        Dim ReturnNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int, 4)
        ReturnNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.VarChar, 50)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim Returnds As New ReturnValue
        Try
            conn.Open()
            com.ExecuteNonQuery()
            Returnds.ReturnMsgNo = ReturnNo.Value
            Returnds.ReturnMsg = ReturnMsg.Value
        Catch ex As Exception
            Returnds.ReturnMsgNo = -1
            Returnds.ReturnMsg = "update提醒額度時發生錯誤：" & ex.Message
            Return Returnds
        Finally
            conn.Close()
        End Try

        Return Returnds
    End Function

    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function EMS_CS_Remind_Query(ByVal StoreId As String) As DataSet
        ConnStr(19)
        ' 建立命令(預存程序名,使用連線) 
        Dim cmd As New SqlClient.SqlCommand
        cmd.CommandTimeout = 300
        cmd.CommandType = CommandType.Text
        cmd.Connection = conn

        Dim Sql As String = "SELECT * FROM [MyCard_CardService_MAS].[dbo].[View_MCS_HeadStoreRemind] where 1=1 And State = 1"
        If StoreId <> "" Then
            Sql = Sql & " And HeadStoreId=@StoreId"
        End If
        cmd.CommandText = Sql
        If StoreId <> "-1" Then
            cmd.Parameters.Add("@StoreId", SqlDbType.VarChar, 10).Value = StoreId
        End If
        'conn.Open()
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("WebService查詢提醒額度資料失敗", ex)
        End Try
        Return DSssid
    End Function

    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function GetHeadStore() As DataSet
        ConnStr(19)
        ' 建立命令(預存程序名,使用連線) 
        Dim cmd As New SqlClient.SqlCommand
        cmd.CommandTimeout = 300
        cmd.CommandType = CommandType.Text
        cmd.Connection = conn

        cmd.CommandText = "SELECT Sn, HeadStoreId, StoreName, StoreShortName, Address, Telephone, BusinessNumber, Manager, CreditLimit, Pos,State, LockState, DebtState FROM [MyCard_CardService_MAS].[dbo].[View_MCS_HeadStore] where  State = 1  order by StoreName"

        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("WebService查詢資料失敗", ex)
        End Try

        Return DSssid
    End Function

    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function EMS_CS_Message_Insert(ByVal TextContent As String, ByVal Status As Integer, ByVal CreateUser As String) As ReturnValue
        ConnStr(19)

        Dim com As New SqlClient.SqlCommand("MyCard_CardService_MAS.dbo.EMS_CS_Message_Insert", conn)
        com.Connection = conn
        com.CommandType = CommandType.StoredProcedure
        com.CommandTimeout = 60000

        com.Parameters.Add("@TextContent", SqlDbType.VarChar, 60).Value = TextContent
        com.Parameters.Add("@Status", SqlDbType.TinyInt).Value = Status
        com.Parameters.Add("@CreateUser", SqlDbType.VarChar, 20).Value = CreateUser

        Dim ReturnMsgNo As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = com.Parameters.Add("@ReturnMsg", SqlDbType.VarChar, 20)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim Returnds As New ReturnValue
        Try
            conn.Open()
            com.ExecuteNonQuery()
            Returnds.ReturnMsgNo = ReturnMsgNo.Value
            Returnds.ReturnMsg = ReturnMsg.Value
        Catch ex As Exception
            Returnds.ReturnMsgNo = -1
            Returnds.ReturnMsg = "新增失敗：" & ex.Message
            Return Returnds
        Finally
            conn.Close()
        End Try

        Return Returnds
    End Function

    ''' 回傳DataSet 
    ''' </remarks>
    <WebMethod()> _
    Public Function View_EMS_Message() As DataSet
        ConnStr(19)
        ' 建立命令(預存程序名,使用連線) 
        Dim cmd As New SqlClient.SqlCommand
        cmd.CommandTimeout = 300
        cmd.CommandType = CommandType.Text
        cmd.Connection = conn

        cmd.CommandText = "SELECT * FROM [MyCard_CardService_MAS].[dbo].[View_EMS_Message] "
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DSssid As New DataSet
        Try
            SDAssid.Fill(DSssid)
        Catch ex As Exception
            Throw New Exception("查詢資料失敗", ex)
        Finally
            conn.Close()
        End Try

        Return DSssid
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