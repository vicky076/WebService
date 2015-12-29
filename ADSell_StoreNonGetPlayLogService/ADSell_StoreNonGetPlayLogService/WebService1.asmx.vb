Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Data.SqlClient
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class WebService1
    Inherits System.Web.Services.WebService
    Dim myDbfnc As New MyCardConn.Dbfuction
    Dim myconn As New MyCardADSellConn.Dbfuction
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
                    Case 22
                        conn = myconn.MyCard_ADSell_DB
                End Select
            Case Else
                conn = myDbfnc.open100Db
        End Select
    End Sub

    ''' <summary>
    ''' 檢查未撈取影音店家
    ''' </summary>
    ''' <returns>ReturnValue</returns>
    ''' <remarks>
    ''' 回傳ReturnValue 
    ''' </remarks>
    <WebMethod()> _
    Public Function MyCard_ADSell_StoreNonGetPlayLog(ByVal P_dDate_S As DateTime, ByVal P_dDate_E As DateTime) As ReturnValue
        ConnStr(22)
        ' 建立命令(預存程序名,使用連線) 
        Dim cmd As New SqlClient.SqlCommand("MyCard_ADSell_DB.dbo.MyCard_ADSell_StoreNonGetPlayLog", conn)
        ' 設置命令型態為預存程序 
        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandTimeout = 300

        cmd.Parameters.Add("@P_dDate_S", SqlDbType.DateTime).Value = P_dDate_S
        cmd.Parameters.Add("@P_dDate_E", SqlDbType.DateTime).Value = P_dDate_E

        Dim ReturnMsgNo As SqlClient.SqlParameter = cmd.Parameters.Add("@P_iReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = cmd.Parameters.Add("@P_cReturnMsg ", SqlDbType.VarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output

        Dim ReturnValue As New ReturnValue
        Dim SDAssid As New System.Data.SqlClient.SqlDataAdapter(cmd)
        Dim DBViewDS As New DataSet

        ' 建立命令(預存程序名,使用連線) 
        Try
            conn.Open()
            SDAssid.Fill(DBViewDS)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            conn.Close()
        End Try

        ReturnValue.ReturnMsgNo = ReturnMsgNo.Value
        ReturnValue.ReturnMsg = ReturnMsg.Value
        ReturnValue.ReturnDS = DBViewDS
        Return ReturnValue
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