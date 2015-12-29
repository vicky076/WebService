Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Net
Imports System.IO

<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class FacebookService
    Inherits System.Web.Services.WebService

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="PostUrl">POST網址</param>
    ''' <param name="PostString">POST參數</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' 20120328 qq 增加傳入POST網址參數
    ''' </remarks>
    <WebMethod()> _
    Public Function PostnGetResponse(ByVal PostUrl As String, ByVal PostString As String) As String
        If Not String.IsNullOrEmpty(PostString) Then
            'Dim restServerUrl As String = My.Settings.restServerUrl
            Dim parameterString As Byte() = Encoding.ASCII.GetBytes(PostString)

            Dim WebRequest As HttpWebRequest = HttpWebRequest.Create(PostUrl.Trim)
            WebRequest.Method = "POST"
            WebRequest.Accept = "application/xml"
            WebRequest.ContentType = "application/x-www-form-urlencoded"
            WebRequest.ContentLength = parameterString.Length

            Dim newStream As Stream = WebRequest.GetRequestStream()
            newStream.Write(parameterString, 0, parameterString.Length)
            newStream.Close()

            Dim WebResponse As HttpWebResponse = CType(WebRequest.GetResponse(), HttpWebResponse)

            Dim sr As StreamReader = New StreamReader(WebResponse.GetResponseStream(), Encoding.Default)
            'Convert the stream to a string
            Dim ReturnString As String = sr.ReadToEnd()
            sr.Close()
            WebResponse.Close()

            Return ReturnString
        Else
            Return ""
        End If
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="PostUrl">POST網址</param>
    ''' <param name="PostString">POST參數</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' 20130618 紹安 回傳JSON
    ''' </remarks>
    <WebMethod()> _
    Public Function PostnGetResponseJSON(ByVal PostUrl As String, ByVal PostString As String) As String
        If Not String.IsNullOrEmpty(PostString) Then
            'Dim restServerUrl As String = My.Settings.restServerUrl
            Dim parameterString As Byte() = Encoding.ASCII.GetBytes(PostString)

            Dim WebRequest As HttpWebRequest = HttpWebRequest.Create(PostUrl.Trim)
            WebRequest.Method = "POST"
            WebRequest.Accept = "application/json"
            WebRequest.ContentType = "application/x-www-form-urlencoded"
            WebRequest.ContentLength = parameterString.Length

            Dim newStream As Stream = WebRequest.GetRequestStream()
            newStream.Write(parameterString, 0, parameterString.Length)
            newStream.Close()

            Dim WebResponse As HttpWebResponse = CType(WebRequest.GetResponse(), HttpWebResponse)

            Dim sr As StreamReader = New StreamReader(WebResponse.GetResponseStream(), Encoding.Default)
            'Convert the stream to a string
            Dim ReturnString As String = sr.ReadToEnd()
            sr.Close()
            WebResponse.Close()

            Return ReturnString
        Else
            Return ""
        End If
    End Function

    <WebMethod()> _
    Public Function GetResponseHTML(ByVal ResponseWord As String) As String
        Try
            Dim oHttpRequest As HttpWebRequest = WebRequest.Create(ResponseWord)
            'ServicePointManager.ServerCertificateValidationCallback = AddressOf ValidateCertificate
            Dim ohttpResponse As HttpWebResponse = oHttpRequest.GetResponse
            Dim MyStream As IO.Stream
            MyStream = ohttpResponse.GetResponseStream
            Dim StreamReader As New IO.StreamReader(MyStream, System.Text.Encoding.Default)

            GetResponseHTML = StreamReader.ReadToEnd()
        Catch ex As Exception
            Throw New Exception("GetResponseHTML失敗|" & ResponseWord & "|" & ex.Message)
        End Try

        Return GetResponseHTML

    End Function

End Class

Public Class POSTResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String

    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
    End Sub
End Class
Public Class ResponseHTMLResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String

    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
    End Sub
End Class