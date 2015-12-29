Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Net
Imports System.IO
Imports System.Security.Cryptography
Imports System.Xml
Imports System.Text
Imports System.Linq
Imports System.Collections
Imports System.Net.Dns
Imports Newtonsoft.Json
Imports System.Data.SqlClient


<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class fbservice
    Inherits System.Web.Services.WebService

    Dim myDbfnc As New MyCardConn.Dbfuction
    Dim Con As New SqlClient.SqlConnection()
    Dim MyCardDbwebservice As New dbwebservice.DbWebService
    Dim FacebookService As New FacebookService.FacebookService

    Dim app_id As String
    Dim APIKey As String
    Dim Secret As String
    Dim app_id_HK As String
    Dim APIKey_HK As String
    Dim Secret_HK As String
    Dim restServerUrl As String
    Sub New()
        app_id = My.Settings.app_id
        APIKey = My.Settings.APIKey
        Secret = My.Settings.Secret
        app_id_HK = My.Settings.app_id_HK
        APIKey_HK = My.Settings.APIKey_HK
        Secret_HK = My.Settings.Secret_HK
        restServerUrl = My.Settings.restServerUrl
    End Sub

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
                    Case 16
                        Con = myDbfnc.open16Db
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

    '20101118 qq 回傳字串以導頁至Facebook做登入動作
    '20110107 qq 修改使用oauth做登入，參數改傳form_id
    ''' <summary>
    ''' 回傳字串以導頁至Facebook做登入動作
    ''' </summary>
    ''' <param name="CallBackUrl">登入後導回的頁面</param>
    ''' <returns>ConnectString</returns>
    ''' <remarks>
    ''' CallBackUrl參數，例：http://apps.facebook.com/mycardpoints/
    ''' 回傳字串用Response.Write(ConnectString)方式以導頁至Facebook做登入動作
    ''' </remarks>
    <WebMethod()> _
    Public Function ConnectUrl(ByVal CallBackUrl As String) As String
        Dim LoginUrl As String = "https://graph.facebook.com/oauth/authorize"
        Dim ConnectString As String = "<script type='text/javascript'>top.location.href = '" + LoginUrl + "?client_id=" + app_id + "&redirect_uri=" + CallBackUrl + "&scope=user_birthday,email'</script>"

        Return ConnectString
    End Function

    '20101118 qq 轉給用戶Facebook Credits
    ''' <summary>
    ''' 轉給用戶Facebook Credits
    ''' </summary>
    ''' <param name="TransferValue">轉給用戶多少Facebook幣</param>
    ''' <param name="uid">用戶的Facebook UID</param>
    ''' <param name="discount">折扣</param>
    ''' <param name="txn_id">廠商交易序號</param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>轉給用戶Facebook Credits
    ''' 如果ReturnMsgNo=1表示FB給幣成功，ReturnMsg回傳FB OrderId，把FB OrderId記錄進DB中。
    ''' </remarks>
    <WebMethod()> _
    Public Function GetFBCredits(ByVal TransferValue As Integer, ByVal uid As String, ByVal discount As Decimal, ByVal txn_id As String) As FBCreditRE
        Dim NewDiscount As Decimal = 100 - discount
        '20111123 qq 要將discount轉成整數
        Dim CorrectDiscount As Integer
        Dim ReturnValue As New FBCreditRE
        Try
            CorrectDiscount = CInt(NewDiscount)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "GetFBCredits|Discount轉換失敗|" & CorrectDiscount, GetLocalhostIP)
            ReturnValue.ReturnMsgNo = -113
            ReturnValue.ReturnMsg = "Discount轉換失敗|" & CorrectDiscount
            ReturnValue.ReturnErrorCode = "FBWS1101"
            Return ReturnValue
        End Try
        Dim PosSig As String
        'PosSig = "api_key=" & APIKey & "discount=" & discount & "from_id=" & app_id & "method=facebook.payments.transferparams={""amount"":" & TransferValue & "}to_id=" & uid & Secret
        PosSig = "api_key=" & APIKey & "from_id=" & app_id & "method=facebook.payments.transferparams={""amount"":" & TransferValue & ",""txn_id"":""" & txn_id & """,""discount"":" & CorrectDiscount & "}to_id=" & uid & Secret
        Dim sig As String = GenMd5(PosSig)
        Dim PostString As String
        'PostString = "api_key=" & APIKey & "&discount=" & discount & "&from_id=" & app_id & "&method=facebook.payments.transfer&params={""amount"":" & TransferValue & "}&to_id=" & uid & "&sig=" & sig
        PostString = "api_key=" & APIKey & "&from_id=" & app_id & "&method=facebook.payments.transfer&params={""amount"":" & TransferValue & ",""txn_id"":""" & txn_id & """,""discount"":" & CorrectDiscount & "}&to_id=" & uid & "&sig=" & sig
        '記錄PostString
        MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), PostString, GetLocalhostIP)

        Dim ReturnXml As String = ""
        Try
            ReturnXml = FacebookService.PostnGetResponse(restServerUrl, PostString)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "PostnGetRespons失敗|FBWS1101|PostString|" & PostString & "|" & ex.Message, GetLocalhostIP)
            ReturnValue.ReturnMsgNo = -101
            ReturnValue.ReturnMsg = "PostnGetRespons失敗"
            ReturnValue.ReturnErrorCode = "FBWS1101"
            Return ReturnValue
        End Try
        '記錄ReturnXml
        MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), ReturnXml, GetLocalhostIP)

        Dim doc As New XmlDocument()
        doc.LoadXml(ReturnXml)

        Dim payments_transfer_response_v As XmlNodeList
        Dim payments_transfer_response As String = ""
        Dim error_code_v As XmlNodeList
        Dim error_code As String
        Dim error_msg_v As XmlNodeList
        Dim error_msg As String
        Try
            payments_transfer_response_v = doc.GetElementsByTagName("payments_transfer_response")
            payments_transfer_response = payments_transfer_response_v(0).InnerText
        Catch ex As Exception
            error_code_v = doc.GetElementsByTagName("error_code")
            error_code = error_code_v(0).InnerText
            error_msg_v = doc.GetElementsByTagName("error_msg")
            error_msg = error_msg_v(0).InnerText
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "交易序號|" & txn_id & "|取XML節點失敗|FBWS1102|error_code|" & error_code & "|error_msg|" & error_msg & "|" & ex.Message, GetLocalhostIP)
            ReturnValue.ReturnMsgNo = -102
            ReturnValue.ReturnMsg = "error_code|" & error_code & "|error_msg|" & error_msg
            ReturnValue.ReturnErrorCode = "FBWS1102"
            Return ReturnValue
        End Try
        ReturnValue.ReturnMsgNo = 1
        ReturnValue.ReturnMsg = payments_transfer_response
        Return ReturnValue
        'WriteLog("PostString:" & PostString & "|ReturnXml:" & ReturnXml)
        'lblMsg.Text = Session("FBname") & " 恭喜您獲得1枚Facebook幣。 OrderId:" & payments_transfer_response

    End Function

    '20110223 qq 轉給用戶Facebook Credits(香港用)
    ''' <summary>
    ''' 轉給用戶Facebook Credits(香港用)
    ''' </summary>
    ''' <param name="TransferValue">轉給用戶多少Facebook幣</param>
    ''' <param name="uid">用戶的Facebook UID</param>
    ''' <param name="discount">折扣</param>
    ''' <param name="txn_id">廠商交易序號</param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>轉給用戶Facebook Credits</remarks>
    <WebMethod()> _
    Public Function GetFBCreditsHK(ByVal TransferValue As Integer, ByVal uid As String, ByVal discount As Decimal, ByVal txn_id As String) As FBCreditRE
        Dim NewDiscount As Decimal = 100 - discount
        '20111123 qq 要將discount轉成整數
        Dim CorrectDiscount As Integer
        Dim ReturnValue As New FBCreditRE
        Try
            CorrectDiscount = CInt(NewDiscount)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "GetFBCreditsHK|Discount轉換失敗|" & CorrectDiscount, GetLocalhostIP)
            ReturnValue.ReturnMsgNo = -114
            ReturnValue.ReturnMsg = "Discount轉換失敗|" & CorrectDiscount
            ReturnValue.ReturnErrorCode = "FBWS1101"
            Return ReturnValue
        End Try
        Dim PosSig As String
        PosSig = "api_key=" & APIKey_HK & "from_id=" & app_id_HK & "method=facebook.payments.transferparams={""amount"":" & TransferValue & ",""txn_id"":""" & txn_id & """,""discount"":" & CorrectDiscount & "}to_id=" & uid & Secret_HK

        Dim sig As String = GenMd5(PosSig)
        Dim PostString As String
        PostString = "api_key=" & APIKey_HK & "&from_id=" & app_id_HK & "&method=facebook.payments.transfer&params={""amount"":" & TransferValue & ",""txn_id"":""" & txn_id & """,""discount"":" & CorrectDiscount & "}&to_id=" & uid & "&sig=" & sig

        '記錄PostString
        MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), PostString, GetLocalhostIP)

        'Dim ReturnValue As New FBCreditRE

        Dim ReturnXml As String = ""
        Try
            ReturnXml = FacebookService.PostnGetResponse(restServerUrl, PostString)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "PostnGetRespons失敗|FBWS1109|PostString|" & PostString & "|" & ex.Message, GetLocalhostIP)
            ReturnValue.ReturnMsgNo = -109
            ReturnValue.ReturnMsg = "PostnGetRespons失敗"
            ReturnValue.ReturnErrorCode = "FBWS1109"
            Return ReturnValue
        End Try
        '記錄ReturnXml
        MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), ReturnXml, GetLocalhostIP)

        Dim doc As New XmlDocument()
        doc.LoadXml(ReturnXml)

        Dim payments_transfer_response_v As XmlNodeList
        Dim payments_transfer_response As String = ""
        Dim error_code_v As XmlNodeList
        Dim error_code As String
        Dim error_msg_v As XmlNodeList
        Dim error_msg As String
        Try
            payments_transfer_response_v = doc.GetElementsByTagName("payments_transfer_response")
            payments_transfer_response = payments_transfer_response_v(0).InnerText
        Catch ex As Exception
            error_code_v = doc.GetElementsByTagName("error_code")
            error_code = error_code_v(0).InnerText
            error_msg_v = doc.GetElementsByTagName("error_msg")
            error_msg = error_msg_v(0).InnerText
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "交易序號|" & txn_id & "|取XML節點失敗|FBWS1110|error_code|" & error_code & "|error_msg|" & error_msg & "|" & ex.Message, GetLocalhostIP)
            ReturnValue.ReturnMsgNo = -110
            ReturnValue.ReturnMsg = "error_code|" & error_code & "|error_msg|" & error_msg
            ReturnValue.ReturnErrorCode = "FBWS1110"
            Return ReturnValue
        End Try
        ReturnValue.ReturnMsgNo = 1
        ReturnValue.ReturnMsg = payments_transfer_response
        Return ReturnValue

    End Function
    ''' <summary>
    ''' 20130618 紹安 新版加點 API 
    ''' sku,productDenomination,productCurrency應為一組由DB查詢的結果
    ''' </summary>
    ''' <param name="sku">API文件sku</param>
    ''' <param name="merchant">店家</param>
    ''' <param name="aPin">Mycard交易序號</param>
    ''' <param name="fbThirdPartyId">登入後呼叫GetThirdpartyid</param>
    ''' <param name="userPC">登入後呼叫GetCurrency</param>
    ''' <param name="productDenomination">API文件productDenomination</param>
    ''' <param name="productCurrency">API文件productCurrency</param>
    ''' <param name="ts">UNIX時間 GetUnixTime</param>
    ''' <returns>回傳交易結果 resultcode判斷</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function GetFbCreditAPAC(ByVal sku As String, ByVal merchant As String, ByVal aPin As String, ByVal fbThirdPartyId As String, ByVal userPC As String, ByVal productDenomination As String, ByVal productCurrency As String, ByVal ts As Integer) As CreditResult
        Dim TDS_Domain As String = My.Settings.TDS_Domain
        Dim TDS_function As String = My.Settings.TDS_function
        Dim TDS_APIKEY As String = My.Settings.TDS_APIKEY
        Dim fbThirdPartyAppId As String = My.Settings.app_id
        Dim TDS_APIsecret As String = My.Settings.TDS_Secret
        'If productCurrency.ToUpper() = "TWD" Or productCurrency.ToUpper() = "HKD" Then
        '    productDenomination = (CInt(productDenomination) * 100).ToString()
        'End If
        productDenomination = Decimal.Round(Decimal.Round(Decimal.Parse(productDenomination), 2) * 100, 0).ToString()
        'APAC_PostData("D6C0A63E6AB0427f9CA77F2A311B3914", "ABCDEF123456789", "merchantName", "aPinValue", "AB888989898", "100003777777777", "USD", "1000", "USD", "1280357665", "pAZI83IsUeI7m6rRzbIlbmiRf3cf2US3oppkbSGD01rJQn94XCdK04atUVojcGE")
        Return APAC_PostData(TDS_APIKEY, sku, merchant, aPin, fbThirdPartyId, fbThirdPartyAppId, userPC, productDenomination, productCurrency, ts, TDS_APIsecret)
    End Function
    ''' <summary>
    ''' 這Function沒用到-統一用台灣正式的APPID
    ''' 20130618 紹安 新版加點 API 
    ''' sku,productDenomination,productCurrency應為一組由DB查詢的結果
    ''' </summary>
    ''' <param name="sku">API文件sku</param>
    ''' <param name="merchant">店家</param>
    ''' <param name="aPin">Mycard交易序號</param>
    ''' <param name="fbThirdPartyId">登入後呼叫GetThirdpartyid</param>
    ''' <param name="userPC">登入後呼叫GetCurrency</param>
    ''' <param name="productDenomination">API文件productDenomination</param>
    ''' <param name="productCurrency">API文件productCurrency</param>
    ''' <param name="ts">UNIX時間 GetUnixTime</param>
    ''' <returns>回傳交易結果 resultcode判斷</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function GetFbCreditAPAC_HK(ByVal sku As String, ByVal merchant As String, ByVal aPin As String, ByVal fbThirdPartyId As String, ByVal userPC As String, ByVal productDenomination As String, ByVal productCurrency As String, ByVal ts As Integer) As CreditResult
        Dim TDS_Domain As String = My.Settings.TDS_Domain
        Dim TDS_function As String = My.Settings.TDS_function
        Dim TDS_APIKEY As String = My.Settings.TDS_APIKEY
        Dim fbThirdPartyAppId As String = My.Settings.app_id_HK
        Dim TDS_APIsecret As String = My.Settings.TDS_Secret
        Return APAC_PostData(TDS_APIKEY, sku, merchant, aPin, fbThirdPartyId, fbThirdPartyAppId, userPC, productDenomination, productCurrency, ts, TDS_APIsecret)
    End Function
    Function APAC_PostData(ByVal TDS_APIKEY As String, ByVal sku As String, ByVal merchant As String, ByVal aPin As String, ByVal fbThirdPartyId As String, ByVal fbThirdPartyAppId As String, ByVal userPC As String, ByVal productDenomination As String, ByVal productCurrency As String, ByVal ts As Integer, ByVal TDS_APIsecret As String) As CreditResult
        Dim dic As Dictionary(Of String, String) = New Dictionary(Of String, String)
        dic.Add("apiKey", TDS_APIKEY)
        dic.Add("sku", sku)
        dic.Add("merchant", merchant)
        dic.Add("aPin", aPin)
        dic.Add("fbThirdPartyId", fbThirdPartyId)
        dic.Add("fbThirdPartyAppId", fbThirdPartyAppId)
        dic.Add("userPC", userPC)
        dic.Add("productDenomination", productDenomination)
        dic.Add("productCurrency", productCurrency)
        dic.Add("ts", ts)

        Dim keys As New List(Of String)(dic.Keys)
        keys.Sort()

        Dim GetString As String = ""
        For Each s As String In keys
            GetString += s + "=" + dic(s)
        Next
        Dim sig As String = GetSHA1(My.Settings.TDS_function + GetString + TDS_APIsecret)
        Dim PostStr As String = ""
        For Each s As String In keys
            PostStr += s + "=" + dic(s) + "&"
        Next
        PostStr += "sig=" + sig

        '記錄PostString
        MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), PostStr, GetLocalhostIP)

        '記錄POST參數
        Dim FBBasicRe As New FBBasicRE
        Try
            FBBasicRe = CPSS_RepairCodeUpdate(aPin, PostStr)
        Catch ex As Exception
            FBBasicRe.ReturnMsgNo = -99
            FBBasicRe.ReturnMsg = ex.Message
        End Try
        If FBBasicRe.ReturnMsgNo <> 1 Then
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), aPin & "|" & PostStr & "|" & FBBasicRe.ReturnMsgNo & "," & FBBasicRe.ReturnMsg, GetLocalhostIP)
        End If

        Dim ReturnValue As New CreditResult
        Try
            GetString = FacebookService.PostnGetResponseJSON(My.Settings.TDS_Domain + My.Settings.TDS_function, PostStr)
        Catch ex As Exception
            ReturnValue.ReturnMsgNo = -207
            ReturnValue.ReturnMsg = "APAC_PostData失敗|FBWS1207|" & ex.Message
            Return ReturnValue
        End Try
        If GetString <> "" Then
            '記錄GetString
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "aPin|" & aPin & "|" & GetString, GetLocalhostIP)

            Dim UsersData As New FBreturnResult
            Try
                Dim jsonReader As JsonReader = New JsonTextReader(New StringReader(GetString))

                UsersData = JsonConvert.DeserializeObject(GetString, GetType(FBreturnResult))
            Catch ex As Exception
                'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "-106|JSON取值錯誤|" & ex.Message, GetLocalhostIP)
                ReturnValue.ReturnMsgNo = -208
                ReturnValue.ReturnMsg = "JSON取值錯誤|FBWS1208|" & ex.Message
                Return ReturnValue
            End Try

            If UsersData.resultCode <> "0" Then
                ReturnValue.ReturnMsgNo = CInt("-" & UsersData.resultCode)
                ReturnValue.ReturnMsg = UsersData.description
                Return ReturnValue
            End If

            ReturnValue.ReturnMsgNo = 1
            ReturnValue.ReturnMsg = "取得用戶資料成功"
            ReturnValue.FBreturnResult = UsersData

            Return ReturnValue
        Else
            'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "-107|取得用戶資料失敗|回傳空值", GetLocalhostIP)
            ReturnValue.ReturnMsgNo = -209
            ReturnValue.ReturnMsg = "APAC_PostData失敗|回傳空值|FBWS1209"
            Return ReturnValue
        End If
    End Function

    ''' <summary>
    ''' 20140711 qq GMG加點 API 增加活動代碼額外贈點 
    ''' sku,productDenomination,productCurrency應為一組由DB查詢的結果
    ''' </summary>
    ''' <param name="sku">API文件sku</param>
    ''' <param name="merchant">店家</param>
    ''' <param name="aPin">Mycard交易序號</param>
    ''' <param name="fbThirdPartyId">登入後呼叫GetThirdpartyid</param>
    ''' <param name="userPC">登入後呼叫GetCurrency</param>
    ''' <param name="productDenomination">API文件productDenomination</param>
    ''' <param name="productCurrency">API文件productCurrency</param>
    ''' <param name="ts">UNIX時間 GetUnixTime</param>
    ''' <returns>回傳交易結果 resultcode判斷</returns>
    ''' <param name="ProjNo">活動代碼</param>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function GetFbCreditAPACForBonus(ByVal sku As String, ByVal merchant As String, ByVal aPin As String, ByVal fbThirdPartyId As String, ByVal userPC As String, ByVal productDenomination As String, ByVal productCurrency As String, ByVal ts As Integer, ByVal ProjNo As String) As CreditResultForBonus
        Dim TDS_Domain As String = My.Settings.TDS_Domain
        Dim TDS_function As String = My.Settings.TDS_function
        Dim TDS_APIKEY As String = My.Settings.TDS_APIKEY
        Dim fbThirdPartyAppId As String = My.Settings.app_id
        Dim TDS_APIsecret As String = My.Settings.TDS_Secret

        '4捨5入成整數
        productDenomination = Decimal.Round(Decimal.Round(Decimal.Parse(productDenomination), 2) * 100, 0).ToString()

        Return APAC_PostDataForBonus(TDS_APIKEY, sku, merchant, aPin, fbThirdPartyId, fbThirdPartyAppId, userPC, productDenomination, productCurrency, ts, TDS_APIsecret, ProjNo)
    End Function
    Function APAC_PostDataForBonus(ByVal TDS_APIKEY As String, ByVal sku As String, ByVal merchant As String, ByVal aPin As String, ByVal fbThirdPartyId As String, ByVal fbThirdPartyAppId As String, ByVal userPC As String, ByVal productDenomination As String, ByVal productCurrency As String, ByVal ts As Integer, ByVal TDS_APIsecret As String, ByVal ProjNo As String) As CreditResultForBonus
        Dim dic As Dictionary(Of String, String) = New Dictionary(Of String, String)
        dic.Add("apiKey", TDS_APIKEY)
        dic.Add("sku", sku)
        dic.Add("merchant", merchant)
        dic.Add("aPin", aPin)
        dic.Add("fbThirdPartyId", fbThirdPartyId)
        dic.Add("fbThirdPartyAppId", fbThirdPartyAppId)
        dic.Add("userPC", userPC)
        dic.Add("productDenomination", productDenomination)
        dic.Add("productCurrency", productCurrency)
        dic.Add("ts", ts)

        Dim keys As New List(Of String)(dic.Keys)
        keys.Sort()

        Dim GetString As String = ""
        For Each s As String In keys
            GetString += s + "=" + dic(s)
        Next
        Dim sig As String = GetSHA1(My.Settings.TDS_function + GetString + TDS_APIsecret)
        Dim PostStr As String = ""
        For Each s As String In keys
            PostStr += s + "=" + dic(s) + "&"
        Next
        PostStr += "sig=" + sig

        '記錄PostString
        MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), aPin & "|" & PostStr, GetLocalhostIP)

        Dim RedeemString As String = aPin & "|" & PostStr

        '判斷活動日期
        Dim ProjNoSetDueDate As Date = CDate(My.Settings.ProjNoSetDueDate)
        If Now < ProjNoSetDueDate Then
            Dim ProjNoSetString As String = My.Settings.ProjNoSet
            Dim ProjNoSets() As String = Split(ProjNoSetString, ",")

            For index = 0 To ProjNoSets.Length - 1
                Dim ProjNos() As String = Split(ProjNoSets(index), "|")
                '判斷活動代號
                Dim ProjNoConfig As String = ProjNos(0)
                Dim ProjNoAmountConfig As String = ProjNos(1)
                Dim ProjNoSkuConfig As String = ProjNos(2)
                If ProjNo.ToLower.Trim = ProjNoConfig.ToLower.Trim Then
                    Dim aPin_b As String = aPin & "-B"

                    Dim PostStr_b As String = GenPostStr(TDS_APIKEY, ProjNoSkuConfig, merchant, aPin_b, fbThirdPartyId, fbThirdPartyAppId, userPC, ProjNoAmountConfig, productCurrency, ts, TDS_APIsecret)

                    '記錄PostString
                    MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), aPin_b & "|" & PostStr_b, GetLocalhostIP)

                    RedeemString = RedeemString & "," & aPin_b & "|" & PostStr_b
                End If
            Next

        End If

        Dim RedeemStrings() As String = Split(RedeemString, ",")

        For RedeemCount As Integer = 0 To RedeemStrings.Length - 1
            Dim RedeemInfo() As String = Split(RedeemStrings(RedeemCount), "|")
            Dim aPin_Redeem As String = RedeemInfo(0)
            Dim PostStr_Redeem As String = RedeemInfo(1)
            '記錄POST參數
            Dim FBBasicRe As New FBBasicRE
            Try
                FBBasicRe = CPSS_RepairCodeUpdate(aPin_Redeem, PostStr_Redeem)
            Catch ex As Exception
                FBBasicRe.ReturnMsgNo = -99
                FBBasicRe.ReturnMsg = ex.Message
            End Try
            If FBBasicRe.ReturnMsgNo <> 1 Then
                MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), aPin_Redeem & "|" & PostStr_Redeem & "|" & FBBasicRe.ReturnMsgNo & "," & FBBasicRe.ReturnMsg, GetLocalhostIP)
            End If
        Next

        Dim ReturnValue As New CreditResultForBonus
        Dim FBReturnMsgNo As Integer = 0
        Dim FBReturnMsg As String = ""

        For RedeemCount As Integer = 0 To RedeemStrings.Length - 1
            Dim RedeemInfo() As String = Split(RedeemStrings(RedeemCount), "|")
            Dim aPin_Redeem As String = RedeemInfo(0)
            Dim PostStr_Redeem As String = RedeemInfo(1)
            Try
                GetString = FacebookService.PostnGetResponseJSON(My.Settings.TDS_Domain + My.Settings.TDS_function, PostStr_Redeem)
            Catch ex As Exception
                FBReturnMsgNo = -207
                FBReturnMsg = "APAC_PostData失敗|FBWS1207|" & ex.Message
                GetString = ""
            End Try
            If GetString <> "" Then
                '記錄GetString
                MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "aPin|" & aPin_Redeem & "|" & GetString, GetLocalhostIP)

                Dim UsersData As New FBreturnResult
                Try
                    UsersData = JsonConvert.DeserializeObject(GetString, GetType(FBreturnResult))
                Catch ex As Exception
                    'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "-106|JSON取值錯誤|" & ex.Message, GetLocalhostIP)
                    UsersData.resultCode = "99"
                    UsersData.description = "JSON取值錯誤|FBWS1208|" & ex.Message
                    'Return ReturnValue
                End Try

                If UsersData.resultCode <> "0" Then
                    MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), My.Settings.TDS_Domain & My.Settings.TDS_function & "|" & PostStr_Redeem & "|" & UsersData.resultCode & "|" & UsersData.description, GetLocalhostIP)
                    If RedeemCount = 0 Then
                        ReturnValue.ReturnMsgNo = CInt("-" & UsersData.resultCode)
                        ReturnValue.ReturnMsg = UsersData.description
                    End If
                    'Return ReturnValue
                Else
                    If RedeemCount = 0 Then
                        ReturnValue.ReturnMsgNo = 1
                        ReturnValue.ReturnMsg = "取得用戶資料成功"
                        ReturnValue.FBreturnResult = UsersData
                    Else
                        ReturnValue.FBreturnResultBonus = UsersData
                    End If
                End If
            Else
                MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), My.Settings.TDS_Domain & My.Settings.TDS_function & "|" & PostStr_Redeem & "|" & FBReturnMsgNo & "|" & FBReturnMsg, GetLocalhostIP)
                If RedeemCount = 0 Then
                    ReturnValue.ReturnMsgNo = FBReturnMsgNo
                    ReturnValue.ReturnMsg = FBReturnMsg
                End If
            End If
        Next
        Return ReturnValue
    End Function

    Private Function GenPostStr(ByVal TDS_APIKEY As String, ByVal sku As String, ByVal merchant As String, ByVal aPin As String,
                                ByVal fbThirdpartyid As String, ByVal fbThirdPartyAppId As String, ByVal userPC As String, ByVal productDenomination As String,
                                ByVal productCurrency As String, ByVal ts As String, ByVal TDS_APIsecret As String) As String
        Dim dic As Dictionary(Of String, String) = New Dictionary(Of String, String)
        dic.Add("apiKey", TDS_APIKEY)
        dic.Add("sku", sku)
        dic.Add("merchant", merchant)
        dic.Add("aPin", aPin)
        dic.Add("fbThirdPartyId", fbThirdpartyid)
        dic.Add("fbThirdPartyAppId", fbThirdPartyAppId)
        dic.Add("userPC", userPC)
        dic.Add("productDenomination", productDenomination)
        dic.Add("productCurrency", productCurrency)
        dic.Add("ts", ts)

        Dim keys As New List(Of String)(dic.Keys)
        keys.Sort()

        Dim GetString As String = ""
        For Each s As String In keys
            GetString += s + "=" + dic(s)
        Next
        Dim sig As String = GetSHA1(My.Settings.TDS_function + GetString + TDS_APIsecret)
        Dim PostStr As String = ""
        For Each s As String In keys
            PostStr += s + "=" + dic(s) + "&"
        Next
        PostStr += "sig=" + sig

        Return PostStr
    End Function

    'Private Function PostnGetResponse(ByVal PostString As String) As String
    '    If Not String.IsNullOrEmpty(PostString) Then
    '        Dim restServerUrl As String = My.Settings.restServerUrl
    '        Dim parameterString As Byte() = Encoding.UTF8.GetBytes(PostString)

    '        Dim WebRequest As HttpWebRequest = HttpWebRequest.Create(restServerUrl.Trim)
    '        WebRequest.Method = "POST"
    '        WebRequest.ContentType = "application/x-www-form-urlencoded"
    '        WebRequest.ContentLength = parameterString.Length

    '        Dim newStream As Stream = WebRequest.GetRequestStream()
    '        newStream.Write(parameterString, 0, parameterString.Length)
    '        newStream.Close()

    '        Dim WebResponse As HttpWebResponse = CType(WebRequest.GetResponse(), HttpWebResponse)

    '        Dim sr As StreamReader = New StreamReader(WebResponse.GetResponseStream(), Encoding.Default)
    '        'Convert the stream to a string
    '        Dim ReturnString As String = sr.ReadToEnd()
    '        sr.Close()
    '        WebResponse.Close()

    '        Return ReturnString
    '    Else
    '        Return ""
    '    End If
    'End Function

    <WebMethod()> _
    Public Function GenMd5(ByVal strData As String) As String
        Dim md5Hasher As MD5 = MD5.Create()
        Dim data As Byte() = md5Hasher.ComputeHash(Encoding.ASCII.GetBytes(strData))

        Dim sBuilder As New StringBuilder
        Dim i As Integer
        For i = 0 To data.Length - 1
            sBuilder.Append(data(i).ToString("x2"))
        Next i
        Return sBuilder.ToString()
    End Function
    Public Function GetSHA1(ByVal strData As String) As String
        Dim sha As New SHA1Managed
        Dim strResult As String = ""
        sha.Initialize()
        Dim bytes As Byte() = sha.ComputeHash(Encoding.UTF8.GetBytes(strData))
        sha.Clear()
        For Each b As Byte In bytes
            strResult += b.ToString("x2")
        Next
        Return strResult.ToUpper()
    End Function

    '<WebMethod()> _
    'Public Function GetResponseHTML(ByVal ResponseWord As String) As String
    '    Try
    '        Dim oHttpRequest As HttpWebRequest = WebRequest.Create(ResponseWord)
    '        'ServicePointManager.ServerCertificateValidationCallback = AddressOf ValidateCertificate
    '        Dim ohttpResponse As HttpWebResponse = oHttpRequest.GetResponse
    '        Dim MyStream As IO.Stream
    '        MyStream = ohttpResponse.GetResponseStream
    '        Dim StreamReader As New IO.StreamReader(MyStream, System.Text.Encoding.Default)

    '        GetResponseHTML = StreamReader.ReadToEnd()
    '    Catch ex As Exception
    '        Throw New Exception("GetResponseHTML失敗|" & ResponseWord & "|" & ex.Message)
    '    End Try

    '    Return GetResponseHTML

    'End Function

    '20110107 qq 取得AccessToken
    ''' <summary>
    ''' 取得AccessToken
    ''' </summary>
    ''' <param name="CallBackUrl">登入時一樣的回傳頁面</param>
    ''' <param name="code">facebook登入成功後取得的參數code</param>
    ''' <returns>AccessToken字串</returns>
    ''' <remarks>
    ''' 如果回傳空值，請將用戶重新導回登入頁面
    ''' </remarks>
    <WebMethod()> _
    Public Function GetAccessToken(ByVal CallBackUrl As String, ByVal code As String) As String
        Dim access_tokenUrl As String = "https://graph.facebook.com/oauth/access_token"
        Dim ConnectString As String = access_tokenUrl + "?client_id=" + app_id + "&redirect_uri=" + CallBackUrl + "&client_secret=" + Secret + "&code=" + code
        Dim AccessToken As String
        Try
            AccessToken = FacebookService.GetResponseHTML(ConnectString)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), ex.Message, GetLocalhostIP)
            Return ""
        End Try

        Return AccessToken

    End Function

    '20110107 qq 取得用戶資料
    ''' <summary>
    ''' 取得用戶資料
    ''' </summary>
    ''' <param name="AccessToken">登入時取得的AccessToken</param>
    ''' <returns>ReturnMsgNo,ReturnMsg,ReturnDS</returns>
    ''' <remarks>
    ''' 回傳DataSet的欄位格式：id,name,birthday,gender,email,first_name,last_name
    ''' </remarks>
    <WebMethod()> _
    Public Function GetUserData(ByVal AccessToken As String) As FBReturnRE
        '{"id":"100000059727750","name":"\u9ec3\u7e6a\u8ae0","first_name":"\u7e6a\u8ae0","last_name":"\u9ec3","link":"http:\/\/www.facebook.com\/profile.php?id=100000059727750","birthday":"05\/26\/1979","location":{"id":"110922325599480","name":"Taichung, Taiwan"},"bio":"\u79ae\u8c8c\u5f88\u91cd\u8981~\r\n\u5e38\u8aaa: \u8acb. \u8b1d\u8b1d. \u5c0d\u4e0d\u8d77.~\r\n\u52dd\u904e\u8b80\u7684\u842c\u5377\u66f8~","education":[{"school":{"id":"110001145696056","name":"\u65b0\u6c11\u5546\u5de5"},"year":{"id":"112492302110039","name":"1999"},"type":"High School"},{"school":{"id":"110001145696056","name":"\u65b0\u6c11\u5546\u5de5"},"year":{"id":"112492302110039","name":"1999"},"type":"High School"}],"gender":"female","email":"leo891106@yahoo.com.tw","timezone":8,"locale":"zh_TW","verified":true,"updated_time":"2011-01-11T02:43:58+0000"}
        Dim GetString As String = "" '"{""id"":""100001387252134"",""name"":""\u9ec3\u7e6a\u8ae0"",""first_name"":""Tim"",""last_name"":""Jackson"",""link"":""http:\/\/www.facebook.com\/profile.php?id=100001387252134"",""gender"":""male"",""birthday"":""05\/26\/1979"",""email"":""qq@soft-world.com.tw"",""timezone"":8,""locale"":""zh_TW""}"
        Dim ReturnValue As New FBReturnRE
        Try
            GetString = FacebookService.GetResponseHTML("https://graph.facebook.com/me?access_token=" & AccessToken)
        Catch ex As Exception
            'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "-103|取得用戶資料失敗|" & ex.Message, GetLocalhostIP)
            ReturnValue.ReturnMsgNo = -103
            ReturnValue.ReturnMsg = "取得用戶資料失敗|" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS1103"
            Return ReturnValue
        End Try

        MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "GetUserData|" & GetString, GetLocalhostIP)
        If GetString <> "" Then
            Dim Users As New Users
            Dim UsersData As New Users
            Try
                Dim jsonReader As JsonReader = New JsonTextReader(New StringReader(GetString))
                UsersData = JsonConvert.DeserializeObject(GetString, GetType(Users))
            Catch ex As Exception
                'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "-108|JSON取值錯誤|" & ex.Message, GetLocalhostIP)
                ReturnValue.ReturnMsgNo = -108
                ReturnValue.ReturnMsg = "JSON取值錯誤|FBWS1108|" & ex.Message
                ReturnValue.ReturnErrorCode = "FBWS1108"
                Return ReturnValue
            End Try
            Dim UserDataSet As New DataSet
            UserDataSet.Tables.Add("User")
            UserDataSet.Tables("User").Columns.Add("id")
            UserDataSet.Tables("User").Columns.Add("name")
            UserDataSet.Tables("User").Columns.Add("gender")
            UserDataSet.Tables("User").Columns.Add("birthday")
            UserDataSet.Tables("User").Columns.Add("email")
            UserDataSet.Tables("User").Columns.Add("first_name")
            UserDataSet.Tables("User").Columns.Add("last_name")
            Dim datarow As DataRow = UserDataSet.Tables("User").NewRow
            datarow.Item(0) = IIf(IsNothing(UsersData.id), "", UsersData.id)
            datarow.Item(1) = IIf(IsNothing(UsersData.name), "", UsersData.name)
            datarow.Item(2) = IIf(IsNothing(UsersData.gender), "", UsersData.gender)
            datarow.Item(3) = IIf(IsNothing(UsersData.birthday), "", Replace(UsersData.birthday, "\", ""))
            datarow.Item(4) = IIf(IsNothing(UsersData.email), "", UsersData.email)
            datarow.Item(5) = IIf(IsNothing(UsersData.first_name), "", UsersData.first_name)
            datarow.Item(6) = IIf(IsNothing(UsersData.last_name), "", UsersData.last_name)
            UserDataSet.Tables("User").Rows.Add(datarow)
            ReturnValue.ReturnMsgNo = 1
            ReturnValue.ReturnMsg = "取得用戶資料成功"
            ReturnValue.ReturnDS = UserDataSet
            Return ReturnValue
        Else
            'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "-104|取得用戶資料失敗|回傳空值", GetLocalhostIP)
            ReturnValue.ReturnMsgNo = -104
            ReturnValue.ReturnMsg = "取得用戶資料失敗|回傳空值|FBWS1104"
            ReturnValue.ReturnErrorCode = "FBWS1104"
            Return ReturnValue
        End If

    End Function

    '20110107 qq 取得用戶uid
    ''' <summary>
    ''' 取得用戶uid
    ''' </summary>
    ''' <param name="AccessToken">登入時取得的AccessToken</param>
    ''' <returns>ReturnMsgNo,ReturnMsg,ReturnUID</returns>
    ''' <remarks>
    ''' </remarks>
    <WebMethod()> _
    Public Function GetUserUID(ByVal AccessToken As String) As UIDReturnRE
        Dim GetString As String = ""
        Dim ReturnValue As New UIDReturnRE
        Try
            GetString = FacebookService.GetResponseHTML("https://graph.facebook.com/me?access_token=" & AccessToken)
        Catch ex As Exception
            'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "-105|取得用戶資料失敗|" & ex.Message, GetLocalhostIP)
            ReturnValue.ReturnMsgNo = -105
            ReturnValue.ReturnMsg = "取得用戶資料失敗|FBWS1105|" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS1105"
            Return ReturnValue
        End Try
        If GetString <> "" Then
            Dim Users As New Users
            Dim UsersData As New Users
            Try
                Dim jsonReader As JsonReader = New JsonTextReader(New StringReader(GetString))
                UsersData = JsonConvert.DeserializeObject(GetString, GetType(Users))
            Catch ex As Exception
                'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "-106|JSON取值錯誤|" & ex.Message, GetLocalhostIP)
                ReturnValue.ReturnMsgNo = -106
                ReturnValue.ReturnMsg = "JSON取值錯誤|FBWS1106|" & ex.Message
                ReturnValue.ReturnErrorCode = "FBWS1106"
                Return ReturnValue
            End Try

            ReturnValue.ReturnMsgNo = 1
            ReturnValue.ReturnMsg = "取得用戶資料成功"
            ReturnValue.ReturnUID = UsersData.id
            Return ReturnValue
        Else
            'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "-107|取得用戶資料失敗|回傳空值", GetLocalhostIP)
            ReturnValue.ReturnMsgNo = -107
            ReturnValue.ReturnMsg = "取得用戶資料失敗|回傳空值|FBWS1107"
            ReturnValue.ReturnErrorCode = "FBWS1107"
            Return ReturnValue
        End If

    End Function

    '20110608 qq 取得錯誤訊息
    ''' <summary>
    ''' 取得錯誤訊息
    ''' </summary>
    ''' <param name="Type">1:中文, 其他:英文</param>
    ''' <param name="ErrorCode">錯誤代碼</param>
    ''' <returns>ReturnMsgNo,ReturnMsg</returns>
    ''' <remarks>
    ''' </remarks>
    <WebMethod()> _
    Public Function GetErrorMsg(ByVal Type As Integer, ByVal ErrorCode As String) As FBReturnRE
        ConnStr(16)
        Dim ReturnValue As New FBReturnRE
        Dim sqlcomm As String
        If ErrorCode = "" Then
            ReturnValue.ReturnMsgNo = -997
            If Type = 1 Then
                ReturnValue.ReturnMsg = "系統發生錯誤!(" & ErrorCode & ")"
            Else
                ReturnValue.ReturnMsg = "System Error!(" & ErrorCode & ")"
            End If
            Return ReturnValue
        Else
            sqlcomm = "Select ErrorCode,ErrorType,ErrorMsg,ErrorMsgEn,FaultType,SystemMsg From PointsBilling_Log.dbo.View_PointsBilling_ErrorType where ErrorCode like @ErrorCode"
        End If

        Dim com As New SqlCommand(sqlcomm, Con)
        com.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 10).Value = ErrorCode
        Dim da As New SqlDataAdapter(com)
        Dim ds As New DataSet()

        Try
            da.Fill(ds)
        Catch ex As Exception
            ReturnValue.ReturnMsgNo = -998
            If Type = 1 Then
                ReturnValue.ReturnMsg = "系統發生例外錯誤!(" & ErrorCode & ")"
            Else
                ReturnValue.ReturnMsg = "Exception Error!(" & ErrorCode & ")"
            End If
            Return ReturnValue
        Finally
            Con.Close()
        End Try


        If ds.Tables.Count > 0 Then
            If ds.Tables(0).Rows.Count > 0 Then
                ReturnValue.ReturnMsgNo = 1
                If Type = 1 Then
                    ReturnValue.ReturnMsg = IIf(IsDBNull(ds.Tables(0).Rows(0).Item("ErrorMsg")), "", ds.Tables(0).Rows(0).Item("ErrorMsg"))
                Else
                    ReturnValue.ReturnMsg = IIf(IsDBNull(ds.Tables(0).Rows(0).Item("ErrorMsgEn")), "", ds.Tables(0).Rows(0).Item("ErrorMsgEn"))
                End If
            Else
                ReturnValue.ReturnMsgNo = -999
                If Type = 1 Then
                    ReturnValue.ReturnMsg = "系統發生錯誤!(" & ErrorCode & ")"
                Else
                    ReturnValue.ReturnMsg = "System Error!(" & ErrorCode & ")"
                End If
            End If
        Else
            ReturnValue.ReturnMsgNo = -999
            If Type = 1 Then
                ReturnValue.ReturnMsg = "系統發生錯誤!(" & ErrorCode & ")"
            Else
                ReturnValue.ReturnMsg = "System Error!(" & ErrorCode & ")"
            End If
        End If

        Return ReturnValue

    End Function

    '20130617 紹安 取得UNIX timestamp
    <WebMethod()> _
    Public Function GetUnixTime() As Integer
        Return (DateTime.UtcNow - New DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds
    End Function

    '20110808 qq 解析FB POST過來的SignedRequest，解密並驗證得到相關資料
    ''' <summary>
    ''' 解析FB POST過來的SignedRequest，解密並驗證得到相關資料
    ''' </summary>
    ''' <param name="SignedRequest">FB POST過來的SignedRequest</param>
    ''' <returns>ReturnMsgNo,ReturnMsg,ReturnDS</returns>
    ''' <remarks>
    ''' </remarks>
    <WebMethod()> _
    Public Function GetUserDataFromSignedRequest(ByVal SignedRequest As String) As GetUserDataRE
        Dim ArreyStr() As String
        ArreyStr = Split(SignedRequest, ".")

        Dim Sig As String = ""
        Dim PayLoadStr As String = ""
        Dim Data As String = ""
        Dim ReturnValue As New GetUserDataRE
        Try
            Sig = ArreyStr(0)
            PayLoadStr = ArreyStr(1)

            '驗證簽章
            If CheckHMACSHA256(Sig, PayLoadStr) Then

                ArreyStr(1) = Replace(ArreyStr(1), "-", "+")
                ArreyStr(1) = Replace(ArreyStr(1), "_", "/")
                If ArreyStr(1) <> "" Then
                    If Len(ArreyStr(1)) Mod 4 <> 0 Then
                        Select Case Len(ArreyStr(1)) Mod 4
                            Case 1
                                ArreyStr(1) = ArreyStr(1) & "==="
                            Case 2
                                ArreyStr(1) = ArreyStr(1) & "=="
                            Case 3
                                ArreyStr(1) = ArreyStr(1) & "="
                        End Select
                    End If
                End If

                Data = UTF8Encoding.UTF8.GetString(Convert.FromBase64String(ArreyStr(1)))
                'Data = "{""user_id"":""1649933799""}"
                Dim Users As New Users
                Dim UsersData As New Users_signed
                Try
                    Dim jsonReader As JsonReader = New JsonTextReader(New StringReader(Data))
                    UsersData = JsonConvert.DeserializeObject(Data, GetType(Users_signed))
                Catch ex As Exception
                    'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "-108|JSON取值錯誤|" & ex.Message, GetLocalhostIP)
                    ReturnValue.ReturnMsgNo = -112
                    ReturnValue.ReturnMsg = "JSON取值錯誤|" & Data & "|" & ex.ToString
                    Return ReturnValue
                End Try
                Dim UserDataSet As New DataSet
                UserDataSet.Tables.Add("User")
                'UserDataSet.Tables("User").Columns.Add("user")
                UserDataSet.Tables("User").Columns.Add("algorithm")
                UserDataSet.Tables("User").Columns.Add("issued_at")
                UserDataSet.Tables("User").Columns.Add("user_id")
                UserDataSet.Tables("User").Columns.Add("oauth_token")
                UserDataSet.Tables("User").Columns.Add("expires")
                Dim datarow As DataRow = UserDataSet.Tables("User").NewRow
                'datarow.Item(0) = IIf(IsNothing(UsersData.user), "", UsersData.user)
                datarow.Item(0) = IIf(IsNothing(UsersData.algorithm), "", UsersData.algorithm)
                datarow.Item(1) = IIf(IsNothing(UsersData.issued_at), "", UsersData.issued_at)
                datarow.Item(2) = IIf(IsNothing(UsersData.user_id), "", UsersData.user_id)
                datarow.Item(3) = IIf(IsNothing(UsersData.oauth_token), "", UsersData.oauth_token)
                datarow.Item(4) = IIf(IsNothing(UsersData.expires), "", UsersData.expires)
                UserDataSet.Tables("User").Rows.Add(datarow)

                ReturnValue.ReturnMsgNo = 1
                ReturnValue.ReturnMsg = "取得資料成功"
                ReturnValue.ReturnDS = UserDataSet
                Return ReturnValue
            Else
                ReturnValue.ReturnMsgNo = -111
                ReturnValue.ReturnMsg = "簽章錯誤"
                Return ReturnValue
            End If
        Catch ex As Exception
            ReturnValue.ReturnMsgNo = -113
            ReturnValue.ReturnMsg = ex.Message
            Return ReturnValue
        End Try

    End Function

    Public Function CheckHMACSHA256(ByVal Sig As String, ByVal PayLoadStr As String) As Boolean
        Dim ComputedSig As String = ""
        Try
            Dim hmacsha256 As New HMACSHA256(UTF8Encoding.UTF8.GetBytes(Secret))
            Dim computedHash As Byte() = hmacsha256.ComputeHash(UTF8Encoding.UTF8.GetBytes(PayLoadStr))
            ComputedSig = Convert.ToBase64String(computedHash)
        Catch ex As Exception
            'Response.Write("<br>" & ex.Message)
            Return False
        End Try
        Sig = Replace(Sig, "-", "+")
        Sig = Replace(Sig, "_", "/")
        If Sig <> "" Then
            If Len(Sig) Mod 4 <> 0 Then
                Select Case Len(Sig) Mod 4
                    Case 1
                        Sig = Sig & "==="
                    Case 2
                        Sig = Sig & "=="
                    Case 3
                        Sig = Sig & "="
                End Select
            End If
        End If
        If Sig = ComputedSig Then
            'Response.Write("<br>簽章正確<br>")
            Return True
        Else
            'Response.Write("<br>簽章錯誤<br>")
            Return False
        End If
    End Function

    ''' <summary>
    ''' 取得AccessToken
    ''' </summary>
    ''' <param name="client_id"></param>
    ''' <param name="client_secret"></param>
    ''' <param name="CallBackUrl"></param>
    ''' <param name="code"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    <WebMethod()> _
    Public Function GetAccessTokenForOtherAPP(ByVal client_id As String, ByVal client_secret As String, ByVal CallBackUrl As String, ByVal code As String) As String
        Dim access_tokenUrl As String = "https://graph.facebook.com/oauth/access_token"
        Dim ConnectString As String = access_tokenUrl + "?client_id=" + client_id + "&redirect_uri=" + CallBackUrl + "&client_secret=" + client_secret + "&code=" + code
        Dim AccessToken As String
        Try
            AccessToken = FacebookService.GetResponseHTML(ConnectString)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "GetAccessToken|" & ex.Message, GetLocalhostIP)
            Return ""
        End Try

        Return AccessToken

    End Function
    '20130617 紹安 取得fbThirdPartyId
    ''' <summary>
    ''' 取得使用者目前的thirdpartyid
    ''' 參考 https://developers.facebook.com/blog/post/431/
    ''' </summary>
    ''' <param name="AccessToken">AccessToken</param>
    ''' <returns>third_party_id</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Function GetThirdpartyid(ByVal AccessToken As String) As fbThirdpartyid

        Dim GetString As String = ""
        Dim ReturnValue As New fbThirdpartyid
        Try
            GetString = FacebookService.GetResponseHTML("https://graph.facebook.com/me?fields=third_party_id&access_token=" & AccessToken)
        Catch ex As Exception
            ReturnValue.ReturnMsgNo = -201
            ReturnValue.ReturnMsg = "取得用戶資料失敗|FBWS1201|" & ex.Message
            ReturnValue.third_party_id = "FBWS1201"
            Return ReturnValue
        End Try
        If GetString <> "" Then
            Dim UsersData As New fbThirdpartyid
            Try
                Dim jsonReader As JsonReader = New JsonTextReader(New StringReader(GetString))
                UsersData = JsonConvert.DeserializeObject(GetString, GetType(fbThirdpartyid))
            Catch ex As Exception
                'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "-106|JSON取值錯誤|" & ex.Message, GetLocalhostIP)
                ReturnValue.ReturnMsgNo = -202
                ReturnValue.ReturnMsg = "JSON取值錯誤|FBWS1202|" & ex.Message
                ReturnValue.third_party_id = "FBWS1202"
                Return ReturnValue
            End Try

            ReturnValue.ReturnMsgNo = 1
            ReturnValue.ReturnMsg = "取得用戶資料成功"
            ReturnValue.third_party_id = UsersData.third_party_id
            Return ReturnValue
        Else
            'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "-107|取得用戶資料失敗|回傳空值", GetLocalhostIP)
            ReturnValue.ReturnMsgNo = -203
            ReturnValue.ReturnMsg = "取得用戶資料失敗|回傳空值|FBWS1203"
            ReturnValue.third_party_id = "FBWS1203"
            Return ReturnValue
        End If
    End Function
    '20130617 紹安
    ''' <summary>
    ''' 取得已登入使用者的地區資料
    ''' 參考 https://developers.facebook.com/docs/howtos/user-currency/
    ''' </summary>
    ''' <param name="AccessToken">AccessToken</param>
    ''' <returns>user_currency=使用者地區\ncurrency = 使用者地區的詳細資料</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Function GetCurrency(ByVal AccessToken As String) As UserCurrency

        Dim GetString As String = ""
        Dim ReturnValue As New UserCurrency
        Try
            GetString = FacebookService.GetResponseHTML("https://graph.facebook.com/me?fields=currency&access_token=" & AccessToken)
        Catch ex As Exception
            ReturnValue.ReturnMsgNo = -204
            ReturnValue.ReturnMsg = "取得用戶資料失敗|FBWS1204|" & ex.Message
            Return ReturnValue
        End Try
        If GetString <> "" Then
            Dim UsersData As New UserCurrency
            Try
                Dim jsonReader As JsonReader = New JsonTextReader(New StringReader(GetString))

                UsersData = JsonConvert.DeserializeObject(GetString, GetType(UserCurrency))
            Catch ex As Exception
                'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "-106|JSON取值錯誤|" & ex.Message, GetLocalhostIP)
                ReturnValue.ReturnMsgNo = -205
                ReturnValue.ReturnMsg = "JSON取值錯誤|FBWS1205|" & ex.Message
                Return ReturnValue
            End Try

            ReturnValue.ReturnMsgNo = 1
            ReturnValue.ReturnMsg = "取得用戶資料成功"
            ReturnValue.user_currency = UsersData.currency.user_currency
            ReturnValue.currency = UsersData.currency

            Return ReturnValue
        Else
            'MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "-107|取得用戶資料失敗|回傳空值", GetLocalhostIP)
            ReturnValue.ReturnMsgNo = -206
            ReturnValue.ReturnMsg = "取得用戶資料失敗|回傳空值|FBWS1206"
            Return ReturnValue
        End If
    End Function


    ''' <summary>
    ''' 發布塗鴉牆
    ''' </summary>
    ''' <param name="PostUrl">https://graph.facebook.com/me/feed</param>
    ''' <param name="PostString"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    <WebMethod()> _
    Public Function PostWall(ByVal PostUrl As String, ByVal PostString As String) As String
        Dim ReturnJson As String = ""
        Try
            ReturnJson = FacebookService.PostnGetResponse(PostUrl, PostString)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "PostWall失敗|PostUrl|" & PostUrl & "|PostString|" & PostString & "|" & ex.Message, GetLocalhostIP)
            Return ""
        End Try
        Return ReturnJson
    End Function


    Public Function GetLocalhostIP() As String
        Try
            Dim ipEntry As System.Net.IPHostEntry = GetHostByName(Environment.MachineName)
            Dim IpAddr As System.Net.IPAddress() = ipEntry.AddressList

            GetLocalhostIP = IpAddr(0).ToString
        Catch ex As Exception
            GetLocalhostIP = "無法取得IP"
        End Try
    End Function

    ''' <summary>
    ''' 2013-11-19 
    ''' 記錄FB API失敗的字串
    ''' </summary>
    ''' <param name="TradeSeq">交易序號</param>
    ''' <param name="RepairCode">POST字串</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function CPSS_RepairCodeUpdate(ByVal TradeSeq As String, ByVal RepairCode As String) As FBBasicRE
        ConnStr(15)

        Dim Cmd As New SqlClient.SqlCommand("MyCard_CPSaveService.dbo.CPSS_RepairCodeUpdate", Con)

        Cmd.CommandType = CommandType.StoredProcedure
        Cmd.CommandTimeout = 300

        Cmd.Parameters.Add("@TradeSeq", SqlDbType.VarChar, 50).Value = TradeSeq
        Cmd.Parameters.Add("@RepairCode", SqlDbType.VarChar, 500).Value = RepairCode

        Dim ReturnMsgNo As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsgNo", SqlDbType.Int)
        ReturnMsgNo.Direction = ParameterDirection.Output
        Dim ReturnMsg As SqlClient.SqlParameter = Cmd.Parameters.Add("@ReturnMsg", SqlDbType.NVarChar, 30)
        ReturnMsg.Direction = ParameterDirection.Output
        Dim ErrorCode As SqlClient.SqlParameter = Cmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 8)
        ErrorCode.Direction = ParameterDirection.Output
        Dim LogSn As SqlClient.SqlParameter = Cmd.Parameters.Add("@LogSn", SqlDbType.Int)
        LogSn.Direction = ParameterDirection.Output

        Dim ReturnValue As New FBBasicRE
        Try
            Con.Open()
            Cmd.ExecuteNonQuery()
            ReturnValue.ReturnMsg = IIf(IsDBNull(ReturnMsg.Value), "", ReturnMsg.Value)
            ReturnValue.ReturnMsgNo = IIf(IsDBNull(ReturnMsgNo.Value), 0, ReturnMsgNo.Value)
            ReturnValue.ReturnErrorCode = IIf(IsDBNull(ErrorCode.Value), "", ErrorCode.Value)
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Catch ex As Exception
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "999|CPSS_RepairCodeUpdate|TradeSeq|" & TradeSeq & "|RepairCode|" & RepairCode & "|" & ex.Message, "127.0.0.3")
            ReturnValue.ReturnMsgNo = "-999"
            ReturnValue.ReturnMsg = "系統發生錯誤|" & ex.Message
            ReturnValue.ReturnErrorCode = "FBWS1300"
            ReturnValue.ReturnLogSn = IIf(IsDBNull(LogSn.Value), 0, LogSn.Value)
        Finally
            Con.Close()
        End Try
        Return ReturnValue
    End Function

    ''' <summary>
    ''' 2013-11-19 
    ''' FB補儲重新發送
    ''' </summary>
    ''' <param name="PostStr">POST字串</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function RepairCodePost(ByVal PostStr As String) As CreditResult
        Dim ReturnValue As New CreditResult
        Dim GetString As String = ""
        Try
            GetString = FacebookService.PostnGetResponseJSON(My.Settings.TDS_Domain + My.Settings.TDS_function, PostStr)
        Catch ex As Exception
            ReturnValue.ReturnMsgNo = -301
            ReturnValue.ReturnMsg = "RepairCodePost失敗|FBWS1301|" & ex.Message
            Return ReturnValue
        End Try
        If GetString <> "" Then
            '記錄GetString
            MyCardDbwebservice.MyCardErrorLog(Server.MapPath("fbservice.asmx"), "RepairCodePost|" & GetString, GetLocalhostIP)

            Dim UsersData As New FBreturnResult
            Try
                Dim jsonReader As JsonReader = New JsonTextReader(New StringReader(GetString))

                UsersData = JsonConvert.DeserializeObject(GetString, GetType(FBreturnResult))
            Catch ex As Exception
                ReturnValue.ReturnMsgNo = -302
                ReturnValue.ReturnMsg = "RepairCodePost|JSON取值錯誤|FBWS1302|" & ex.Message
                Return ReturnValue
            End Try

            If UsersData.resultCode <> "0" Then
                ReturnValue.ReturnMsgNo = CInt("-" & UsersData.resultCode)
                ReturnValue.ReturnMsg = UsersData.description
                Return ReturnValue
            End If

            ReturnValue.ReturnMsgNo = 1
            ReturnValue.ReturnMsg = "成功"
            ReturnValue.FBreturnResult = UsersData

            Return ReturnValue
        Else
            ReturnValue.ReturnMsgNo = -303
            ReturnValue.ReturnMsg = "RepairCodePost失敗|回傳空值|FBWS1303"
            Return ReturnValue
        End If
    End Function
End Class

Public Class FBBasicRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnErrorCode As String
    Public ReturnLogSn As Integer
    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
        ReturnErrorCode = ""
        ReturnLogSn = 0
    End Sub
End Class
Public Class FBCreditRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnErrorCode As String
    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
        ReturnErrorCode = ""
    End Sub
End Class
Public Class FBReturnRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnDS As New DataSet
    Public ReturnErrorCode As String
    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
        ReturnErrorCode = ""
    End Sub
End Class
Public Class UIDReturnRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnUID As String
    Public ReturnErrorCode As String
    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
        ReturnUID = ""
        ReturnErrorCode = ""
    End Sub
End Class

Public Class Users
    Public id As String
    Public name As String
    Public gender As String
    Public birthday As String
    Public email As String
    Public first_name As String
    Public last_name As String
End Class

Public Class Users_signed
    'Public user As String
    Public algorithm As String
    Public issued_at As String
    Public user_id As String
    Public oauth_token As String
    Public expires As String
End Class

Public Class GetUserDataRE
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public ReturnDS As New DataSet
    Sub New()
        ReturnMsgNo = 0
        ReturnMsg = ""
    End Sub
End Class

Public Class fbThirdpartyid
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public third_party_id As String
End Class
Public Class currency
    Public user_currency As String
    Public currency_exchange As String
    Public currency_exchange_inverse As String
    Public usd_exchange As String
    Public usd_exchange_inverse As String
    Public currency_offset As String
End Class
Public Class UserCurrency
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public user_currency As String
    Public currency As currency
End Class
Public Class FBreturnResult
    Public resultCode As String
    Public description As String
    Public retryLater As String
    Public aPin As String
    Public action As String
    Public transactionId As String
    Public transactionTime As String
    Public currency As String
    Public amount As String
End Class
Public Class CreditResult
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public FBreturnResult As FBreturnResult
End Class
Public Class CreditResultForBonus
    Public ReturnMsgNo As Integer
    Public ReturnMsg As String
    Public FBreturnResult As FBreturnResult
    Public FBreturnResultBonus As FBreturnResult
End Class