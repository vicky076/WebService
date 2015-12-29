Imports Microsoft.VisualStudio.TestTools.UnitTesting.Web

Imports Microsoft.VisualStudio.TestTools.UnitTesting

Imports Lucky2012WebService



'''<summary>
'''這是 ServiceTest 的測試類別，應該包含
'''所有 ServiceTest 單元測試
'''</summary>
<TestClass()> _
Public Class ServiceTest


    Private testContextInstance As TestContext

    '''<summary>
    '''取得或設定提供目前測試回合的相關資訊與功能
    '''的測試內容。
    '''</summary>
    Public Property TestContext() As TestContext
        Get
            Return testContextInstance
        End Get
        Set(ByVal value As TestContext)
            testContextInstance = Value
        End Set
    End Property

#Region "其他測試屬性"
    '
    '您可以在撰寫測試時，使用下列的其他屬性:
    '
    '在執行類別中的第一項測試之前，先使用 ClassInitialize 執行程式碼
    '<ClassInitialize()>  _
    'Public Shared Sub MyClassInitialize(ByVal testContext As TestContext)
    'End Sub
    '
    '在執行類別中的所有測試之後，使用 ClassCleanup 執行程式碼
    '<ClassCleanup()>  _
    'Public Shared Sub MyClassCleanup()
    'End Sub
    '
    '在執行每一項測試之前，先使用 TestInitialize 執行程式碼
    '<TestInitialize()>  _
    'Public Sub MyTestInitialize()
    'End Sub
    '
    '在執行每一項測試之後，使用 TestCleanup 執行程式碼
    '<TestCleanup()>  _
    'Public Sub MyTestCleanup()
    'End Sub
    '
#End Region


    'TODO: 請確定 UrlToTest 屬性有為 ASP.NET 頁面指定 URL (例如，
    ' http://.../Default.aspx)。這是要在 Web 伺服器上執行單元測試時的必要項目，
    ' 無論您是要測試頁面、Web 服務或是 WCF 服務，都是如此。
    '''<summary>
    '''ConductProceduresCheckInsert 的測試
    '''</summary>
    <TestMethod(), _
     HostType("ASP.NET"), _
     AspNetDevelopmentServerHost("%PathToWebRoot%\Lucky2012WebService\Lucky2012WebService", "/"), _
     UrlToTest("http://localhost:1456/WebForm1.aspx")> _
    Public Sub ConductProceduresCheckInsertTest()
        Dim target As Service = New Service ' TODO: 初始化為適當值
        Dim ChoiceType As Integer = 2 ' TODO: 初始化為適當值
        Dim MyCardID As String = "1" ' TODO: 初始化為適當值
        Dim TradeSeq As String = "1" ' TODO: 初始化為適當值
        Dim CreateIp As String = "127.0.0.1" ' TODO: 初始化為適當值
        Dim ActId As String = "mcact10104" ' TODO: 初始化為適當值
        Dim expected As New ConductProceduresCheckInsertREsult ' TODO: 初始化為適當值
        expected.ReturnMsgNo = -11
        Dim actual As ConductProceduresCheckInsertREsult
        actual = target.ConductProceduresCheckInsert(ChoiceType, MyCardID, TradeSeq, CreateIp, ActId)
        Assert.AreEqual(expected.ReturnMsgNo, actual.ReturnMsgNo)
        'Assert.Inconclusive("驗證這個測試方法的正確性。")
    End Sub

    'TODO: 請確定 UrlToTest 屬性有為 ASP.NET 頁面指定 URL (例如，
    ' http://.../Default.aspx)。這是要在 Web 伺服器上執行單元測試時的必要項目，
    ' 無論您是要測試頁面、Web 服務或是 WCF 服務，都是如此。
    '''<summary>
    '''MyCardMemberForSavCheck 的測試
    '''</summary>
    <TestMethod(), _
     HostType("ASP.NET"), _
     AspNetDevelopmentServerHost("%PathToWebRoot%\Lucky2012WebService\Lucky2012WebService", "/"), _
     UrlToTest("http://localhost:1456/WebForm1.aspx")> _
    Public Sub MyCardMemberForSavCheckTest()
        Dim target As Service = New Service ' TODO: 初始化為適當值
        Dim ActId As String = "mcact10104" ' TODO: 初始化為適當值
        Dim MyCardID As String = "1" ' TODO: 初始化為適當值
        Dim expected As New MyCardMemberForSavCheckREsult ' TODO: 初始化為適當值
        expected.ReturnMsgNo = -4
        Dim actual As MyCardMemberForSavCheckREsult
        actual = target.MyCardMemberForSavCheck(ActId, MyCardID)
        Assert.AreEqual(expected.ReturnMsgNo, actual.ReturnMsgNo)
        'Assert.Inconclusive("驗證這個測試方法的正確性。")
    End Sub
End Class
