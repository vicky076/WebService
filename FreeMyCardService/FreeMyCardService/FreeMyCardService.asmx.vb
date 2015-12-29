Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel

<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class FreeMyCardService
    Inherits System.Web.Services.WebService
    
    <WebMethod()> _
    Public Function ExecuteCancelPoint(ByVal FreeMyCardId As String, ByVal Trade_Seq As String, ByVal FreeTradeSeq As String) As Integer
        Dim wsFreeMyCard As New FreeMyCardAwardStatus.FreeMyCardAwardStatus
        Dim dsFreeCard As Integer = -1
        Try
            dsFreeCard = wsFreeMyCard.ExecuteCancelPoint(FreeMyCardId, Trade_Seq, FreeTradeSeq)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return dsFreeCard
    End Function

End Class