﻿'------------------------------------------------------------------------------
' <auto-generated>
'     這段程式碼是由工具產生的。
'     執行階段版本:4.0.30319.18063
'
'     對這個檔案所做的變更可能會造成錯誤的行為，而且如果重新產生程式碼，
'     變更將會遺失。
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml.Serialization

'
'原始程式碼已由 Microsoft.VSDesigner 自動產生，版本 4.0.30319.18063。
'
Namespace FreeMyCardAwardStatus
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.17929"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Web.Services.WebServiceBindingAttribute(Name:="FreeMyCardAwardStatusSoap", [Namespace]:="http://tempuri.org/")>  _
    Partial Public Class FreeMyCardAwardStatus
        Inherits System.Web.Services.Protocols.SoapHttpClientProtocol
        
        Private SearchPointOperationCompleted As System.Threading.SendOrPostCallback
        
        Private ExecuteCancelPointOperationCompleted As System.Threading.SendOrPostCallback
        
        Private useDefaultCredentialsSetExplicitly As Boolean
        
        '''<remarks/>
        Public Sub New()
            MyBase.New
            Me.Url = Global.FreeMyCardService.My.MySettings.Default.FreeMyCardService_FreeMyCardAwardStatus_FreeMyCardAwardStatus
            If (Me.IsLocalFileSystemWebService(Me.Url) = true) Then
                Me.UseDefaultCredentials = true
                Me.useDefaultCredentialsSetExplicitly = false
            Else
                Me.useDefaultCredentialsSetExplicitly = true
            End If
        End Sub
        
        Public Shadows Property Url() As String
            Get
                Return MyBase.Url
            End Get
            Set
                If (((Me.IsLocalFileSystemWebService(MyBase.Url) = true)  _
                            AndAlso (Me.useDefaultCredentialsSetExplicitly = false))  _
                            AndAlso (Me.IsLocalFileSystemWebService(value) = false)) Then
                    MyBase.UseDefaultCredentials = false
                End If
                MyBase.Url = value
            End Set
        End Property
        
        Public Shadows Property UseDefaultCredentials() As Boolean
            Get
                Return MyBase.UseDefaultCredentials
            End Get
            Set
                MyBase.UseDefaultCredentials = value
                Me.useDefaultCredentialsSetExplicitly = true
            End Set
        End Property
        
        '''<remarks/>
        Public Event SearchPointCompleted As SearchPointCompletedEventHandler
        
        '''<remarks/>
        Public Event ExecuteCancelPointCompleted As ExecuteCancelPointCompletedEventHandler
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/SearchPoint", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function SearchPoint(ByVal strMyID As String, ByVal strDealNo_MC As String, ByVal strDealNo_FMC As String) As Integer
            Dim results() As Object = Me.Invoke("SearchPoint", New Object() {strMyID, strDealNo_MC, strDealNo_FMC})
            Return CType(results(0),Integer)
        End Function
        
        '''<remarks/>
        Public Overloads Sub SearchPointAsync(ByVal strMyID As String, ByVal strDealNo_MC As String, ByVal strDealNo_FMC As String)
            Me.SearchPointAsync(strMyID, strDealNo_MC, strDealNo_FMC, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub SearchPointAsync(ByVal strMyID As String, ByVal strDealNo_MC As String, ByVal strDealNo_FMC As String, ByVal userState As Object)
            If (Me.SearchPointOperationCompleted Is Nothing) Then
                Me.SearchPointOperationCompleted = AddressOf Me.OnSearchPointOperationCompleted
            End If
            Me.InvokeAsync("SearchPoint", New Object() {strMyID, strDealNo_MC, strDealNo_FMC}, Me.SearchPointOperationCompleted, userState)
        End Sub
        
        Private Sub OnSearchPointOperationCompleted(ByVal arg As Object)
            If (Not (Me.SearchPointCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent SearchPointCompleted(Me, New SearchPointCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/ExecuteCancelPoint", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function ExecuteCancelPoint(ByVal strMyID As String, ByVal strDealNo_MC As String, ByVal strDealNo_FMC As String) As Integer
            Dim results() As Object = Me.Invoke("ExecuteCancelPoint", New Object() {strMyID, strDealNo_MC, strDealNo_FMC})
            Return CType(results(0),Integer)
        End Function
        
        '''<remarks/>
        Public Overloads Sub ExecuteCancelPointAsync(ByVal strMyID As String, ByVal strDealNo_MC As String, ByVal strDealNo_FMC As String)
            Me.ExecuteCancelPointAsync(strMyID, strDealNo_MC, strDealNo_FMC, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub ExecuteCancelPointAsync(ByVal strMyID As String, ByVal strDealNo_MC As String, ByVal strDealNo_FMC As String, ByVal userState As Object)
            If (Me.ExecuteCancelPointOperationCompleted Is Nothing) Then
                Me.ExecuteCancelPointOperationCompleted = AddressOf Me.OnExecuteCancelPointOperationCompleted
            End If
            Me.InvokeAsync("ExecuteCancelPoint", New Object() {strMyID, strDealNo_MC, strDealNo_FMC}, Me.ExecuteCancelPointOperationCompleted, userState)
        End Sub
        
        Private Sub OnExecuteCancelPointOperationCompleted(ByVal arg As Object)
            If (Not (Me.ExecuteCancelPointCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent ExecuteCancelPointCompleted(Me, New ExecuteCancelPointCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        Public Shadows Sub CancelAsync(ByVal userState As Object)
            MyBase.CancelAsync(userState)
        End Sub
        
        Private Function IsLocalFileSystemWebService(ByVal url As String) As Boolean
            If ((url Is Nothing)  _
                        OrElse (url Is String.Empty)) Then
                Return false
            End If
            Dim wsUri As System.Uri = New System.Uri(url)
            If ((wsUri.Port >= 1024)  _
                        AndAlso (String.Compare(wsUri.Host, "localHost", System.StringComparison.OrdinalIgnoreCase) = 0)) Then
                Return true
            End If
            Return false
        End Function
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.17929")>  _
    Public Delegate Sub SearchPointCompletedEventHandler(ByVal sender As Object, ByVal e As SearchPointCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.17929"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class SearchPointCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As Integer
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),Integer)
            End Get
        End Property
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.17929")>  _
    Public Delegate Sub ExecuteCancelPointCompletedEventHandler(ByVal sender As Object, ByVal e As ExecuteCancelPointCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.17929"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class ExecuteCancelPointCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As Integer
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),Integer)
            End Get
        End Property
    End Class
End Namespace
