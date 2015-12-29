﻿'------------------------------------------------------------------------------
' <auto-generated>
'     這段程式碼是由工具產生的。
'     執行階段版本:4.0.30319.17929
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
'原始程式碼已由 Microsoft.VSDesigner 自動產生，版本 4.0.30319.17929。
'
Namespace FacebookService

    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.17929"), _
     System.Diagnostics.DebuggerStepThroughAttribute(), _
     System.ComponentModel.DesignerCategoryAttribute("code"), _
     System.Web.Services.WebServiceBindingAttribute(Name:="FacebookServiceSoap", [Namespace]:="http://tempuri.org/")> _
    Partial Public Class FacebookService
        Inherits System.Web.Services.Protocols.SoapHttpClientProtocol

        Private PostnGetResponseOperationCompleted As System.Threading.SendOrPostCallback

        Private PostnGetResponseJSONOperationCompleted As System.Threading.SendOrPostCallback

        Private GetResponseHTMLOperationCompleted As System.Threading.SendOrPostCallback

        Private useDefaultCredentialsSetExplicitly As Boolean

        '''<remarks/>
        Public Sub New()
            MyBase.New()
            Me.Url = Global.FBWebService.My.MySettings.Default.FBWebService_FacebookService_FacebookService
            If (Me.IsLocalFileSystemWebService(Me.Url) = True) Then
                Me.UseDefaultCredentials = True
                Me.useDefaultCredentialsSetExplicitly = False
            Else
                Me.useDefaultCredentialsSetExplicitly = True
            End If
        End Sub

        Public Shadows Property Url() As String
            Get
                Return MyBase.Url
            End Get
            Set(value As String)
                If (((Me.IsLocalFileSystemWebService(MyBase.Url) = True) _
                            AndAlso (Me.useDefaultCredentialsSetExplicitly = False)) _
                            AndAlso (Me.IsLocalFileSystemWebService(Value) = False)) Then
                    MyBase.UseDefaultCredentials = False
                End If
                MyBase.Url = Value
            End Set
        End Property

        Public Shadows Property UseDefaultCredentials() As Boolean
            Get
                Return MyBase.UseDefaultCredentials
            End Get
            Set(value As Boolean)
                MyBase.UseDefaultCredentials = Value
                Me.useDefaultCredentialsSetExplicitly = True
            End Set
        End Property

        '''<remarks/>
        Public Event PostnGetResponseCompleted As PostnGetResponseCompletedEventHandler

        '''<remarks/>
        Public Event PostnGetResponseJSONCompleted As PostnGetResponseJSONCompletedEventHandler

        '''<remarks/>
        Public Event GetResponseHTMLCompleted As GetResponseHTMLCompletedEventHandler

        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/PostnGetResponse", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)> _
        Public Function PostnGetResponse(ByVal PostUrl As String, ByVal PostString As String) As String
            Dim results() As Object = Me.Invoke("PostnGetResponse", New Object() {PostUrl, PostString})
            Return CType(results(0), String)
        End Function

        '''<remarks/>
        Public Overloads Sub PostnGetResponseAsync(ByVal PostUrl As String, ByVal PostString As String)
            Me.PostnGetResponseAsync(PostUrl, PostString, Nothing)
        End Sub

        '''<remarks/>
        Public Overloads Sub PostnGetResponseAsync(ByVal PostUrl As String, ByVal PostString As String, ByVal userState As Object)
            If (Me.PostnGetResponseOperationCompleted Is Nothing) Then
                Me.PostnGetResponseOperationCompleted = AddressOf Me.OnPostnGetResponseOperationCompleted
            End If
            Me.InvokeAsync("PostnGetResponse", New Object() {PostUrl, PostString}, Me.PostnGetResponseOperationCompleted, userState)
        End Sub

        Private Sub OnPostnGetResponseOperationCompleted(ByVal arg As Object)
            If (Not (Me.PostnGetResponseCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg, System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent PostnGetResponseCompleted(Me, New PostnGetResponseCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub

        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/PostnGetResponseJSON", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)> _
        Public Function PostnGetResponseJSON(ByVal PostUrl As String, ByVal PostString As String) As String
            Dim results() As Object = Me.Invoke("PostnGetResponseJSON", New Object() {PostUrl, PostString})
            Return CType(results(0), String)
        End Function

        '''<remarks/>
        Public Overloads Sub PostnGetResponseJSONAsync(ByVal PostUrl As String, ByVal PostString As String)
            Me.PostnGetResponseJSONAsync(PostUrl, PostString, Nothing)
        End Sub

        '''<remarks/>
        Public Overloads Sub PostnGetResponseJSONAsync(ByVal PostUrl As String, ByVal PostString As String, ByVal userState As Object)
            If (Me.PostnGetResponseJSONOperationCompleted Is Nothing) Then
                Me.PostnGetResponseJSONOperationCompleted = AddressOf Me.OnPostnGetResponseJSONOperationCompleted
            End If
            Me.InvokeAsync("PostnGetResponseJSON", New Object() {PostUrl, PostString}, Me.PostnGetResponseJSONOperationCompleted, userState)
        End Sub

        Private Sub OnPostnGetResponseJSONOperationCompleted(ByVal arg As Object)
            If (Not (Me.PostnGetResponseJSONCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg, System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent PostnGetResponseJSONCompleted(Me, New PostnGetResponseJSONCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub

        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetResponseHTML", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)> _
        Public Function GetResponseHTML(ByVal ResponseWord As String) As String
            Dim results() As Object = Me.Invoke("GetResponseHTML", New Object() {ResponseWord})
            Return CType(results(0), String)
        End Function

        '''<remarks/>
        Public Overloads Sub GetResponseHTMLAsync(ByVal ResponseWord As String)
            Me.GetResponseHTMLAsync(ResponseWord, Nothing)
        End Sub

        '''<remarks/>
        Public Overloads Sub GetResponseHTMLAsync(ByVal ResponseWord As String, ByVal userState As Object)
            If (Me.GetResponseHTMLOperationCompleted Is Nothing) Then
                Me.GetResponseHTMLOperationCompleted = AddressOf Me.OnGetResponseHTMLOperationCompleted
            End If
            Me.InvokeAsync("GetResponseHTML", New Object() {ResponseWord}, Me.GetResponseHTMLOperationCompleted, userState)
        End Sub

        Private Sub OnGetResponseHTMLOperationCompleted(ByVal arg As Object)
            If (Not (Me.GetResponseHTMLCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg, System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent GetResponseHTMLCompleted(Me, New GetResponseHTMLCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub

        '''<remarks/>
        Public Shadows Sub CancelAsync(ByVal userState As Object)
            MyBase.CancelAsync(userState)
        End Sub

        Private Function IsLocalFileSystemWebService(ByVal url As String) As Boolean
            If ((url Is Nothing) _
                        OrElse (url Is String.Empty)) Then
                Return False
            End If
            Dim wsUri As System.Uri = New System.Uri(url)
            If ((wsUri.Port >= 1024) _
                        AndAlso (String.Compare(wsUri.Host, "localHost", System.StringComparison.OrdinalIgnoreCase) = 0)) Then
                Return True
            End If
            Return False
        End Function
    End Class

    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.17929")> _
    Public Delegate Sub PostnGetResponseCompletedEventHandler(ByVal sender As Object, ByVal e As PostnGetResponseCompletedEventArgs)

    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.17929"), _
     System.Diagnostics.DebuggerStepThroughAttribute(), _
     System.ComponentModel.DesignerCategoryAttribute("code")> _
    Partial Public Class PostnGetResponseCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs

        Private results() As Object

        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub

        '''<remarks/>
        Public ReadOnly Property Result() As String
            Get
                Me.RaiseExceptionIfNecessary()
                Return CType(Me.results(0), String)
            End Get
        End Property
    End Class

    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.17929")> _
    Public Delegate Sub PostnGetResponseJSONCompletedEventHandler(ByVal sender As Object, ByVal e As PostnGetResponseJSONCompletedEventArgs)

    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.17929"), _
     System.Diagnostics.DebuggerStepThroughAttribute(), _
     System.ComponentModel.DesignerCategoryAttribute("code")> _
    Partial Public Class PostnGetResponseJSONCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs

        Private results() As Object

        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub

        '''<remarks/>
        Public ReadOnly Property Result() As String
            Get
                Me.RaiseExceptionIfNecessary()
                Return CType(Me.results(0), String)
            End Get
        End Property
    End Class

    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.17929")> _
    Public Delegate Sub GetResponseHTMLCompletedEventHandler(ByVal sender As Object, ByVal e As GetResponseHTMLCompletedEventArgs)

    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.0.30319.17929"), _
     System.Diagnostics.DebuggerStepThroughAttribute(), _
     System.ComponentModel.DesignerCategoryAttribute("code")> _
    Partial Public Class GetResponseHTMLCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs

        Private results() As Object

        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub

        '''<remarks/>
        Public ReadOnly Property Result() As String
            Get
                Me.RaiseExceptionIfNecessary()
                Return CType(Me.results(0), String)
            End Get
        End Property
    End Class
End Namespace
