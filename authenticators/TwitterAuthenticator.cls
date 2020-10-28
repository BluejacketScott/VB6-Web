VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwitterAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule EmptyMethod
'@Folder("Web.Authenticators")
''
' Twitter Authenticator v3.0.5
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Custom IWebAuthenticator for application-only authentication in Twitter's V1.1 REST API
'
' Details:
' - https://dev.twitter.com/docs/auth/application-only-auth
' - https://github.com/VBA-tools/VBA-Web/wiki/Implementing-your-own-IWebAuthenticator
'
' Errors:
' 11041 / 80042b21 / -2147210463 - Error retreiving token
'
' @example
'   Dim Auth As New TwitterAuthenticator
'   Auth.Setup "Your consumer key", "Your consumer secret"
'
'   Set Client.Authenticator = Auth
'
' @class TwitterAuthenticator
' @implements IWebAuthenticator v4.*
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Private Type TTwitterAuthenticator
    ConsumerKey As String
    ConsumerSecret As String
    Token As String
End Type

Private this As TTwitterAuthenticator

Public Property Get ConsumerKey() As String
    ConsumerKey = this.ConsumerKey
End Property

Public Property Let ConsumerKey(ByVal value As String)
    this.ConsumerKey = value
End Property

Public Property Get ConsumerSecret() As String
    ConsumerSecret = this.ConsumerSecret
End Property

Public Property Let ConsumerSecret(ByVal value As String)
    this.ConsumerSecret = value
End Property

Public Property Get Token() As String
    Token = this.Token
End Property

Public Property Let Token(ByVal value As String)
    this.Token = value
End Property

Public Sub Setup(ByVal ConsumerKey As String, ByVal ConsumerSecret As String)
    Me.ConsumerKey = ConsumerKey
    Me.ConsumerSecret = ConsumerSecret
End Sub

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param in|out {WebRequest} Request The request about to be executed
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByVal Request As WebRequest)
    If Me.Token = vbNullString Then
        Me.Token = Me.GetToken(Client)
    End If

    Request.SetHeader "Authorization", "Bearer " & Me.Token
End Sub

''
' Hook for taking action after request has been executed
'
' @param {WebClient} Client The client that executed request
' @param {WebRequest} Request The request that was just executed
' @param in|out {WebResponse} Response to request
''
Private Sub IWebAuthenticator_AfterExecute(ByVal Client As WebClient, ByVal Request As WebRequest, ByVal Response As WebResponse)
    ' e.g. Handle 401 Unauthorized or other issues
End Sub

''
' Hook for updating http before send
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param in|out {WinHttpRequest} Http
''
Private Sub IWebAuthenticator_PrepareHttp(ByVal Client As WebClient, ByVal Request As WebRequest, ByVal Http As Object)
    ' e.g. Update option, headers, etc.
End Sub

''
' Hook for updating cURL before send
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param in|out {String} Curl
''
Private Sub IWebAuthenticator_PrepareCurl(ByVal Client As WebClient, ByVal Request As WebRequest, ByRef Curl As String)
    ' e.g. Add flags to cURL
End Sub

''
' Get token for given client
'
' @param {WebClient} Client
' @return {String}
''
Public Function GetToken(ByVal auth_Client As WebClient) As String
    On Error GoTo auth_Cleanup
    
    Dim auth_TokenClient As WebClient
    Dim auth_Request As New WebRequest
    Dim auth_Response As WebResponse
    
    ' Clone client (to avoid accidental interactions)
    Set auth_TokenClient = auth_Client.Clone
    Set auth_TokenClient.Authenticator = Nothing
    auth_TokenClient.BaseUrl = "https://api.twitter.com/"
    
    ' Prepare token request
    auth_Request.Resource = "oauth2/token"
    auth_Request.Method = WebMethod.HttpPost
    auth_Request.RequestFormat = WebFormat.FormUrlEncoded
    auth_Request.ResponseFormat = WebFormat.Json
    
    auth_Request.SetHeader "Authorization", _
        "Basic " & WebHelpers.Base64Encode(Me.ConsumerKey & ":" & Me.ConsumerSecret)
    auth_Request.AddBodyParameter "grant_type", "client_credentials"
    
    Set auth_Response = auth_TokenClient.Execute(auth_Request)
    
    If auth_Response.StatusCode = WebStatusCode.Ok Then
        GetToken = auth_Response.Data.Item("access_token")
    Else
        Err.Raise 11041 + vbObjectError, Description:=auth_Response.StatusCode & ": " & auth_Response.Content
    End If
    
auth_Cleanup:
    
    Set auth_TokenClient = Nothing
    Set auth_Request = Nothing
    Set auth_Response = Nothing
    
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while retrieving token" & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & IIf(Err.Number < 0, " (" & LCase$(Hex$(Err.Number)) & ")", vbNullString) & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "TwitterAuthenticator.GetToken", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "TwitterAuthenticator.GetToken", auth_ErrorDescription
    End If
End Function


