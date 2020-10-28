VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OAuth2Authenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule EmptyMethod
'@Folder("Web.Authenticators")
''
' OAuth2 Authenticator v3.0.5
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' OAuth2 authenticator
' (Currently using client credentials flow only)
'
' @class OAuth2Authenticator
' @implements IWebAuthenticator v4.*
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Private Type TOAuth2Authenticator
    ClientId As String
    ClientSecret As String
    Username As String
    Password As String
    TokenUrl As String
    TokenKey As String
    Token As String
End Type

Private this As TOAuth2Authenticator

Public Property Get ClientId() As String
    ClientId = this.ClientId
End Property

Public Property Let ClientId(ByVal value As String)
    this.ClientId = value
End Property

Public Property Get ClientSecret() As String
    ClientSecret = this.ClientSecret
End Property

Public Property Let ClientSecret(ByVal value As String)
    this.ClientSecret = value
End Property

Public Property Get Username() As String
    Username = this.Username
End Property

Public Property Let Username(ByVal value As String)
    this.Username = value
End Property

Public Property Get Password() As String
    Password = this.Password
End Property

Public Property Let Password(ByVal value As String)
    this.Password = value
End Property

Public Property Get TokenUrl() As String
    TokenUrl = this.TokenUrl
End Property

Public Property Let TokenUrl(ByVal value As String)
    this.TokenUrl = value
End Property

Public Property Get TokenKey() As String
    TokenKey = this.TokenKey
End Property

Public Property Let TokenKey(ByVal value As String)
    this.TokenKey = value
End Property

Public Property Get Token() As String
    Token = this.Token
End Property

Public Property Let Token(ByVal value As String)
    this.Token = value
End Property

Public Sub Setup(ByVal ClientId As String, ByVal ClientSecret As String, ByVal Username As String, ByVal Password As String)
    Me.ClientId = ClientId
    Me.ClientSecret = ClientSecret
    Me.Username = Username
    Me.Password = Password
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
' @internal
' @param {WebClient} Client
' @return {String}
''
Public Function GetToken(ByVal auth_Client As WebClient) As String
    ' Generate token request using client credentials flow
    ' This is currently setup using Salesforce's implementation
    ' TODO Generalize flow
    ' [Digging Deeper into OAuth 2.0 on Force.com](http://wiki.developerforce.com/page/Digging_Deeper_into_OAuth_2.0_at_Salesforce.com)
    
    On Error GoTo auth_Cleanup
    
    Dim auth_TokenClient As WebClient
    Dim auth_Request As New WebRequest
    Dim auth_Response As WebResponse
    
    ' Clone client (to avoid accidental interactions)
    Set auth_TokenClient = auth_Client.Clone
    Set auth_TokenClient.Authenticator = Nothing
    auth_TokenClient.BaseUrl = vbNullString
    
    ' Prepare token request
    auth_Request.Resource = Me.TokenUrl
    auth_Request.Method = WebMethod.HttpPost
    
    auth_Request.AddQuerystringParam "grant_type", "password"
    auth_Request.AddQuerystringParam "client_id", Me.ClientId
    auth_Request.AddQuerystringParam "client_secret", Me.ClientSecret
    auth_Request.AddQuerystringParam "username", Me.Username
    auth_Request.AddQuerystringParam "password", Me.Password
    
    Set auth_Response = auth_TokenClient.Execute(auth_Request)
    
    If auth_Response.StatusCode = WebStatusCode.Ok Then
        GetToken = auth_Response.Data.Item(Me.TokenKey)
    Else
        Err.Raise 11041 + vbObjectError, _
            Description:=auth_Response.StatusCode & ": " & auth_Response.Content
    End If
    
auth_Cleanup:

    Set auth_TokenClient = Nothing
    Set auth_Request = Nothing
    Set auth_Response = Nothing
    
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while retrieving token." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & IIf(Err.Number < 0, " (" & LCase$(Hex$(Err.Number)) & ")", vbNullString) & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "OAuth2Authenticator.GetToken", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "OAuth2Authenticator.GetToken", auth_ErrorDescription
    End If
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Sub Class_Initialize()
    Me.TokenKey = "access_token"
    Me.TokenUrl = vbNullString
End Sub


