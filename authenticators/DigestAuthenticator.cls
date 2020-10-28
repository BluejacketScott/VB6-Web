VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DigestAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule EmptyMethod
'@Folder("Web.Authenticators")
''
' Digest Authenticator v3.0.8
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Custom IWebAuthenticator for Digest Authentication
' http://en.wikipedia.org/wiki/Digest_access_authentication
'
' @class DigestAuthenticator
' @implements IWebAuthenticator v4.*
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const auth_Qop As String = "auth"
Private auth_pClientNonce As String

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Private Type TDigestAuthenticator
    Username As String
    Password As String
    Realm As String
    ServerNonce As String
    RequestCount As Long
    Opaque As String
End Type

Private this As TDigestAuthenticator

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

Public Property Get Realm() As String
    Realm = this.Realm
End Property

Public Property Let Realm(ByVal value As String)
    this.Realm = value
End Property

Public Property Get ServerNonce() As String
    ServerNonce = this.ServerNonce
End Property

Public Property Let ServerNonce(ByVal value As String)
    this.ServerNonce = value
End Property

Public Property Get RequestCount() As Long
    RequestCount = this.RequestCount
End Property

Public Property Let RequestCount(ByVal value As Long)
    this.RequestCount = value
End Property

Public Property Get Opaque() As String
    Opaque = this.Opaque
End Property

Public Property Let Opaque(ByVal value As String)
    this.Opaque = value
End Property

Public Property Get ClientNonce() As String
    If auth_pClientNonce = vbNullString Then
        auth_pClientNonce = WebHelpers.CreateNonce
    End If
    ClientNonce = auth_pClientNonce
End Property
Public Property Let ClientNonce(ByVal value As String)
    auth_pClientNonce = value
End Property

Public Property Get IsAuthenticated() As Boolean
    If ServerNonce <> vbNullString Then
        IsAuthenticated = True
    End If
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Setup authenticator
'
' @param {String} Username
' @param {String} Password
''
Public Sub Setup(ByVal Username As String, ByVal Password As String)
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
    If Me.IsAuthenticated Then
        Me.RequestCount = Me.RequestCount + 1
        Request.SetHeader "Authorization", CreateHeader(Client, Request)
    End If
End Sub

''
' Hook for taking action after request has been executed
'
' @param {WebClient} Client The client that executed request
' @param {WebRequest} Request The request that was just executed
' @param in|out {WebResponse} Response to request
''
Private Sub IWebAuthenticator_AfterExecute(ByVal Client As WebClient, ByVal Request As WebRequest, ByVal Response As WebResponse)
    If Response.StatusCode = 401 And Not Me.IsAuthenticated Then
        WebHelpers.LogDebug "Extract Authenticate and retry 401 request " & Client.GetFullUrl(Request), "Digest.AfterExecute"
        ExtractAuthenticateInformation Response
        
        Request.SetHeader "Authorization", CreateHeader(Client, Request)
        Response.Update Client.Execute(Request)
    End If
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
    ' http://curl.haxx.se/docs/manpage.html#--digest
    Curl = Curl & " --digest --user " & WebHelpers.PrepareTextForShell(Me.Username) & ":" & WebHelpers.PrepareTextForShell(Me.Password)
End Sub

''
' Create digest header for given Client and Request
'
' @internal
' @param {WebClient} Client
' @param {WebRequest} Request
' @return {String}
''
Public Function CreateHeader(ByVal Client As WebClient, ByVal Request As WebRequest) As String
    Dim auth_Uri As String
    auth_Uri = WebHelpers.GetUrlParts(Client.GetFullUrl(Request)).Item("Path")

    CreateHeader = "Digest " & _
        "username=""" & Me.Username & """, " & _
        "realm=""" & Me.Realm & """, " & _
        "nonce=""" & Me.ServerNonce & """, " & _
        "uri=""" & auth_Uri & """, " & _
        "qop=" & auth_Qop & ", " & _
        "nc=" & web_FormattedRequestCount & ", " & _
        "cnonce=""" & Me.ClientNonce & """, " & _
        "response=""" & web_CalculateResponse(Client, Request) & """, " & _
        "opaque=""" & Me.Opaque & """"
        
    WebHelpers.LogDebug CreateHeader, "DigestAuthenticator.CreateHeader"
End Function

''
' Extract authentication information from 401 response headers
'
' @internal
' @param {WebResponse} Response
''
Public Sub ExtractAuthenticateInformation(ByVal Response As WebResponse)
    Dim auth_Header As String
    Dim web_CrLf As String
    
    auth_Header = WebHelpers.FindInKeyValues(Response.Headers, "WWW-Authenticate")
    web_CrLf = Chr$(13) & Chr$(10)
    
    If auth_Header <> vbNullString And Left$(auth_Header, 6) = "Digest" Then
        Dim auth_Lines As Variant
        auth_Lines = Split(Mid$(auth_Header, 7), web_CrLf)
        
        Dim auth_i As Long
        Dim auth_Key As String
        Dim auth_Value As String
        For auth_i = LBound(auth_Lines) To UBound(auth_Lines)
            auth_Key = LCase$(Trim$(Mid$(auth_Lines(auth_i), 1, InStr(1, auth_Lines(auth_i), "=") - 1)))
            auth_Value = Trim$(Mid$(auth_Lines(auth_i), InStr(1, auth_Lines(auth_i), "=") + 1, Len(auth_Lines(auth_i))))
            
            ' Remove quotes and trailing comma
            auth_Value = Replace(auth_Value, """", vbNullString)
            If Right$(auth_Value, 1) = "," Then
                auth_Value = Left$(auth_Value, Len(auth_Value) - 1)
            End If
            
            ' Find realm, nonce, and opaque
            If auth_Key = "realm" Then Me.Realm = auth_Value
            If auth_Key = "nonce" Then Me.ServerNonce = auth_Value
            If auth_Key = "opaque" Then Me.Opaque = auth_Value
        Next auth_i
        
        WebHelpers.LogDebug "realm=" & Me.Realm & ", nonce=" & Me.ServerNonce & ", opaque=" & Me.Opaque, "DigestAuthenticator.ExtractAuthenticateInformation"
    End If
End Sub

' ============================================= '
' Private Methods
' ============================================= '

''
' Calculate digest response fro given Client and Request
'
' @internal
' @param {WebClient} Client
' @param {WebRequest} Request
' @return {String}
''
Private Function web_CalculateResponse(ByVal web_Client As WebClient, ByVal web_Request As WebRequest) As String
    '@Ignore UseMeaningfulName
    Dim auth_HA1 As String
    '@Ignore UseMeaningfulName
    Dim auth_HA2 As String
    Dim auth_Uri As String
    
    auth_Uri = WebHelpers.GetUrlParts(web_Client.GetFullUrl(web_Request)).Item("Path")
    auth_HA1 = web_CalculateHA1
    auth_HA2 = web_CalculateHA2(WebHelpers.MethodToName(web_Request.Method), auth_Uri)
    
    web_CalculateResponse = WebHelpers.MD5(auth_HA1 & ":" & Me.ServerNonce & ":" & web_FormattedRequestCount & ":" & Me.ClientNonce & ":" & auth_Qop & ":" & auth_HA2)
End Function

''
' Calculate HA1 portion of digest response
'
' @internal
' @return {String}
''
'@Ignore UseMeaningfulName
Private Function web_CalculateHA1() As String
    web_CalculateHA1 = WebHelpers.MD5(Me.Username & ":" & Me.Realm & ":" & Me.Password)
End Function

''
' Calculate HA1 portion of digest response
'
' @internal
' @return {String}
''
'@Ignore UseMeaningfulName
Private Function web_CalculateHA2(ByVal web_Method As String, ByVal web_Uri As String) As String
    web_CalculateHA2 = WebHelpers.MD5(web_Method & ":" & web_Uri)
End Function

''
' Pad request count to 8 places
'
' @internal
' @return {String}
''
Private Function web_FormattedRequestCount() As String
    web_FormattedRequestCount = Right("00000000" & Me.RequestCount, 8)
End Function


