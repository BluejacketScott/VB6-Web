VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OAuth1Authenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule EmptyMethod
'@Folder("Web.Authenticators")
''
' OAuth1 Authenticator v3.0.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' OAuth1 authenticator
'
' @class OAuth1Authenticator
' @implements IWebAuthenticator v4.*
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Const auth_SignatureMethod As String = "HMAC-SHA1"

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Private Type TOAuth1Authenticator
    ConsumerKey As String
    ConsumerSecret As String
    Token As String
    TokenSecret As String
    Realm As String
    Nonce As String
    Timestamp As String
End Type

Private this As TOAuth1Authenticator

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

Public Property Get TokenSecret() As String
    TokenSecret = this.TokenSecret
End Property

Public Property Let TokenSecret(ByVal value As String)
    this.TokenSecret = value
End Property

Public Property Get Realm() As String
    Realm = this.Realm
End Property

Public Property Let Realm(ByVal value As String)
    this.Realm = value
End Property

Public Property Get Nonce() As String
    Nonce = this.Nonce
End Property

Public Property Let Nonce(ByVal value As String)
    this.Nonce = value
End Property

Public Property Get Timestamp() As String
    Timestamp = this.Timestamp
End Property

Public Property Let Timestamp(ByVal value As String)
    this.Timestamp = value
End Property

Public Sub Setup(ByVal ConsumerKey As String, ByVal ConsumerSecret As String, _
    ByVal Token As String, ByVal TokenSecret As String, Optional ByVal Realm As String = vbNullString)
    
    Me.ConsumerKey = ConsumerKey
    Me.ConsumerSecret = ConsumerSecret
    Me.Token = Token
    Me.TokenSecret = TokenSecret
    Me.Realm = Realm
End Sub

Private Sub Class_Initialize()
    Me.Nonce = vbNullString
    Me.Timestamp = vbNullString
End Sub

''
' Hook for taking action before a request is executed
'
' @param {WebClient} Client The client that is about to execute the request
' @param in|out {WebRequest} Request The request about to be executed
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByVal Request As WebRequest)
    ' Add authorization header to request
    Request.SetHeader "Authorization", CreateHeader(Client, Request)
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
' Create header for given client and request
'
' @internal
' @param {WebClient} Client
' @param {WebRequest} Request
' @return {String}
''
Public Function CreateHeader(ByVal auth_Client As WebClient, ByVal auth_Request As WebRequest) As String
    Dim auth_Nonce As String
    Dim auth_Timestamp As String
    Dim auth_Base As String
    Dim auth_SigningKey As String
    Dim auth_Signature As String
    
    ' Load or create nonce and timestamp
    If Me.Nonce <> vbNullString Then
        auth_Nonce = Me.Nonce
    Else
        auth_Nonce = WebHelpers.CreateNonce()
    End If
    If Me.Timestamp <> vbNullString Then
        auth_Timestamp = Me.Timestamp
    Else
        auth_Timestamp = auth_CreateTimestamp
    End If
    
    ' Create needed parts of authorization header
    auth_Base = CreateBaseString(auth_Nonce, auth_Timestamp, auth_Client, auth_Request)
    auth_SigningKey = auth_CreateSigningKey()
    auth_Signature = CreateSignature(auth_Base, auth_SigningKey)
    
    ' Generate header
    CreateHeader = "OAuth "
    
    ' Add realm (if exists)
    If Me.Realm <> vbNullString Then
        CreateHeader = CreateHeader & "realm=""" & Me.Realm & """, "
    End If
    
    ' Construct header parts
    ' [OAuth Core 1.0 Revision A](http://oauth.net/core/1.0a/)
    CreateHeader = CreateHeader & "oauth_consumer_key=""" & Me.ConsumerKey & """, "
    CreateHeader = CreateHeader & "oauth_nonce=""" & auth_Nonce & """, "
    CreateHeader = CreateHeader & "oauth_signature=""" & WebHelpers.UrlEncode(auth_Signature) & """, "
    CreateHeader = CreateHeader & "oauth_signature_method=""" & auth_SignatureMethod & """, "
    CreateHeader = CreateHeader & "oauth_timestamp=""" & auth_Timestamp & """, "
    CreateHeader = CreateHeader & "oauth_token=""" & Me.Token & """, "
    CreateHeader = CreateHeader & "oauth_version=""" & "1.0" & """"
    
    CreateHeader = CreateHeader
End Function

''
' Create base string for given parameters
'
' @internal
' @param {String} Nonce
' @param {String} Timestamp
' @param {WebClient} Client
' @param {WebRequest} Request
' @return {String}
''
Public Function CreateBaseString(ByVal auth_Nonce As String, ByVal auth_Timestamp As String, ByVal auth_Client As WebClient, ByVal auth_Request As WebRequest) As String
    Dim auth_Base As String
    Dim auth_Parameters() As String
    
    ' Add and sort parameters
    auth_Parameters = Split(GetRequestParameters(auth_Client, auth_Request), "&")
    ReDim Preserve auth_Parameters(UBound(auth_Parameters) + 6)
    
    auth_Parameters(UBound(auth_Parameters) - 5) = "oauth_consumer_key=" & Me.ConsumerKey
    auth_Parameters(UBound(auth_Parameters) - 4) = "oauth_nonce=" & auth_Nonce
    auth_Parameters(UBound(auth_Parameters) - 3) = "oauth_signature_method=" & auth_SignatureMethod
    auth_Parameters(UBound(auth_Parameters) - 2) = "oauth_timestamp=" & auth_Timestamp
    auth_Parameters(UBound(auth_Parameters) - 1) = "oauth_token=" & Me.Token
    auth_Parameters(UBound(auth_Parameters)) = "oauth_version=1.0"
    
    auth_Parameters = SortParameters(auth_Parameters)
    auth_Base = Join(auth_Parameters, "&")
    
    CreateBaseString = WebHelpers.MethodToName(auth_Request.Method) & "&" & _
        WebHelpers.UrlEncode(GetRequestUrl(auth_Client, auth_Request)) & "&" & _
        WebHelpers.UrlEncode(auth_Base)
End Function

''
' Create signature with given parameters
'
' @internal
' @param {String} Base
' @param {String} SigningKey
' @return {String}
''
Public Function CreateSignature(ByRef auth_Base As String, ByRef auth_SigningKey As String) As String
    CreateSignature = WebHelpers.HMACSHA1(auth_Base, auth_SigningKey, "Base64")
End Function

''
' Create request url for given client and request
'
' @internal
' @param {WebClient} Client
' @param {WebRequest} Request
' @return {String}
''
Public Function GetRequestUrl(ByVal auth_Client As WebClient, ByVal auth_Request As WebRequest) As String
' From OAuth 1.0 Docs
' http://oauth.net/core/1.0/#anchor14
'
' The Signature Base String includes the request absolute URL, tying the signature to a specific endpoint.
' The URL used in the Signature Base String MUST include the scheme, authority, and path, and MUST exclude the query and fragment as defined by [RFC3986] section 3.
'
' If the absolute request URL is not available to the Service Provider (it is always available to the Consumer),
' it can be constructed by combining the scheme being used, the HTTP Host header, and the relative HTTP request URL.
' If the Host header is not available, the Service Provider SHOULD use the host name communicated to the Consumer in the documentation or other means.
'
' The Service Provider SHOULD document the form of URL used in the Signature Base String to avoid ambiguity due to URL normalization.
' Unless specified, URL scheme and authority MUST be lowercase and include the port number; http default port 80 and https default port 443 MUST be excluded.

    Dim auth_Parts As Dictionary
    Set auth_Parts = WebHelpers.GetUrlParts(auth_Client.GetFullUrl(auth_Request))
    
    ' Url scheme and authority MUST be lowercase
    GetRequestUrl = LCase(auth_Parts.Item("Protocol") & "://" & auth_Parts.Item("Host"))
    
    ' Include port (80 and 443 MUST be excluded)
    If auth_Parts.Item("Port") <> 80 And auth_Parts.Item("Port") <> 443 Then
        GetRequestUrl = GetRequestUrl & ":" & auth_Parts.Item("Port")
    End If
    
    ' Include path
    GetRequestUrl = GetRequestUrl + auth_Parts.Item("Path")
    
    ' MUST exclude query and fragment
End Function

''
' Create request parameters for given client and request
'
' @internal
' @param {WebClient} Client
' @param {WebRequest} Request
' @return {String}
''
Public Function GetRequestParameters(ByVal auth_Client As WebClient, ByVal auth_Request As WebRequest) As String
    Dim auth_Parts As Dictionary
    Set auth_Parts = WebHelpers.GetUrlParts(auth_Client.GetFullUrl(auth_Request))
    
    ' Replace + for spaces with %20
    GetRequestParameters = Replace(auth_Parts.Item("Querystring"), "+", "%20")
End Function

''
' Sort parameters (by value then key)
'
' @internal
' @param {Variant} Parameters
' @return {Variant}
''
Public Function SortParameters(ByRef auth_Parameters As Variant) As Variant
    ' Sort by key then value = sort by combined key-value
    ' (shouldn't be too many parameters, use naive selection sort
    Dim auth_Temp As String
    Dim auth_i As Long
    Dim auth_j As Long
    
    For auth_i = LBound(auth_Parameters) To UBound(auth_Parameters)
        For auth_j = auth_i To UBound(auth_Parameters)
            If auth_Parameters(auth_j) < auth_Parameters(auth_i) Then
                auth_Temp = auth_Parameters(auth_i)
                auth_Parameters(auth_i) = auth_Parameters(auth_j)
                auth_Parameters(auth_j) = auth_Temp
            End If
        Next auth_j
    Next auth_i
    
    SortParameters = auth_Parameters
End Function

' ============================================= '
' Private Methods
' ============================================= '

Private Function auth_CreateSigningKey() As String
    auth_CreateSigningKey = Me.ConsumerSecret & "&" & Me.TokenSecret
End Function

Private Function auth_CreateTimestamp() As String
    auth_CreateTimestamp = CStr(DateDiff("s", #1/1/1970#, WebHelpers.ConvertToUtc(Now)))
End Function


