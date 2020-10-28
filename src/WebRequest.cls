VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@IgnoreModule UseMeaningfulName
'@ModuleAttribute VB_Creatable, True
'@Exposed
'@Folder("Web")
''
' WebRequest v4.1.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' `WebRequest` is used to create detailed requests
' (including formatting, querystrings, headers, cookies, and much more).
'
' Usage:
' ```VB.net
' Dim Request As New WebRequest
' Request.Resource = "users/{Id}"
'
' Request.Method = WebMethod.HttpPut
' Request.RequestFormat = WebFormat.UrlEncoded
' Request.ResponseFormat = WebFormat.Json
'
' Dim Body As New Dictionary
' Body.Add "name", "Tim"
' Body.Add "project", "VBA-Web"
' Set Request.Body = Body
'
' Request.AddUrlSegment "Id", 123
' Request.AddQuerystringParam "api_key", "abcd"
' Request.AddHeader "Authorization", "Token ..."
'
' ' -> PUT (Client.BaseUrl)users/123?api_key=abcd
' '    Authorization: Token ...
' '
' '    name=Tim&project=VBA-Web
' ```
'
' Errors:
' 11020 / 80042b0c / -2147210484 - Cannot add body parameter to non-Dictionary
'
' @class WebRequest
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Type TWebRequest
    Web_pRequestFormat As WebFormat
    Web_pResponseFormat As WebFormat
    Web_pCustomRequestFormat As String
    Web_pCustomResponseFormat As String
    Web_pBody As Variant
    Web_pConvertedBody As Variant
    Web_pContentType As String
    Web_pAccept As String
    Web_pAcceptEncoding As String
    Web_pContentLength As Long
    Web_pId As String
    Resource As String
    Method As WebMethod
    Headers As Collection
    QuerystringParams As Collection
    UrlSegments As Dictionary
    Cookies As Collection
    UserAgent As String
End Type

Private this As TWebRequest

Public Property Get Web_pRequestFormat() As WebFormat
    Web_pRequestFormat = this.Web_pRequestFormat
End Property

Public Property Let Web_pRequestFormat(ByVal value As WebFormat)
    this.Web_pRequestFormat = value
End Property

Public Property Get Web_pResponseFormat() As WebFormat
    Web_pResponseFormat = this.Web_pResponseFormat
End Property

Public Property Let Web_pResponseFormat(ByVal value As WebFormat)
    this.Web_pResponseFormat = value
End Property

Public Property Get Web_pCustomRequestFormat() As String
    Web_pCustomRequestFormat = this.Web_pCustomRequestFormat
End Property

Public Property Let Web_pCustomRequestFormat(ByVal value As String)
    this.Web_pCustomRequestFormat = value
End Property

Public Property Get Web_pCustomResponseFormat() As String
    Web_pCustomResponseFormat = this.Web_pCustomResponseFormat
End Property

Public Property Let Web_pCustomResponseFormat(ByVal value As String)
    this.Web_pCustomResponseFormat = value
End Property

Public Property Get Web_pBody() As Variant
    If Not IsEmpty(this.Web_pBody) Then
        If VarType(this.Web_pBody) = vbString Then
            Web_pBody = this.Web_pBody
        Else
            Set Web_pBody = this.Web_pBody
        End If
    End If
End Property

Public Property Let Web_pBody(ByVal value As Variant)
    this.Web_pBody = value
End Property

Public Property Set Web_pBody(ByVal value As Variant)
    Set this.Web_pBody = value
End Property

Public Property Get Web_pConvertedBody() As Variant
    If Not IsEmpty(this.Web_pConvertedBody) Then
        Set Web_pConvertedBody = this.Web_pConvertedBody
    End If
End Property

Public Property Let Web_pConvertedBody(ByVal value As Variant)
    this.Web_pConvertedBody = value
End Property

Public Property Set Web_pConvertedBody(ByVal value As Variant)
    Set this.Web_pConvertedBody = value
End Property

Public Property Get Web_pContentType() As String
    Web_pContentType = this.Web_pContentType
End Property

Public Property Let Web_pContentType(ByVal value As String)
    this.Web_pContentType = value
End Property

Public Property Get Web_pAccept() As String
    Web_pAccept = this.Web_pAccept
End Property

Public Property Let Web_pAccept(ByVal value As String)
    this.Web_pAccept = value
End Property

Public Property Get Web_pAcceptEncoding() As String
    Web_pAcceptEncoding = this.Web_pAcceptEncoding
End Property

Public Property Let Web_pAcceptEncoding(ByVal value As String)
    this.Web_pAcceptEncoding = value
End Property

Public Property Get Web_pContentLength() As Long
    Web_pContentLength = this.Web_pContentLength
End Property

Public Property Let Web_pContentLength(ByVal value As Long)
    this.Web_pContentLength = value
End Property

Public Property Get Web_pId() As String
    Web_pId = this.Web_pId
End Property

Public Property Let Web_pId(ByVal value As String)
    this.Web_pId = value
End Property

Public Property Get Resource() As String
    Resource = this.Resource
End Property

Public Property Let Resource(ByVal value As String)
    this.Resource = value
End Property

Public Property Get Method() As WebMethod
    Method = this.Method
End Property

Public Property Let Method(ByVal value As WebMethod)
    this.Method = value
End Property

Public Property Get Headers() As Collection
    Set Headers = this.Headers
End Property

Public Property Set Headers(ByVal value As Collection)
    Set this.Headers = value
End Property

Public Property Get QuerystringParams() As Collection
    Set QuerystringParams = this.QuerystringParams
End Property

Public Property Set QuerystringParams(ByVal value As Collection)
    Set this.QuerystringParams = value
End Property

Public Property Get UrlSegments() As Dictionary
    Set UrlSegments = this.UrlSegments
End Property

Public Property Set UrlSegments(ByVal value As Dictionary)
    Set this.UrlSegments = value
End Property

Public Property Get Cookies() As Collection
    Set Cookies = this.Cookies
End Property

Public Property Set Cookies(ByVal value As Collection)
    Set this.Cookies = value
End Property

Public Property Get UserAgent() As String
    UserAgent = this.UserAgent
End Property

Public Property Let UserAgent(ByVal value As String)
    this.UserAgent = value
End Property

Public Property Get Format() As WebFormat
    Format = RequestFormat
End Property
Public Property Let Format(ByVal value As WebFormat)
    Me.RequestFormat = value
    Me.ResponseFormat = value
End Property

''
' Set the format to use for converting the response `Body` to string and for the `Content-Type` header
'
' _Note_ If `WebFormat.Custom` is used, the [`CustomRequestFormat`](#/WebRequest/CustomRequestFormat) must be set.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Body = Array("A", "B", "C")
'
' Request.RequestFormat = WebFormat.Json
'
' ' -> (Header) Content-Type: application/json
' ' -> Convert Body to JSON string
' Request.Body ' = "["A","B","C"]"
' ```
'
' @property RequestFormat
' @type WebFormat
' @default WebFormat.Json
''
Public Property Get RequestFormat() As WebFormat
    RequestFormat = Web_pRequestFormat
End Property
Public Property Let RequestFormat(ByVal value As WebFormat)
    If value <> Web_pRequestFormat Then
        Web_pRequestFormat = value

        ' Clear cached converted body
        Web_pConvertedBody = Empty
    End If
End Property

''
' Set the format to use for converting the response `Content` to `Data` and for the `Accept` header
'
' _Note_ If `WebFormat.Custom` is used, the [`CustomResponseFormat`](#/WebRequest/CustomResponseFormat) must be set.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.ResponseFormat = WebFormat.Json
'
' ' -> (Header) Accept: application/json
'
' Dim Response As WebResponse
' ' ... from Execute
' Response.Content = "{""message"":""Howdy!""}"
'
' ' -> Parse Content to JSON Dictionary
' Debug.Print Response.Data("message") ' -> "Howdy!"
' ```
'
' @property ResponseFormat
' @type WebFormat
' @default WebFormat.Json
''
Public Property Get ResponseFormat() As WebFormat
    ResponseFormat = Web_pResponseFormat
End Property
Public Property Let ResponseFormat(ByVal value As WebFormat)
    If value <> Web_pResponseFormat Then
        Web_pResponseFormat = value

        ' Clear cached converted body
        Web_pConvertedBody = Empty
    End If
End Property

''
' Use converter registered with [`WebHelpers.RegisterConverter`](#/WebHelpers/RegisterConverter)
' to convert `Body` to string and set `Content-Type` header.
'
' (Automatically sets `RequestFormat` to `WebFormat.Custom`)
'
' @example
' ```VB.net
' WebHelpers.RegisterConverter "csv", "text/csv", "Module.ConvertToCsv", "Module.ParseCsv"
'
' Dim Request As New WebRequest
' Request.CustomRequestFormat = "csv"
'
' ' -> (Header) Content-Type: text/csv
' ' -> Body converted to string with Module.ConvertToCsv
' ```
'
' @property CustomRequestFormat
' @type String
''
Public Property Get CustomRequestFormat() As String
    CustomRequestFormat = Web_pCustomRequestFormat
End Property
Public Property Let CustomRequestFormat(ByVal value As String)
    If value <> Web_pCustomRequestFormat Then
        Web_pCustomRequestFormat = value

        ' Clear cached converted body
        Web_pConvertedBody = Empty

        If value <> vbNullString Then
            Web_pRequestFormat = WebFormat.Custom
        End If
    End If
End Property

''
' Use converter registered with [`WebHelpers.RegisterConverter`](#/WebHelpers/RegisterConverter)
' to convert the response `Content` to `Data` and set `Accept` header.
'
' (Automatically sets `ResponseFormat` to `WebFormat.Custom`)
'
' @example
' ```VB.net
' WebHelpers.RegisterConverter "csv", "text/csv", "Module.ConvertToCsv", "Module.ParseCsv"
'
' Dim Request As New WebRequest
' Request.CustomResponseFormat = "csv"
'
' ' -> (Header) Accept: text/csv
' ' -> WebResponse Content converted Data with Module.ParseCsv
' ```
'
' @property CustomResponseFormat
' @type String
''
Public Property Get CustomResponseFormat() As String
    CustomResponseFormat = Web_pCustomResponseFormat
End Property
Public Property Let CustomResponseFormat(ByVal value As String)
    If value <> Web_pCustomResponseFormat Then
        Web_pCustomResponseFormat = value

        ' Clear cached converted body
        Web_pConvertedBody = Empty

        If value <> vbNullString Then
            ResponseFormat = WebFormat.Custom
        End If
    End If
End Property

''
' Set automatically from `RequestFormat` or `CustomRequestFormat`,
' but can be overriden to set `Content-Type` header for request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.ContentType = "text/csv"
'
' ' -> (Header) Content-Type: text/csv
' ```
'
' @property ContentType
' @type String
' @default Media-type of request format
''
Public Property Get ContentType() As String
    If Web_pContentType <> vbNullString Then
        ContentType = Web_pContentType
    Else
        ContentType = WebHelpers.FormatToMediaType(Me.RequestFormat, Me.CustomRequestFormat)
    End If
End Property
Public Property Let ContentType(ByVal value As String)
    Web_pContentType = value
End Property

''
' Set automatically from `ResponseFormat` or `CustomResponseFormat`,
' but can be overriden to set `Accept` header for request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Accept = "text/csv"
'
' ' -> (Header) Accept: text/csv
' ```
'
' @property Accept
' @type String
' @default Media-type of response format
''
Public Property Get Accept() As String
    If Web_pAccept <> vbNullString Then
        Accept = Web_pAccept
    Else
        Accept = WebHelpers.FormatToMediaType(Me.ResponseFormat, Me.CustomResponseFormat)
    End If
End Property
Public Property Let Accept(ByVal value As String)
    Web_pAccept = value
End Property

''
' WinHTTP does not support decompression at this time (Jan. 2019).
' If not Accept-Encoding is passed to the server, [RFC 7231](https://tools.ietf.org/html/rfc7231#section-5.3.4)
' states that "any content-coding is considered acceptable by the user agent"
' -> Explicitly set Accept-Encoding
'
' cURL supports --compressed, which automatically decompresses gzip and other compressed responses
' -> If AcceptEncoding != "identity", enable --compressed flag
''
Public Property Get AcceptEncoding() As String
    If Web_pAcceptEncoding <> vbNullString Then
        AcceptEncoding = Web_pAcceptEncoding
    Else
        AcceptEncoding = "identity"
    End If
End Property
Public Property Let AcceptEncoding(ByVal value As String)
    Web_pAcceptEncoding = value
End Property

''
' Set automatically by length of `Body`,
' but can be overriden to set `Content-Length` header for request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.ContentLength = 200
'
' ' -> (Header) Content-Length: 200
' ```
'
' @property ContentLength
' @type Long
' @default Length of `Body`
''
Public Property Get ContentLength() As Long
    If Web_pContentLength >= 0 Then
        ContentLength = Web_pContentLength
    Else
        ContentLength = Len(Me.Body)
    End If
End Property
Public Property Let ContentLength(ByVal value As Long)
    Web_pContentLength = value
End Property

''
' - Get: Body value converted to string using `RequestFormat` or `CustomRequestFormat`
' - Let: Use `String` or `Array` for Body
' - Set: Use `Collection`, `Dictionary`, or `Object` for Body
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.RequestFormat = WebFormat.Json
'
' ' Let: String|Array
' Request.Body = "text"
' Debug.Print Request.Body ' -> "text"
'
' Request.Body = Array("A", "B", "C")
' Debug.Print Request.Body ' -> "["A","B","C"]"
'
' ' Set: Collection|Dictionary|Object
' Dim Body As Object
' Set Body = New Collection
' Body.Add "Howdy!"
' Set Request.Body = Body
' Debug.Print Request.Body ' -> "["Howdy!"]"
'
' Set Body = New Dictionary
' Body.Add "a", 123
' Body.Add "b", 456
' Set Request.Body = Body
' Debug.Print Request.Body ' -> "{"a":123,"b":456}"
' ```
'
' @property Body
' @type String|Array|Collection|Dictionary|Variant
''
Public Property Get Body() As Variant
    If Not IsEmpty(Web_pBody) Then
        If VarType(Web_pBody) = vbString Then
            Body = Web_pBody
        ElseIf IsEmpty(Web_pConvertedBody) Then
            ' Convert body and cache
            Set Body = WebHelpers.ConvertToFormat(Web_pBody, Me.RequestFormat, Me.CustomRequestFormat)
            Set Web_pConvertedBody = Body
        Else
            Set Body = Web_pConvertedBody
        End If
    End If
End Property
Public Property Let Body(ByVal value As Variant)
    Web_pConvertedBody = Empty
    Web_pBody = value
End Property
Public Property Set Body(ByVal value As Variant)
    Web_pConvertedBody = Empty
    Set Web_pBody = value
End Property

''
' Get `Resource` with Url Segments replaced and Querystring added.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Resource = "examples/{Id}"
' Request.AddUrlSegment "Id", 123
' Request.AddQuerystringParam "message", "Hello"
'
' Debug.Print Request.FormattedResource ' -> "examples/123?message=Hello"
' ```
'
' @property FormattedResource
' @type String
''
Public Property Get FormattedResource() As String
    Dim web_Segment As Variant
    Dim web_Encoding As UrlEncodingMode

    FormattedResource = Me.Resource

    ' Replace url segments
    For Each web_Segment In Me.UrlSegments.Keys
        FormattedResource = Replace(FormattedResource, "{" & web_Segment & "}", WebHelpers.UrlEncode(Me.UrlSegments.Item(web_Segment)))
    Next web_Segment

    ' Add querystring
    If Me.QuerystringParams.Count > 0 Then
        If InStr(FormattedResource, "?") <= 0 Then
            FormattedResource = FormattedResource & "?"
        Else
            FormattedResource = FormattedResource & "&"
        End If

        ' For querystrings, W3C defines form-urlencoded as the required encoding,
        ' but the treatment of space -> "+" (rather than "%20") can cause issues
        '
        ' If the request format is explicitly form-urlencoded, use FormUrlEncoding (space -> "+")
        ' otherwise, use subset of RFC 3986 and form-urlencoded that should work for both cases (space -> "%20")
        If Me.RequestFormat = WebFormat.FormUrlEncoded Then
            web_Encoding = UrlEncodingMode.FormUrlEncoding
        Else
            web_Encoding = UrlEncodingMode.QueryUrlEncoding
        End If
        FormattedResource = FormattedResource & WebHelpers.ConvertToUrlEncoded(Me.QuerystringParams, EncodingMode:=web_Encoding)
    End If
End Property

''
' @internal
' @property Id
' @type String
''
Public Property Get Id() As String
    If Web_pId = vbNullString Then: Web_pId = WebHelpers.CreateNonce
    Id = Web_pId
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Add header to be sent with request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.AddHeader "Authentication", "Bearer ..."
'
' ' -> (Header) Authorization: Bearer ...
' ```
'
' @method AddHeader
' @param {String} Key
' @param {Variant} Value
''
Public Sub AddHeader(ByVal Key As String, ByVal value As Variant)
    Me.Headers.Add WebHelpers.CreateKeyValue(Key, value)
End Sub

''
' Add/replace header to be sent with request.
' `SetHeader` should be used for headers that can only be included once with a request
' (e.g. Authorization, Content-Type, etc.).
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.AddHeader "Authorization", "A..."
' Request.AddHeader "Authorization", "B..."
'
' ' -> Headers:
' '    Authorization: A...
' '    Authorization: B...
'
' Request.SetHeader "Authorization", "C..."
'
' ' -> Headers:
' '    Authorization: C...
' ```
'
' @method SetHeader
' @param {String} Key
' @param {Variant} Value
''
Public Sub SetHeader(ByRef Key As String, ByRef value As Variant)
    WebHelpers.AddOrReplaceInKeyValues Me.Headers, Key, value
End Sub

''
' Url Segments are used to easily add dynamic values to `Resource`.
' Create a Url Segement in `Resource` with curly brackets and then
' replace with dynamic value with `AddUrlSegment`.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Dim User As String
' Dim Id As Long
'
' User = "Tim"
' Id = 123
'
' ' OK: Use string concatenation for dynamic values
' Request.Resource = User & "/messages/" & Id
'
' ' BETTER: Use Url Segments for dynamic values
' Request.Resource = "{User}/messages/{Id}"
' Request.AddUrlSegment "User", User
' Request.AddUrlSegment "Id", Id
'
' Debug.Print Request.FormattedResource ' > "Tim/messages/123"
' ```
'
' @method AddUrlSegment
' @param {String} Key
' @param {String} Value
''
Public Sub AddUrlSegment(ByRef Segment As String, ByVal value As Variant)
    Me.UrlSegments.Item(Segment) = value
End Sub

''
' Add querysting parameter to be used in `FormattedResource` for request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Resource = "messages"
' Request.AddQuerystringParam "from", "Tim"
'
' Request.FormattedResource ' = "messages?from=Tim"
' ```
'
' @method AddQuerystringParam
' @param {String} Key
' @param {Variant} Value
''
Public Sub AddQuerystringParam(ByVal Key As String, ByVal value As Variant)
    Me.QuerystringParams.Add WebHelpers.CreateKeyValue(Key, value)
End Sub

''
' Add cookie to be sent with request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.AddCookie "a", "abc"
' Request.AddCookie "b", 123
'
' ' -> (Header) Cookie: a=abc; b=123;
' ```
'
' @method AddCookie
' @param {String} Key
' @param {Variant} Value
''
Public Sub AddCookie(ByRef Key As String, ByRef value As Variant)
    Me.Cookies.Add WebHelpers.CreateKeyValue( _
        web_EncodeCookieName(Key), _
        WebHelpers.UrlEncode(value, EncodingMode:=UrlEncodingMode.CookieUrlEncoding) _
    )
End Sub

''
' Add `Key-Value` to `Body`.
' `Body` must be a `Dictionary` (if it's an `Array` or `Collection` an error is thrown)
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Format = WebFormat.Json
'
' Request.AddBodyParameter "a", 123
' Debug.Print Request.Body ' -> "{"a":123}"
'
' ' Can add parameters to existing Dictionary
' Dim Body As New Dictionary
' Body.Add "a", 123
'
' Set Request.Body = Body
' Request.AddBodyParameter "b", 456
'
' Debug.Print Request.Body ' -> "{"a":123,"b":456}"
' ```
'
' @method AddBodyParameter
' @param {Variant} Key
' @param {Variant} Value
' @throws 11020 / 80042b0c / -2147210484 - Cannot add body parameter to non-Dictionary
''
Public Sub AddBodyParameter(ByRef Key As Variant, ByRef value As Variant)
    If IsEmpty(Web_pBody) Then
        Set Web_pBody = New Dictionary
    ElseIf Not TypeOf Web_pBody Is Dictionary Then
        Dim web_ErrorDescription As String
        web_ErrorDescription = "Cannot add body parameter to non-Dictionary Body (existing Body must be of type Dictionary)"

        WebHelpers.LogError web_ErrorDescription, "WebRequest.AddBodyParameter", 11020 + vbObjectError
        Err.Raise 11020 + vbObjectError, "WebRequest.AddBodyParameter", web_ErrorDescription
    End If

    If IsObject(value) Then
        Set Web_pBody(Key) = value
    Else
        Web_pBody(Key) = value
    End If

    ' Clear cached converted body
    Web_pConvertedBody = Empty
End Sub

''
' Prepare request for execution
'
' @internal
' @method Prepare
''
Public Sub Prepare()
    ' Add/replace general headers for request
    SetHeader "User-Agent", Me.UserAgent
    SetHeader "Accept", Me.Accept
    SetHeader "Accept-Encoding", Me.AcceptEncoding
    If Me.Method <> WebMethod.HttpGet Or Me.ContentLength > 0 Then
        SetHeader "Content-Type", Me.ContentType
        SetHeader "Content-Length", CStr(Me.ContentLength)
    End If
End Sub

''
' Clone request
'
' @internal
' @method Clone
' @return {WebRequest}
''
Public Function Clone() As WebRequest
    Set Clone = New WebRequest

    ' Note: Clone underlying for properties with default values
    Clone.Resource = Me.Resource
    Clone.Method = Me.Method
    Clone.UserAgent = Me.UserAgent
    Clone.Accept = Web_pAccept
    Clone.ContentType = Web_pContentType
    Clone.ContentLength = Web_pContentLength
    Clone.RequestFormat = Me.RequestFormat
    Clone.ResponseFormat = Me.ResponseFormat
    Clone.CustomRequestFormat = Me.CustomRequestFormat
    Clone.CustomResponseFormat = Me.CustomResponseFormat

    Set Clone.Headers = WebHelpers.CloneCollection(Me.Headers)
    Set Clone.QuerystringParams = WebHelpers.CloneCollection(Me.QuerystringParams)
    Set Clone.UrlSegments = WebHelpers.CloneDictionary(Me.UrlSegments)
    Set Clone.Cookies = WebHelpers.CloneCollection(Me.Cookies)

    If IsObject(Web_pBody) Then
        Set Clone.Body = Web_pBody
    Else
        Clone.Body = Web_pBody
    End If
End Function

''
' Create WebRequest from options
'
' @method CreateFromOptions
' @param {Dictionary} Options
' @param {Collection} [Options.Headers] Collection of `KeyValue`
' @param {Collection} [Options.Cookies] Collection of `KeyValue`
' @param {Collection} [Options.QuerystringParams] Collection of `KeyValue`
' @param {Dictionary} [Options.UrlSegments]
''
Public Sub CreateFromOptions(ByVal Options As Dictionary)
    If Not Options Is Nothing Then
        If Options.Exists("Headers") Then
            Set Me.Headers = Options.Item("Headers")
        End If
        If Options.Exists("Cookies") Then
            Set Me.Cookies = Options.Item("Cookies")
        End If
        If Options.Exists("QuerystringParams") Then
            Set Me.QuerystringParams = Options.Item("QuerystringParams")
        End If
        If Options.Exists("UrlSegments") Then
            Set Me.UrlSegments = Options.Item("UrlSegments")
        End If
    End If
End Sub

' ============================================= '
' Private Functions
' ============================================= '

' Encode cookie name
'
' References:
' - RFC 6265 https://tools.ietf.org/html/rfc6265
Private Function web_EncodeCookieName(ByRef web_CookieName As Variant) As String
    Dim web_CookieVal As String
    Dim web_StringLen As Long

    web_CookieVal = CStr(web_CookieName)
    web_StringLen = Len(web_CookieVal)

    If web_StringLen > 0 Then
        Dim web_Result() As String
        Dim web_i As Long
        Dim web_CharCode As Long
        Dim web_Char As String
        ReDim web_Result(web_StringLen)

        ' ALPHA / DIGIT / "!" / "#" / "$" / "&" / "'" / "*" / "+" / "-" / "." / "^" / "_" / "`" / "|" / "~"
        ' Note: "%" is allowed in spec, but is currently excluded due to parsing issues

        ' Loop through string characters
        For web_i = 1 To web_StringLen
            ' Get character and ascii code
            web_Char = Mid$(web_CookieVal, web_i, 1)
            web_CharCode = Asc(web_Char)

            Select Case web_CharCode
                Case 65 To 90, 97 To 122
                    ' ALPHA
                    web_Result(web_i) = web_Char
                Case 48 To 57
                    ' DIGIT
                    web_Result(web_i) = web_Char
                Case 33, 35, 36, 38, 39, 42, 43, 45, 46, 94, 95, 96, 124, 126
                    ' "!" / "#" / "$" / "&" / "'" / "*" / "+" / "-" / "." / "^" / "_" / "`" / "|" / "~"
                    web_Result(web_i) = web_Char

                Case 0 To 15
                    web_Result(web_i) = "%0" & Hex(web_CharCode)
                Case Else
                    web_Result(web_i) = "%" & Hex(web_CharCode)
            End Select
        Next web_i

        web_EncodeCookieName = Join$(web_Result, vbNullString)
    End If
End Function

Private Sub Class_Initialize()
    ' Set default values
    Me.RequestFormat = WebFormat.Json
    Me.ResponseFormat = WebFormat.Json
    Me.UserAgent = WebUserAgent

    Set Me.Headers = New Collection
    Set Me.QuerystringParams = New Collection
    Set Me.UrlSegments = New Dictionary
    Set Me.Cookies = New Collection
    Me.ContentLength = -1
End Sub


