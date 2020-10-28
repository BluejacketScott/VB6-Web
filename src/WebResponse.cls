VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@ModuleAttribute VB_Creatable, True
'@Folder("Web")
''
' WebResponse v4.1.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Wrapper for http/cURL responses that includes parsed Data based on WebRequest.ResponseFormat.
'
' Usage:
' ```VB.net
' Dim Response As WebResponse
' Set Response = Client.Execute(Request)
'
' If Response.StatusCode = WebStatusCode.Ok Then
'   ' Response.Headers, Response.Cookies
'   ' Response.Data -> Parsed Response.Content based on Request.ResponseFormat
'   ' Response.Body -> Raw response bytes
' Else
'   Debug.Print "Error: " & Response.StatusCode & " - " & Response.Content
' End If
' ```
'
' Errors:
' 11030 / 80042b16 / -2147210474 - Error creating from http
' 11031 / 80042b17 / -2147210473 - Error creating from cURL
' 11032 / 80042b18 / -2147210472 - Error extracting headers
'
' @class WebResponse
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private Type TWebResponse
    web_CrLf As String
    StatusCode As WebStatusCode
    StatusDescription As String
    Content As String
    Body As Variant
    Data As Object
    Headers As Collection
    Cookies As Collection
End Type

Private this As TWebResponse

Public Property Get web_CrLf() As String
    web_CrLf = this.web_CrLf
End Property

Public Property Let web_CrLf(ByVal value As String)
    this.web_CrLf = value
End Property

Public Property Get StatusCode() As WebStatusCode
    StatusCode = this.StatusCode
End Property

Public Property Let StatusCode(ByVal value As WebStatusCode)
    this.StatusCode = value
End Property

Public Property Get StatusDescription() As String
    StatusDescription = this.StatusDescription
End Property

Public Property Let StatusDescription(ByVal value As String)
    this.StatusDescription = value
End Property

Public Property Get Content() As String
    Content = this.Content
End Property

Public Property Let Content(ByVal value As String)
    this.Content = value
End Property

Public Property Get Body() As Variant
    Body = this.Body
End Property

Public Property Let Body(ByVal value As Variant)
    this.Body = value
End Property

Public Property Set Body(ByVal value As Variant)
    Set this.Body = value
End Property

Public Property Get Data() As Object
    Set Data = this.Data
End Property

Public Property Set Data(ByVal value As Object)
    Set this.Data = value
End Property

Public Property Get Headers() As Collection
    Set Headers = this.Headers
End Property

Public Property Set Headers(ByVal value As Collection)
    Set this.Headers = value
End Property

Public Property Get Cookies() As Collection
    Set Cookies = this.Cookies
End Property

Public Property Set Cookies(ByVal value As Collection)
    Set this.Cookies = value
End Property

Public Sub Update(ByVal Updated As WebResponse)
    Me.StatusCode = Updated.StatusCode
    Me.StatusDescription = Updated.StatusDescription
    Me.Content = Updated.Content
    Me.Body = Updated.Body
    Set Me.Headers = Updated.Headers
    Set Me.Cookies = Updated.Cookies
    Set Me.Data = Updated.Data
End Sub

''
' Create response from http
'
' @internal
' @method CreateFromHttp
' @param {WebClient} Client
' @param {WebRequest} Request
' @param {WinHttpRequest} Http
' @throws 11030 / 80042b16 / -2147210474 - Error creating from http
''
Public Sub CreateFromHttp(ByVal Client As WebClient, ByVal Request As WebRequest, ByVal Http As Object)
    On Error GoTo web_ErrorHandling

    Me.StatusCode = Http.Status
    Me.StatusDescription = Http.StatusText
    Me.Content = Http.ResponseText
    Me.Body = Http.ResponseBody

    web_LoadValues Http.GetAllResponseHeaders, Me.Content, Me.Body, Request

    Exit Sub

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred while creating response from http" & vbNewLine & _
        Err.Number & IIf(Err.Number < 0, " (" & LCase$(Hex$(Err.Number)) & ")", vbNullString) & ": " & Err.Description

    WebHelpers.LogError web_ErrorDescription, "WebResponse.CreateFromHttp", 11030 + vbObjectError
    Err.Raise 11030 + vbObjectError, "WebResponse.CreateFromHttp", web_ErrorDescription
End Sub

''
' Create response from cURL
'
' @internal
' @method CreateFromCurl
' @param {WebClient} Client
' @param {WebRequest} Request
' @param {String} Result
' @throws 11031 / 80042b17 / -2147210473 - Error creating from cURL
''
Public Sub CreateFromCurl(ByVal Client As WebClient, ByRef Request As WebRequest, ByVal Result As String)
    On Error GoTo web_ErrorHandling

    Dim web_Lines() As String

    web_Lines = Split(Result, web_CrLf)

    Me.StatusCode = web_ExtractStatusFromCurlResponse(web_Lines)
    Me.StatusDescription = web_ExtractStatusTextFromCurlResponse(web_Lines)
    Me.Content = web_ExtractResponseTextFromCurlResponse(web_Lines)
    Me.Body = WebHelpers.StringToAnsiBytes(Me.Content)

    web_LoadValues web_ExtractHeadersFromCurlResponse(web_Lines), Me.Content, Me.Body, Request

    Exit Sub

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred while creating response from cURL" & vbNewLine & _
        Err.Number & IIf(Err.Number < 0, " (" & LCase$(Hex$(Err.Number)) & ")", vbNullString) & ": " & Err.Description

    WebHelpers.LogError web_ErrorDescription, "WebResponse.CreateFromCurl", 11031 + vbObjectError
    Err.Raise 11031 + vbObjectError, "WebResponse.CreateFromCurl", web_ErrorDescription
End Sub

''
' Extract headers from response headers
'
' @internal
' @method ExtractHeaders
' @param {String} ResponseHeaders
' @return {Collection} Headers
' @throws 11032 / 80042b18 / -2147210472 - Error extracting headers
''
Public Function ExtractHeaders(ByVal ResponseHeaders As String) As Collection
    On Error GoTo web_ErrorHandling

    Dim web_Lines As Variant
    Dim web_i As Long
    Dim web_Headers As New Collection
    Dim web_Header As Dictionary
    Dim web_ColonPosition As Long
    Dim web_Multiline As Boolean

    web_Lines = Split(ResponseHeaders, web_CrLf)

    For web_i = LBound(web_Lines) To (UBound(web_Lines) + 1)
        If web_i > UBound(web_Lines) Then
            web_Headers.Add web_Header
        ElseIf web_Lines(web_i) <> vbNullString Then
            web_ColonPosition = InStr(1, web_Lines(web_i), ":")
            If web_ColonPosition = 0 And Not web_Header Is Nothing Then
                ' Assume part of multi-line header
                web_Multiline = True
            ElseIf web_Multiline Then
                ' Close out multi-line string
                web_Multiline = False
                web_Headers.Add web_Header
            ElseIf Not web_Header Is Nothing Then
                ' Add previous header
                web_Headers.Add web_Header
            End If

            If Not web_Multiline Then
                Set web_Header = WebHelpers.CreateKeyValue( _
                    Key:=Trim(Mid$(web_Lines(web_i), 1, web_ColonPosition - 1)), _
                    value:=Trim(Mid$(web_Lines(web_i), web_ColonPosition + 1, Len(web_Lines(web_i)))) _
                )
            Else
                web_Header.Item("Value") = web_Header.Item("Value") & web_CrLf & web_Lines(web_i)
            End If
        End If
    Next web_i

    Set ExtractHeaders = web_Headers
    Exit Function

web_ErrorHandling:

    Dim web_ErrorDescription As String
    web_ErrorDescription = "An error occurred while extracting headers" & vbNewLine & _
        Err.Number & IIf(Err.Number < 0, " (" & LCase$(Hex$(Err.Number)) & ")", vbNullString) & ": " & Err.Description

    WebHelpers.LogError web_ErrorDescription, "WebResponse.CreateFromCurl", 11032 + vbObjectError
    Err.Raise 11032 + vbObjectError, "WebResponse.CreateFromCurl", web_ErrorDescription
End Function

''
' Extract cookies from response headers
'
' @internal
' @method ExtractCookies
' @param {Collection} Headers
' @return {Collection} Cookies
''
Public Function ExtractCookies(ByVal Headers As Collection) As Collection
    Dim web_Header As Dictionary
    Dim web_Cookie As String
    Dim web_Key As String
    Dim web_Value As String
    Dim web_Cookies As New Collection

    For Each web_Header In Headers
        If web_Header.Item("Key") = "Set-Cookie" Then
            web_Cookie = web_Header.Item("Value")
            If InStr(1, web_Cookie, "=") > 0 Then
                web_Key = Mid$(web_Cookie, 1, InStr(1, web_Cookie, "=") - 1)
                web_Value = Mid$(web_Cookie, InStr(1, web_Cookie, "=") + 1, Len(web_Cookie))

                ' Ignore text after semi-colon
                If InStr(1, web_Value, ";") > 0 Then
                    web_Value = Mid$(web_Value, 1, InStr(1, web_Value, ";") - 1)
                End If

                ' Ignore surrounding quotes
                If Left$(web_Value, 1) = """" Then
                    web_Value = Mid$(web_Value, 2, Len(web_Value) - 2)
                End If

                web_Cookies.Add WebHelpers.CreateKeyValue(web_Key, WebHelpers.UrlDecode(web_Value, PlusAsSpace:=False, EncodingMode:=UrlEncodingMode.CookieUrlEncoding))
            Else
                WebHelpers.LogWarning _
                    "Unrecognized cookie format: " & web_Cookie, "WebResponse.ExtractCookies"
            End If
        End If
    Next web_Header

    Set ExtractCookies = web_Cookies
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Sub web_LoadValues(ByVal web_Headers As String, ByVal web_Content As String, ByVal web_Body As Variant, ByVal web_Request As WebRequest)
    ' Convert content to data by format
    If web_Request.ResponseFormat <> WebFormat.PlainText Then
        On Error Resume Next
        Set Me.Data = _
            WebHelpers.ParseByFormat(web_Content, web_Request.ResponseFormat, web_Request.CustomResponseFormat, web_Body)

        If Err.Number <> 0 Then
            WebHelpers.LogError Err.Description, Err.Source, Err.Number
            Err.Clear
        End If
        On Error GoTo 0
    End If

    ' Extract headers
    Set Me.Headers = ExtractHeaders(web_Headers)

    ' Extract cookies
    Set Me.Cookies = ExtractCookies(Me.Headers)
End Sub

Private Function web_ExtractStatusFromCurlResponse(ByRef web_CurlResponseLines() As String) As Long
    Dim web_StatusLineParts() As String

    web_StatusLineParts = Split(web_CurlResponseLines(web_FindStatusLine(web_CurlResponseLines)), " ")
    web_ExtractStatusFromCurlResponse = CLng(web_StatusLineParts(1))
End Function

Private Function web_ExtractStatusTextFromCurlResponse(ByRef web_CurlResponseLines() As String) As String
    Dim web_StatusLineParts() As String
    
    web_StatusLineParts = Split(web_CurlResponseLines(web_FindStatusLine(web_CurlResponseLines)), " ", 3)
    web_ExtractStatusTextFromCurlResponse = web_StatusLineParts(2)
End Function

Private Function web_ExtractHeadersFromCurlResponse(ByRef web_CurlResponseLines() As String) As String
    Dim web_StatusLineIndex As Long
    Dim web_BlankLineIndex As Long
    Dim web_HeaderLines() As String
    Dim web_WriteIndex As Long
    Dim web_ReadIndex As Long

    ' Find status line and blank line before body
    web_StatusLineIndex = web_FindStatusLine(web_CurlResponseLines)
    web_BlankLineIndex = web_FindBlankLine(web_CurlResponseLines)

    ' Extract headers string
    ReDim web_HeaderLines(0 To web_BlankLineIndex - 2 - web_StatusLineIndex)

    web_WriteIndex = 0
    For web_ReadIndex = (web_StatusLineIndex + 1) To web_BlankLineIndex - 1
        web_HeaderLines(web_WriteIndex) = web_CurlResponseLines(web_ReadIndex)
        web_WriteIndex = web_WriteIndex + 1
    Next web_ReadIndex

    web_ExtractHeadersFromCurlResponse = Join$(web_HeaderLines, web_CrLf)
End Function

Private Function web_ExtractResponseTextFromCurlResponse(ByRef web_CurlResponseLines() As String) As String
    Dim web_BlankLineIndex As Long
    Dim web_BodyLines() As String
    Dim web_WriteIndex As Long
    Dim web_ReadIndex As Long

    ' Find blank line before body
    web_BlankLineIndex = web_FindBlankLine(web_CurlResponseLines)

    ' Extract body string
    ReDim web_BodyLines(0 To UBound(web_CurlResponseLines) - web_BlankLineIndex - 1)

    web_WriteIndex = 0
    For web_ReadIndex = web_BlankLineIndex + 1 To UBound(web_CurlResponseLines)
        web_BodyLines(web_WriteIndex) = web_CurlResponseLines(web_ReadIndex)
        web_WriteIndex = web_WriteIndex + 1
    Next web_ReadIndex

    web_ExtractResponseTextFromCurlResponse = Join$(web_BodyLines, web_CrLf)
End Function

Private Function web_FindStatusLine(ByRef web_CurlResponseLines() As String) As Long
    ' Special case for cURL: 100 Continue is included before final status code
    ' -> ignore 100 and find final status code (next non-100 status code)
    For web_FindStatusLine = LBound(web_CurlResponseLines) To UBound(web_CurlResponseLines)
        If Trim$(web_CurlResponseLines(web_FindStatusLine)) <> vbNullString Then
            If Split(web_CurlResponseLines(web_FindStatusLine), " ")(1) <> "100" Then
                Exit Function
            End If
        End If
    Next web_FindStatusLine
End Function

Private Function web_FindBlankLine(ByRef web_CurlResponseLines() As String) As Long
    For web_FindBlankLine = (web_FindStatusLine(web_CurlResponseLines) + 1) To UBound(web_CurlResponseLines)
        If Trim$(web_CurlResponseLines(web_FindBlankLine)) = vbNullString Then
            Exit Function
        End If
    Next web_FindBlankLine
End Function

Private Sub Class_Initialize()
    web_CrLf = Chr$(13) & Chr$(10)

    Set Headers = New Collection
    Set Cookies = New Collection
End Sub


