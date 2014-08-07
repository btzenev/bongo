Imports System.IO
Imports System.Text
Imports System.Web.Mail
Imports System.Text.RegularExpressions
Imports System.Globalization

Public Module Common

    Public Const GlobalUser As String = "assent"
  Public Const GlobalRole As Integer = -12345
  Public Const AdminRole As Integer = -1
  Public Const GuestRole As Integer = 0
  Public Const EveryRole As Integer = -123

  Private mstr_langs As String
  Private mobj_langs As New Dictionary(Of String, Dictionary(Of String, String))
  Private mobj_config As New Dictionary(Of String, String)

  Public ReadOnly Property Config() As Dictionary(Of String, String)
    Get
      Return mobj_config
    End Get
  End Property

  Public Property Config(ByVal vstr_key As String) As String
    Get
      If mobj_config.ContainsKey(vstr_key) Then Return mobj_config(vstr_key)
      Return String.Empty
    End Get
    Set(ByVal value As String)
      If mobj_config.ContainsKey(vstr_key) Then
        mobj_config(vstr_key) = value
      Else
        mobj_config.Add(vstr_key, value)
      End If
    End Set
  End Property

  Public ReadOnly Property AppID() As String
    Get
      Return ToStr(Config("cms.id"), New Guid("00000000-0000-0000-0000-000000000000").ToString)
    End Get
  End Property

  Public ReadOnly Property Version() As String
    Get
      If Not Config.ContainsKey("cms.version") Then Config("cms.version") = My.Application.Info.Version.ToString
      Return Config("cms.version")
    End Get
  End Property

  Public ReadOnly Property AppPath() As String
    Get
      Return System.Web.HttpRuntime.AppDomainAppVirtualPath
    End Get
  End Property

#Region " HttpContext "

  Public ReadOnly Property HttpContext() As System.Web.HttpContext
    Get
      Return System.Web.HttpContext.Current
    End Get
  End Property

  Public Property HttpContext(ByVal vstr_name As String) As Object
    Get
      Return System.Web.HttpContext.Current.Items(vstr_name)
    End Get
    Set(ByVal vobj_value As Object)
      On Error Resume Next
      System.Web.HttpContext.Current.Items.Remove(vstr_name)
      System.Web.HttpContext.Current.Items.Add(vstr_name, vobj_value)
    End Set
  End Property

  Public ReadOnly Property Response() As System.Web.HttpResponse
    Get
      Return HttpContext.Response
    End Get
  End Property

  Public ReadOnly Property Server() As System.Web.HttpServerUtility
    Get
      Return HttpContext.Server
    End Get
  End Property

  Public ReadOnly Property Request() As System.Web.HttpRequest
    Get
      Return HttpContext.Request
    End Get
  End Property

  Public ReadOnly Property Application() As System.Web.HttpApplicationState
    Get
      Return HttpContext.Application
    End Get
  End Property

  Public ReadOnly Property Session() As System.Web.SessionState.HttpSessionState
    Get
      Return HttpContext.Session
    End Get
  End Property

  Public Property User() As System.Security.Principal.IPrincipal
    Get
      Return HttpContext.User
    End Get
    Set(ByVal Value As System.Security.Principal.IPrincipal)
      HttpContext.User = Value
    End Set
  End Property

#End Region

  Public ReadOnly Property Debug() As Boolean
    Get
      Return HttpContext.IsDebuggingEnabled
    End Get
  End Property

  Public ReadOnly Property RawUrl() As String
    Get
      On Error Resume Next
      Return Url.PathAndQuery
    End Get
  End Property

  Public ReadOnly Property FullUrl() As String
    Get
      On Error Resume Next
      Return Url.ToString
    End Get
  End Property

  Public ReadOnly Property Url() As Uri
    Get
      On Error Resume Next
      Dim obj_uri As Uri = ToObj(Of Uri)(HttpContext("cms.url"))
      If obj_uri Is Nothing Then
        Dim str_url As String = HttpContext.Current.Request.RawUrl.ToLower
        Dim int_404 As Integer = str_url.IndexOf("404;")
        If int_404 >= 0 Then
          int_404 = str_url.IndexOf("://", int_404) + 3
          int_404 = str_url.IndexOf("/", int_404)
          str_url = str_url.Substring(int_404)
        End If
        str_url = Request.Url.Scheme & "://" & Request.Url.Host & str_url
        obj_uri = New Uri(str_url)
        HttpContext("cms.url") = obj_uri
      End If
      Return obj_uri
    End Get
  End Property

  Public ReadOnly Property Languages() As Dictionary(Of String, Dictionary(Of String, String))
    Get
      On Error Resume Next
      If MultiLingual Then
        Dim str_file As String = MapPath("/languages.xml")
        If IO.File.Exists(str_file) Then
          Dim str_crc As String = GetFileCRC(str_file)
          If mstr_langs <> str_crc Then
            mstr_langs = str_crc
            mobj_langs = New Dictionary(Of String, Dictionary(Of String, String))
            Dim obj_xml As New Xml.XmlDocument()
            obj_xml.Load(str_file)
            For Each obj_map As Xml.XmlNode In obj_xml.FirstChild.ChildNodes
              If Not mobj_langs.ContainsKey(obj_map.Name) Then
                mobj_langs.Add(obj_map.Name, New Dictionary(Of String, String))
                For Each obj_langs As Xml.XmlNode In obj_map.ChildNodes
                  If Not mobj_langs(obj_map.Name).ContainsKey(obj_langs.Name) Then
                    mobj_langs(obj_map.Name).Add(obj_langs.Name, obj_langs.InnerText)
                  End If
                Next
              End If
            Next
          End If
        End If
      End If
      Return mobj_langs
    End Get
  End Property

  'Public ReadOnly Property Extensionless() As Boolean
  '  Get
  '    If Not Config.ContainsKey("site.extensionless") Then Config("site.extensionless") = False
  '    Return Config("site.extensionless")
  '  End GetB
  'End Property

  Public ReadOnly Property MultiLingual() As Boolean
    Get
      'If Not Config.ContainsKey("site.multilingual") Then Config("site.multilingual") = 0
      Return Config("site.multilingual") '(ToInt(Config("site.multilingual"), 0) = 1)
    End Get
  End Property

  Public ReadOnly Property Args() As ArgList
    Get
      On Error Resume Next
      Dim obj_args As ArgList = ToObj(Of ArgList)(HttpContext("args"))
      If obj_args Is Nothing Then
        obj_args = New ArgList(RawUrl)
        HttpContext("args") = obj_args
      End If
      Return obj_args
    End Get
  End Property

  Public Sub RewritePath(ByVal vstr_url As String, Optional ByVal rebaseClientPath As Boolean = False)
    Dim str_url As String = vstr_url.ToLower
    HttpContext.Current.RewritePath(vstr_url, rebaseClientPath)
    str_url = Request.Url.Scheme & "://" & Request.Url.Host & str_url
    HttpContext("cms.url") = New Uri(str_url)
  End Sub

End Module
