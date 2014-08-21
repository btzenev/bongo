Imports System.IO
Imports System.Text
Imports System.Configuration
Imports System.Security
Imports System.Security.Principal
Imports System.Web
Imports System.Web.SessionState
Imports System.Web.Security
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Web.Routing

Public Class Application
  Inherits System.Web.HttpApplication

    Private Shared mstr_password As String = "plovdiv"
  Private Shared mobj_domains As New List(Of String)

  Public Shared Status As Integer = 0

  Private Class RouteHandler
    Implements IRouteHandler

    Public Function GetHttpHandler(ByVal requestContext As System.Web.Routing.RequestContext) As System.Web.IHttpHandler Implements System.Web.Routing.IRouteHandler.GetHttpHandler
      Dim str_url As String = Url.LocalPath.ToLower
      If str_url.StartsWith("/cms") Then
        Select Case Args(1)
          Case "image" : Return New Context.CMS.Image
          Case "post" : Return New Context.CMS.Post
          Case "select" : Return New Context.CMS.Select
          Case "login" : Return New Context.CMS.Login
          Case "logout", "logoff"
            Security.Logout()
            requestContext.HttpContext.Response.Redirect("/", True)
          Case "reload"
            Global.Pages.Load()
            Plugins.Load()
            requestContext.HttpContext.Response.Redirect("/", True)
          Case Else 'admin only functionality
            If requestContext.HttpContext.Request.IsAuthenticated() Then
              Return New Context.CMS.Default
            Else
              Security.RedirectToLoginPage(url:="/cms/login", returnurl:="/cms/")
            End If
        End Select
      ElseIf str_url.StartsWith("/logout") Then
        Security.Logout()
        requestContext.HttpContext.Response.Redirect("/", True)
      End If
      Return New Context.Default
    End Function

  End Class

#Region " Verify "

  Private Shared Sub Verify(ByVal vstr_name As String)
    If Not Config.ContainsKey(vstr_name) Then Throw New System.Exception("APPLICATION: [" & vstr_name & "] is invalid. please check your web.config file.")
  End Sub

  Private Shared Sub Verify(ByVal vstr_name As String, ByVal vstr_default As String)
    On Error Resume Next
    If Config.ContainsKey(vstr_name) Then
      Config(vstr_name) = ToStr(Config(vstr_name), vstr_default)
    Else
      Config.Add(vstr_name, vstr_default)
    End If
  End Sub

  Private Shared Sub Verify(ByVal vstr_name As String, ByVal vint_default As Integer)
    On Error Resume Next
    If Config.ContainsKey(vstr_name) Then
      Config(vstr_name) = ToInt(Config(vstr_name), vint_default)
    Else
      Config.Add(vstr_name, vint_default)
    End If
  End Sub

  Private Shared Sub Verify(ByVal vstr_name As String, ByVal vb_default As Boolean)
    On Error Resume Next
    If Config.ContainsKey(vstr_name) Then
      Config(vstr_name) = ToBool(Config(vstr_name), vb_default)
    Else
      Config.Add(vstr_name, vb_default)
    End If
  End Sub

#End Region

  Protected Friend Shared Function getPassword()
    Return mstr_password
  End Function

    Public Shared Sub Initialize()

        ' On Error Resume Next
        If Status = 1 Then
            If HttpContext.Current IsNot Nothing Then
                Threading.Thread.Sleep(100)
                HttpContext.Current.Response.Redirect(RawUrl, True)
                HttpContext.Current.Response.End()
            End If
        ElseIf Status = 0 Then
            Status = 1

            'reload web.config data into settings global context
            Config.Clear()
            For Each str_key As String In ConfigurationManager.AppSettings.Keys
                Config.Add(str_key, ConfigurationManager.AppSettings(str_key))
            Next

            'Verify("cms.id")

            ''verify license information
            'Dim str_lic As String = MapPath("/license.dat")
            'If Not IO.File.Exists(str_lic) Then Throw New System.Exception("APPLICATION: License not found.")

            'str_lic = ReadFile(str_lic).Trim
            'If String.IsNullOrEmpty(str_lic) Then Throw New System.Exception("APPLICATION: License not configured.")

            'Dim obj_key As New Encryption.Data(AppID)
            'Dim obj_data As New Encryption.Data()
            'obj_data.Hex = str_lic
            'Dim obj_syn As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
            'obj_data = obj_syn.Decrypt(obj_data, obj_key)

            'mobj_domains = New List(Of String)
            'For Each str_key As String In obj_data.ToString.ToLower.Split(vbNewLine)
            '  If Not String.IsNullOrEmpty(str_key) Then
            '    str_key = str_key.Trim
            '    If str_key.StartsWith("pass:") Then
            '      mstr_password = str_key.Replace("pass:", "")
            '    Else
            '      mobj_domains.Add(str_key)
            '    End If
            '  End If
            'Next
            'If mobj_domains.Count = 0 Then Throw New System.Exception("APPLICATION: License not configured.")

            'verifying settings and placing default values where needed
            Verify("cms.version", My.Application.Info.Version.ToString & " " & My.Application.Info.Description)
            Verify("cms.name", "CMS")
            Verify("cms.site", 1)
            'Verify("site.language", 2)
            Verify("site.multilingual", 0)
            'Config("site.multilingual") = (ToInt(Config("site.multilingual"), 0) = 1)

            'Verify("email.error", False)
            'Verify("email.port", 25)
            'Verify("email.host", "")
            'Verify("email.username", "")
            'Verify("email.password", "")

            Verify("page.login", "/cms/login")
            Verify("page.members", "/")
            Dim isCMS As Integer = 1
            isCMS = Config("cms.site")
            If isCMS = 0 Then

                RouteTable.Routes.RouteExistingFiles = False
                RouteTable.Routes.Add("default", New Route("{*value}", New RouteHandler))
                Plugins.Load()
                If Data.Current.Open Then
                    Pages.Load()
                    Data.Current.Close()
                Else
                    Exception("App", "Unable to Load Databases", True)
                End If
            End If
        End If

    End Sub

  Protected Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
    Global.Application.Initialize()
  End Sub

  Protected Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
    Dim str_host As String = Url.Host.ToLower
        Dim str_url As String = Url.LocalPath.ToLower
        Dim str_query As String = Url.Query.TrimStart("?")
        Throw New System.Exception(Join(mobj_domains.ToArray, ","))
        'check license for domain
        'If Not mobj_domains.Contains(str_host) Then Throw New System.Exception("APPLICATION: Not Licensed.")
        If Global.Application.Status = 1 Then Response.Redirect("/default.aspx", True)


        If Not File.Exists(MapPath(str_url)) And Not Pages.List Is Nothing Then

            Dim str_ext As String = Path.GetExtension(str_url).TrimStart("."c).ToLower
            If str_ext.Length = 0 And Not str_url.EndsWith("/") Then
                str_url &= "/" 'default.aspx"
                If str_query.Length > 0 Then str_url &= "?" & str_query
                Response.Redirect(str_url, True)
            Else

                If str_url = "/" Or str_url.StartsWith("/default.aspx") Then
                    str_url = "/default.aspx?pageid=" & Config("page.landing")
                    'ElseIf str_url.StartsWith("/cms/image/") Then
                    '  str_url = str_url.Substring(11)
                    '  Dim int_start As Integer = str_url.IndexOf("/")
                    '  Dim str_dim As String = str_url.Substring(0, int_start)
                    '  str_url = "/cms/image.aspx?dim=" & str_dim & "&file=" & str_url.Substring(int_start)
                ElseIf str_url.StartsWith("/cms/") Then
                    str_url = "/cms/default.aspx"
                Else
                    If str_url.Contains("/default.aspx") And str_url.Length > 13 Then
                        str_url = str_url.Replace("/default.aspx", "")
                    ElseIf str_url.EndsWith("/") And str_url.Length > 1 Then
                        str_url = str_url.Substring(0, str_url.Length - 1)
                    End If

                    Dim obj_page As Page = Pages.Get(str_url)
                    If obj_page IsNot Nothing Then
                        If Not Security.IsAdmin() And NotEmpty(obj_page.Redirect) Then
                            Response.Status = "301 Moved Permanently"
                            Response.AddHeader("Location", obj_page.Redirect)
                            Response.End()
                            Return
                        End If
                        str_url = "/default.aspx?pageid=" & obj_page.ID
                    Else
                        Throw New HttpException(404, "Not Found")
                    End If
                End If
                If NotEmpty(str_query) Then str_url &= IIf(str_url.IndexOf("?") > 0, "&", "?") & str_query
                HttpContext.Current.RewritePath(str_url)
            End If
        End If
  End Sub

  Protected Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
    If Request.IsAuthenticated Then
      Dim obj_cookie As HttpCookie = Request.Cookies(FormsAuthentication.FormsCookieName)
      If Not obj_cookie Is Nothing Then
        Dim obj_ticket As FormsAuthenticationTicket = FormsAuthentication.Decrypt(obj_cookie.Value)
        If FormsAuthentication.SlidingExpiration Then
          obj_ticket = FormsAuthentication.RenewTicketIfOld(obj_ticket)
          'Response.Cookies.Remove(FormsAuthentication.FormsCookieName)
          obj_cookie = New HttpCookie(FormsAuthentication.FormsCookieName, FormsAuthentication.Encrypt(obj_ticket))
          Response.Cookies.Add(obj_cookie)
        End If
        Security.SetPrincipal(obj_ticket.Name, obj_ticket.UserData)
      End If
    End If
  End Sub

  Protected Sub Application_Error(ByVal sender As Object, ByVal e As System.EventArgs)
    Dim obj_app As HttpApplication = DirectCast(sender, HttpApplication)
    Dim obj_ex As System.Exception = Server.GetLastError.GetBaseException()
    Application("LastError") = Server.GetLastError
    Data.Current.Close()
  End Sub

  Protected Sub Application_EndRequest(ByVal sender As Object, ByVal e As EventArgs)
    On Error Resume Next
    Data.Current.Close()
  End Sub

End Class
