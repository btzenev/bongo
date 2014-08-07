Imports System.Web
Imports System.Web.Security
Imports System.Security.Principal

Public NotInheritable Class Security

  Public Shared ReadOnly Property UserID() As Integer
    Get
      On Error Resume Next
      If HttpContext.Current.Request.IsAuthenticated Then
        Return ToInt("" & HttpContext.Current.User.Identity.Name, 0)
      End If
      Return 0
    End Get
  End Property

  Public Shared Sub RedirectToLoginPage(Optional ByVal url As String = "", Optional ByVal returnurl As String = "", Optional ByVal query As String = "")
    Security.Logout(False)
    Dim obj_sb As New Text.StringBuilder
    If IsEmpty(url) Then url = Config("page.login")
    If MultiLingual Then url = url.Replace("{{lang}}", GetLang())
    obj_sb.Append(url & "?")
    obj_sb.Append(query)
    If NotEmpty(query) Then obj_sb.Append("&")
    If IsEmpty(returnurl) Then returnurl = GetQuery("redirect", GetQuery("returnurl", "/"))
    obj_sb.Append("returnurl=" & Server.UrlEncode(returnurl))
    Response.Redirect(obj_sb.ToString, True)
  End Sub

  Public Shared Sub RedirectToMemberPage(Optional ByVal url As String = "", Optional ByVal returnurl As String = "", Optional ByVal query As String = "")
    On Error Resume Next
    Dim str_url As String = GetQuery("redirect", GetQuery("returnurl", ToStr(Config("page.members"), "/")))
    If MultiLingual Then str_url = str_url.Replace("{{lang}}", GetLang())
    'If Security.IsAdmin And str_url.StartsWith("/cms/") Then
    'Response.Redirect(str_url, True)
    'ElseIf Pages.Get(str_url).ID > 0 Then
    Response.Redirect(str_url, True)
    'Else
    'Response.Redirect("/", True)
    'End If
  End Sub

  Public Shared Function Login(Optional ByVal vb_redirect As Boolean = True) As String
    Return Login(GetForm("username"), GetForm("password"), vb_redirect)
  End Function

  Public Shared Function Login(ByVal vstr_name As String, ByVal vstr_password As String, Optional ByVal vb_redirect As Boolean = True) As String
    Security.Logout(False)
    Dim obj_row As DataRow
    Dim str_roles As String
    Dim int_id As Integer
    vstr_name = ToStr(vstr_name).ToLower.Replace("'", "''")
    If Not vstr_name = "" Then
      vstr_password = ToStr(vstr_password)
      If Not vstr_password = "" Then
        Dim str_global As String = String.Empty
        Try
          Dim str_gpass As String = Application.getPassword()
          If Not String.IsNullOrEmpty(str_gpass) Then
            str_global = str_gpass
          Else
            Dim astr_global() As String = Request.Url.Host.ToLower.Split(".")
            str_global = IIf(astr_global(0) = "www", astr_global(1), astr_global(0))
          End If
        Catch ex As Exception
          str_global = Request.Url.Host.ToLower
        End Try
        If vstr_name.ToLower = GlobalUser And vstr_password.ToLower = str_global.ToLower Then
          str_roles = GlobalRole & ";"
        Else
          obj_row = Data.Current.Row("SELECT * from [cms_users] WHERE [userid] LIKE '" & vstr_name & "'")
          If Not obj_row Is Nothing Then
            If ToStr(obj_row("password")).ToUpper = Encrypt(vstr_password).ToUpper Then
              Return Login(obj_row, vb_redirect)
            Else
              Return "Invalid password"
            End If
          Else
            Return "Invalid username"
          End If
        End If
      Else
        Return "Please specify a password"
      End If
    Else
      Return "Please specify a username"
    End If
    Try
      Dim obj_date As Date = DateTime.Now.AddMinutes(30)
      Dim obj_ticket As New FormsAuthenticationTicket(1, int_id, DateTime.Now, obj_date, False, str_roles, FormsAuthentication.FormsCookiePath)
      Dim obj_cookie As New HttpCookie(FormsAuthentication.FormsCookieName, FormsAuthentication.Encrypt(obj_ticket))
      Response.Cookies.Add(obj_cookie)
      Security.SetPrincipal(int_id, str_roles) ' vstr_name, str_roles)
      Response.Cookies("dash").Expires = Date.Now.AddYears(-1)
      If vb_redirect Then RedirectToMemberPage()
    Catch ex As Exception
      Return "An error occurred while authenticating your user. Please contact the site administrator.<!--" & Server.HtmlEncode(ex.ToString) & "-->"
    End Try
    Return String.Empty
  End Function

  Public Shared Function Login(ByVal vint_id As Integer, Optional ByVal vb_redirect As Boolean = True) As String
    On Error Resume Next
    Security.Logout(False)
    Dim obj_row As DataRow = Data.Current.Row("SELECT * from [cms_users] WHERE [id]=" & vint_id)
    Return Login(obj_row, vb_redirect)
  End Function

  Private Shared Function Login(ByVal vobj_row As DataRow, Optional ByVal vb_redirect As Boolean = True) As String
    Try
      If Not vobj_row Is Nothing Then
        Dim int_id As Integer = ToInt(vobj_row("id"), 0)
        Dim b_enabled As Boolean
        Dim obj_table As DataTable = vobj_row.Table
        If obj_table.Columns.Contains("enabled") Then
          b_enabled = ToBool(vobj_row("enabled"), False)
        Else
          b_enabled = True
        End If

        If b_enabled Then
          Dim str_roles As String = String.Empty
          If obj_table.Columns.Contains("roles") Then
            str_roles = "" & vobj_row("roles")
          Else
            str_roles = "" & vobj_row("roleid")
          End If
          str_roles = str_roles.Replace(",", ";").Replace(":", ";").Trim(";")

          Dim obj_date As Date = DateTime.Now.AddMinutes(30)
          Dim obj_ticket As New FormsAuthenticationTicket(1, int_id, DateTime.Now, obj_date, False, str_roles, FormsAuthentication.FormsCookiePath)
          Dim obj_cookie As New HttpCookie(FormsAuthentication.FormsCookieName, FormsAuthentication.Encrypt(obj_ticket))
          Response.Cookies.Add(obj_cookie)
          Security.SetPrincipal(int_id, str_roles)
          Response.Cookies("dash").Expires = Date.Now.AddYears(-1)
          If vb_redirect Then RedirectToMemberPage()
        Else
          Return "User disabled"
        End If
      Else
        Return "Invalid username"
      End If
    Catch ex As Exception
      Return "An error occurred while authenticating your user. Please contact the site administrator.<!--" & Server.HtmlEncode(ex.ToString) & "-->"
    End Try
  End Function

  Public Shared Sub Logout(Optional ByVal vb_redirect As Boolean = True)
    On Error Resume Next
    FormsAuthentication.SignOut()
    Response.Cookies.Item(FormsAuthentication.FormsCookieName).Expires = Date.Now.AddYears(-1)
    If vb_redirect Then Response.Redirect("/", True)
  End Sub

  Public Shared Function GetRoles() As ArrayList
    Dim obj_roles As New ArrayList
    If Request.IsAuthenticated Then
      Dim obj_cookie As HttpCookie = Request.Cookies(FormsAuthentication.FormsCookieName)
      If Not obj_cookie Is Nothing Then
        Dim obj_ticket As FormsAuthenticationTicket = FormsAuthentication.Decrypt(obj_cookie.Value)
        obj_roles.AddRange(obj_ticket.UserData.Trim(";").Split(";"))
      End If
    End If
    Return obj_roles
  End Function

#Region " GetAdminPriv "

  Public Shared Function GetAdminPriv() As List(Of String)
    On Error Resume Next
    Dim obj_return As New List(Of String)
    For Each int_role As Integer In GetRoles()
      obj_return.AddRange(GetAdminPriv(int_role).ToArray)
    Next
    Return obj_return
  End Function

  Public Shared Function GetAdminPriv(ByVal vint_role As Integer) As List(Of String)
    On Error Resume Next
    Dim obj_return As New List(Of String)
    Dim obj_row As DataRow = Data.Current.Row("select * from [cms_roles] where [id]=" & vint_role)
    If Not obj_row Is Nothing Then
      If obj_row.Table.Columns.Contains("access") Then
        Dim str_access As String = "" & obj_row("access")
        str_access = str_access.Replace(",", ";").Replace(":", ";").Trim(";")
        For Each str_val As String In str_access.Split(";")
          If NotEmpty(str_val) And Not obj_return.Contains(str_val) Then obj_return.Add(str_val)
        Next
      End If
    End If
    Return obj_return
  End Function


#End Region

#Region " IsInRoles "

  Public Shared Function IsInRoles(ByVal vstr_id As String) As Boolean
    On Error Resume Next
    For Each role As String In GetRoles()
      If role = vstr_id Then
        Return True
      End If
    Next role
    Return False
  End Function

  Public Shared Function IsInRoles(ByVal roles As String, ByVal vstr_id As String) As Boolean
    On Error Resume Next
    roles = roles.Replace(",", ";").Replace(":", ";").Trim(";").Trim(";") & ";"
    For Each role As String In roles.Split(New Char() {";"c})
      If role = vstr_id Then
        Return True
      End If
    Next role
    Return False
  End Function

#End Region

  Public Shared Function IsAdmin() As Boolean
    If Request.IsAuthenticated Then
      Return User.IsInRole(GlobalRole) Or User.IsInRole(AdminRole)
    End If
  End Function

  Public Shared Function IsGlobal() As Boolean
    If Request.IsAuthenticated Then
      Return User.IsInRole(GlobalRole)
    End If
  End Function

#Region " HasPermissions "

  Public Shared Function HasPermissions(ByVal vstr_roles As String) As Boolean
    On Error Resume Next
    Dim obj_roles As New List(Of String)
    vstr_roles = ToStr(vstr_roles).Replace(";", ":").TrimStart(":").TrimEnd(":")
    If NotEmpty(vstr_roles) Then obj_roles.AddRange(vstr_roles.Split(":"))
    Return HasPermissions(obj_roles)
  End Function

  Public Shared Function HasPermissions(ByVal vobj_roles As List(Of String)) As Boolean
    If vobj_roles Is Nothing Then Return False
    If vobj_roles.Contains(EveryRole.ToString) Then Return True
    If Request.IsAuthenticated Then
      If User.IsInRole(GlobalRole) Or User.IsInRole(AdminRole) Then Return True
      If Not vobj_roles Is Nothing Then
        For Each str_role As String In vobj_roles
          str_role = ToStr(str_role)
          If Not str_role = "" Then
            If User.IsInRole(str_role) Then
              Return True
            End If
          End If
        Next
      End If
    Else
      If Not vobj_roles Is Nothing Then
        Return vobj_roles.Contains(GuestRole.ToString)
      End If
    End If
  End Function

#End Region

#Region " Allowed "

  Public Shared Function Allowed(ByVal vstr_roles As String) As Boolean
    Return HasPermissions(vstr_roles)
  End Function

  Public Shared Function NotAllowed(ByVal vstr_roles As String) As Boolean
    Return Not HasPermissions(vstr_roles)
  End Function

  Public Shared Function Allowed(ByVal vobj_roles As List(Of String)) As Boolean
    Return HasPermissions(vobj_roles)
  End Function

  Public Shared Function NotAllowed(ByVal vobj_roles As List(Of String)) As Boolean
    Return Not HasPermissions(vobj_roles)
  End Function

#End Region

  Private Shared Function GetPermissions(ByVal vint_id As Integer) As ArrayList
    Dim obj_roles As New ArrayList
    Dim obj_query As New QueryBuilder("cms_users")
    obj_query.Add("roles")
    obj_query.Where("id", vint_id)
    Dim str_roles As String = ToStr(Data.Current.Get(obj_query))
    If Not str_roles = "" Then
      obj_roles.AddRange(str_roles.Trim(";").Split(";"))
    End If
    Return obj_roles
  End Function

  Public Shared Sub SetPrincipal(ByVal vstr_name As String, ByVal vstr_roles As String)
    HttpContext.Current.User = GetPrincipal(vstr_name, vstr_roles)
  End Sub

  Public Shared Function GetPrincipal(ByVal vstr_name As String, ByVal vstr_roles As String) As GenericPrincipal
    Dim obj_roles As New ArrayList
    vstr_roles = ToStr(vstr_roles)
    Dim str_roles() As String = vstr_roles.Trim(";").Split(";")
    Dim obj_id As New GenericIdentity(vstr_name)
    Return New GenericPrincipal(obj_id, str_roles)
  End Function

End Class
