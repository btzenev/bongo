Namespace Context.CMS

  Public Class [Default]
    Inherits System.Web.UI.Page

    Private mobj_priv As New List(Of String)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      SetCache(HttpCacheability.NoCache)

      If Request.IsAuthenticated() Then
        If Security.IsAdmin Then
          If Args.Count > 1 Then
            RenderPlugin(Args(1))
          Else
            RenderCMS()
          End If
        Else
          mobj_priv = Security.GetAdminPriv()
          If (mobj_priv.Count > 0) Then
            Dim str_plugin As String = Args(1)
            If Args.Count > 1 Then
              If mobj_priv.Contains("admin_" & str_plugin) Then
                RenderPlugin(str_plugin)
              Else
                Response.Redirect("/cms/", True)
              End If
            Else
              RenderCMS()
            End If
          Else
            Response.Redirect("/", True)
          End If
        End If
      Else
                Security.RedirectToLoginPage(url:="/cms/login", returnurl:="/cms/")
      End If

    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload
      Data.Current.Close()
    End Sub

    Private Sub RenderPlugin(ByVal vstr_plugin As String)
      If vstr_plugin.StartsWith("admin_") Then
        Response.Redirect(RawUrl.Replace(vstr_plugin, vstr_plugin.Substring(6)), True)
        Return
      End If

      Try
        Dim obj_plugin As Admin.Interface = Admin.Get("admin_" & vstr_plugin)
        If obj_plugin IsNot Nothing Then
          If Args.Count = 2 Then
            If obj_plugin.Full Then
              Response.Write("" & obj_plugin.Render)
            Else
              Dim obj_tp As New HTMLTemplate("/cms/templates/admin.htm")
              obj_plugin.Menu(obj_tp)
              obj_tp.SetItem("iframe", "/cms/" & vstr_plugin & "/default.aspx")
              obj_tp.SetItem("link", vstr_plugin & ".aspx")
              obj_tp.SetItem("title", obj_plugin.Name)
              obj_tp.SetItem("name", obj_plugin.Name)
              obj_tp.SetItem("cms.version", Config("cms.version"))
                            obj_tp.SetItem("cms.title", NullSafe(Config("cms.name"), "AMS"))
              obj_tp.Render(Response.OutputStream)
            End If
          Else
            Response.Write("" & obj_plugin.Render)
          End If
        Else
          Response.Write("Not available at this time")
          'Response.Redirect("/cms/", True)
        End If
      Catch ex As Exception
        Response.Write(FormatError(ex))
      End Try
    End Sub

    Private Sub RenderCMS()
      On Error Resume Next
      Dim obj_tp As New HTMLTemplate("/cms/templates/manager.htm")
      Dim obj_plugin As Admin.Interface

            Dim obj_plugins As New List(Of String)(Plugins.Admin.List)
      obj_plugins.Sort()

      If obj_plugins.Contains("admin_site") Then
        If Security.IsAdmin Then
          obj_tp("admin").AddItem("value", "site")
          obj_tp("admin").SetItem("name", "<b>Site Management</b>")
        End If
        obj_plugins.Remove("admin_site")
      End If

      If obj_plugins.Contains("admin_asset") Then
        obj_plugins.Remove("admin_asset")
        If Security.IsAdmin Then obj_plugins.Add("admin_asset")
      End If

      'If obj_plugins.Contains("admin_data") Then
      '  obj_plugins.Remove("admin_data")
      '  If Security.IsGlobal Then obj_plugins.Add("admin_data")
      'End If

      'If obj_plugins.Contains("admin_tools") Then
      '  obj_plugins.Remove("admin_tools")
      '  If Security.IsGlobal Then obj_plugins.Add("admin_tools")
      'End If

      Dim str_key, str_name As String
      Dim b_show As Boolean = False

      For Each str_plugin As String In obj_plugins
        b_show = (Security.IsGlobal Or Security.IsAdmin Or mobj_priv.Contains(str_plugin))
        If b_show And Config.ContainsKey(str_plugin) Then
          b_show = Security.IsGlobal Or (NullSafe(Config(str_plugin), 0) = 1)
        End If
        If b_show Then
          str_key = str_plugin.Replace("admin_", "")
          obj_plugin = Plugins.Admin.Get(str_plugin)
          If obj_plugin IsNot Nothing Then
            str_name = obj_plugin.Name
            If str_key = "asset" Or str_key = "data" Or str_key = "tools" Then str_name = "<i>" & str_name & "</i>"
            obj_tp("admin").AddItem("value", str_key)
            obj_tp("admin").SetItem("name", str_name)
          End If
          obj_plugin = Nothing
        End If
      Next

      Dim str_user As String = User.Identity.Name.ToLower
            If str_user = "assent" Then
                str_user = "Global Administrator"
            End If
      obj_tp.SetItem("user", str_user)
      obj_tp.SetItem("cms.version", Config("cms.version"))
            obj_tp.SetItem("cms.title", NullSafe(Config("cms.name"), "BongoAMS"))
      obj_tp.Render(Response.OutputStream)
    End Sub

  End Class

End Namespace
