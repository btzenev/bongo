Namespace Context

  Public Class [Default]
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      ' On Error Resume Next
      If Global.Application.Status = 0 Then Global.Application.Initialize()
      Page.Culture = NullSafe(Config("culture." & GetLang()), "en-US")
      SetCache(HttpCacheability.Public)
      'Response.Cache.SetCacheability(HttpCacheability.NoCache)
      If Data.Current.Open Then 'connect to database
        Dim obj_page As Page
        If Args.Count = 0 Then
          obj_page = Pages.Get(ToInt(Config("page.landing")))
        Else
          obj_page = Pages.Get(Url.LocalPath)
        End If

        If obj_page Is Nothing And Args.Count > 0 Then

          Dim str_url As String = "/"
          For Each str_arg As String In Args
            str_url &= str_arg & "/"
            Dim obj_tmp As Page = Pages.Get(str_url)
            If obj_tmp Is Nothing Then Exit For
            obj_page = obj_tmp
          Next

          If obj_page IsNot Nothing Then
            If obj_page.Application = 0 Then
              obj_page = Nothing
            End If
          End If
        End If

        If obj_page IsNot Nothing Then
          If Security.HasPermissions(obj_page.Read) Then
            obj_page.Render(Response.OutputStream)
          Else
            Security.RedirectToLoginPage(returnurl:=obj_page.Url)
          End If
        Else
          HttpContext.Current.Response.StatusCode = 404
          HttpContext.Current.Response.StatusDescription = "404 - Not Found"
          HttpContext.Current.ApplicationInstance.CompleteRequest()
          Throw New HttpException(404, "Not Found")
        End If

      Else
        On Error GoTo 0
        Throw New System.Exception("PAGE: Unable to Load Database.")
      End If
    End Sub

    Private Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Unload, MyBase.Error
      Data.Current.Close()
      If Global.Application.Status = 1 Then Global.Application.Status = 2
    End Sub

    'Private Sub RenderPage()

    'End Sub

    'Private Sub RenderPlugin1()
    '    If GetQuery("action") <> "" Then
    '        Dim obj_plugin As Client.Interface = Plugins.Client.Get(Request.QueryString("action"))
    '        If Not obj_plugin Is Nothing Then
    '            obj_plugin.ModuleID = -1
    '            obj_plugin.PageID = Pages.ID(Request.Url.LocalPath.ToString.ToLower)
    '            'obj_plugin, "URL", Request.Url.LocalPath.ToString.ToLower)
    '            Response.Write("" & obj_plugin.Render)
    '            obj_plugin = Nothing
    '        End If
    '    End If
    'End Sub

  End Class

End Namespace