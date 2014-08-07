Namespace Context.CMS

    Public Class Errors
        Inherits System.Web.UI.Page

        Private mstr_url As String

        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            On Error Resume Next
            Response.StatusCode = 404
            mstr_url = GetQuery("aspxerrorpath")
            If mstr_url.Length = 0 Then
                mstr_url = Server.UrlDecode(Request.QueryString.ToString.ToLower)
                Dim int_start As Integer = mstr_url.IndexOf(Request.Url.Host) + Request.Url.Host.Length
                int_start = mstr_url.IndexOf("/", int_start)
                mstr_url = mstr_url.Substring(int_start).TrimEnd("/")
            End If
            If Not Redirect() Then
                If Config("site.multilingual") Then
                    If Lang() Then Return
                Else
          Dim int_page As Integer = Global.Pages.ID("/error")
          If int_page <= 0 Then int_page = Global.Pages.ID("/errors")
          If RenderPage(int_page) Then Return
        End If
        Generic()
      End If
    End Sub

    Private Function Lang() As Boolean
      On Error Resume Next
      Dim str_lang As String = mstr_url.Substring(1, 2)
      Dim int_root As Integer = Global.Pages.ID("/" & str_lang & "/error")
      If int_root <= 0 Then int_root = Global.Pages.ID("/" & str_lang & "/errors")
      If int_root <= 0 Then int_root = Global.Pages.ID("/en/error")
      If int_root <= 0 Then int_root = Global.Pages.ID("/en/errors")
      Return RenderPage(int_root)
    End Function

        Private Function RenderPage(ByVal vint_id As Integer) As Boolean
            If vint_id > 0 Then
        Dim obj_page As Page = Global.Pages.Get(vint_id)
                Try
                    obj_page.Render(Response.OutputStream)
                    Return True
                Catch obj_ex As System.Exception
                    Response.Write("<!--" & vbNewLine)
                    Response.Write(Server.HtmlEncode(obj_ex.ToString) & vbNewLine)
                    Response.Write("-->" & vbNewLine)
                End Try
            End If
        End Function

        Private Function Redirect() As Boolean
            On Error Resume Next
            Dim str_link As String = String.Empty
            Dim i As Integer = 1
            Dim obj_match As Match
            Dim b_regex As Boolean = NullSafe(Config("cms.redirect.regex"), 0) = 1
            'Response.Write(mstr_url & "<br>")
            Do While Config.ContainsKey("cms.redirect.match." & i)
                If b_regex Then
                    obj_match = Regex.Match(mstr_url, Config("cms.redirect.match." & i), RegexOptions.IgnoreCase)
                    'Response.Write(Config("cms.redirect.match." & i) & "=" & obj_match.Success & "<br>")
                    If obj_match.Success Then
                        str_link = Config("cms.redirect.replace." & i) 'FULL URL
                        If Not String.IsNullOrEmpty(str_link) Then
                            For j As Integer = 0 To obj_match.Captures.Count - 1
                                str_link = str_link.Replace("$" & j + 1, obj_match.Captures(j).Value)
                            Next
                            Response.RedirectLocation = str_link
                            Response.StatusCode = 301
                            Response.End()
                            Return True
                        End If
                    End If
                Else
                    If mstr_url = Config("cms.redirect.match." & i).trimend("/") Then
                        str_link = Config("cms.redirect.replace." & i) 'FULL URL
                        If Not String.IsNullOrEmpty(str_link) Then
                            Response.RedirectLocation = str_link
                            Response.StatusCode = 301
                            Response.End()
                            Return True
                        End If
                    End If
                End If
                i += 1
                If i > 250 Then Exit Do
            Loop
            'Return True
        End Function

        Private Sub Generic()
            On Error Resume Next
            Dim str_path As String = GetQuery("aspxerrorpath", mstr_url)
            Dim obj_ex As System.Exception = Application("LastError") ' Server.GetLastError().GetBaseException()
            'If obj_ex Is Nothing Then obj_ex = New System.Exception("An Error Occurred")
            Dim str_file As String = "/cms/templates/errors.htm"
            Dim obj_tp As New HTMLTemplate(str_file)
            If obj_tp IsNot Nothing Then
                obj_tp.SetItem("title", "Page Not Available")
                obj_tp.SetItem("description", "The page you requested cannot be found. The page you are looking for might have been removed, had its name changed, or is temporarily unavailable.")
                obj_tp.SetItem("ver.cms", "" & Config("cms.version"))
                obj_tp.SetItem("ver.net", System.Environment.Version.ToString)
                obj_tp.SetItem("path", str_path)
                If obj_ex IsNot Nothing Then
                    obj_tp.SetItem("raw", Server.HtmlEncode(obj_ex.ToString).Replace("&gt;", ""))
                End If
                obj_tp.Render(Response.OutputStream)
            Else
                Response.Write("<html><head><title>Page Not Available</title></head><body><H1>Page Not Available</H1>")
                Response.Write("<!--" & Server.HtmlEncode(obj_ex.ToString).Replace("&gt;", "") & "-->")
                Response.Write("</body></html>")
            End If
        End Sub

    End Class

End Namespace
