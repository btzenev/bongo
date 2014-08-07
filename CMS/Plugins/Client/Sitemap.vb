Namespace Plugins.Client

  Public Class Sitemap
    Inherits Plugins.Client.Base

    Private mint_offset As Integer

    Public Overrides ReadOnly Property Name() As String
      Get
        Return "Sitemap Plugin"
      End Get
    End Property

    Public Overrides Function Render() As String
      mint_offset = NullSafe(Config("sitemap.offset"), 0)
      Dim obj_sb As New System.Text.StringBuilder()
      obj_sb.Append("<div class=""sitemap"">")
            PagesLoad(obj_sb, GetRoot(Page.ID), 0)
      obj_sb.Append("</div>" & vbNewLine)
      Return obj_sb.ToString
    End Function

    Private Sub PagesLoad(ByVal vobj_sb As System.Text.StringBuilder, ByVal vint_parent As Integer, ByVal vint_depth As Integer)
      Dim obj_children As List(Of Page) = Global.Pages.Children(vint_parent)
      If obj_children.Count > 0 Then
        vobj_sb.Append(vbNewLine & "<ul class=""sitemap_depth_" & vint_depth & """>" & vbNewLine)
        For Each obj_page As Page In obj_children
          If Security.HasPermissions(obj_page.Read) Then
            Dim b_write As Boolean = Security.IsAdmin
            If obj_page.Visible = 1 Or b_write Then
              vobj_sb.Append("<li><a href=""" & obj_page.Url & """>" & obj_page.Title & "</a>")
              If b_write And obj_page.Visible = 0 Then vobj_sb.Append("&nbsp;(hidden)")
              If obj_page.Url <> "/fr" Then
                PagesLoad(vobj_sb, obj_page.ID, vint_depth + 1)
              End If
              vobj_sb.Append("</li>" & vbNewLine)
            End If
          End If
        Next
        vobj_sb.Append("</ul>" & vbNewLine)
      End If
    End Sub

  End Class

End Namespace