Namespace Plugins.Client

  Public Class Subtree
    Inherits Plugins.Client.Base

    Private mobj_parents As New List(Of Integer)

    Public Overrides ReadOnly Property Name() As String
      Get
        Return "Subtree Plugin"
      End Get
    End Property

    Public Overrides Function Render() As String
      On Error Resume Next
      Pages.Load()
      Dim obj_sb As New System.Text.StringBuilder()
      mobj_parents = GetParents(Page.ID, True)
      PagesLoad(obj_sb, GetRoot(Page.ID, True), 0)
      Return obj_sb.ToString
    End Function

    Private Sub PagesLoad(ByVal vobj_sb As System.Text.StringBuilder, ByVal vint_parent As Integer, ByVal vint_depth As Integer)
      Dim obj_pages As List(Of Page) = Global.Pages.Children(vint_parent)
      Dim b_ul As Boolean = False
      If obj_pages.Count > 0 Then
        For Each obj_pageinfo As Page In obj_pages
          If Security.HasPermissions(obj_pageinfo.Read) Then
            If Security.IsAdmin Then obj_pageinfo.Visible = 1
            If obj_pageinfo.Visible = 1 Then
              If Not b_ul Then
                vobj_sb.Append(vbNewLine & "<ul class=""depth_" & vint_depth & """>" & vbNewLine)
                b_ul = True
              End If
              vobj_sb.Append("<li")
              If Page.ID = obj_pageinfo.ID Then vobj_sb.Append(" class=""navactive""")
              vobj_sb.Append(">")
              vobj_sb.Append("<a href=""" & obj_pageinfo.Url & """><span>" & obj_pageinfo.Title & "</span></a>")
              If Page.ID = obj_pageinfo.ID Or mobj_parents.Contains(obj_pageinfo.ID) Then
                PagesLoad(vobj_sb, obj_pageinfo.ID, vint_depth + 1)
              End If
              vobj_sb.Append("</li>" & vbNewLine)
            End If
          End If
        Next
        If b_ul Then vobj_sb.Append("</ul>" & vbNewLine)
      End If
    End Sub
  End Class

End Namespace