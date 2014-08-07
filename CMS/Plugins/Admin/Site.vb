Namespace Plugins.Admin

  Public Class Site
    Inherits Base

    Public Overrides ReadOnly Property Name() As String
      Get
        Return "Site Management"
      End Get
    End Property

    Public Overrides ReadOnly Property Full() As Boolean
      Get
        Return True
      End Get
    End Property

    Public Overrides Function Render() As String
      If Security.IsAdmin Then

        If Args.Count = 2 Or (Args.Count = 3 And Args(2) = "default") Then
          Dim obj_plugin As Plugins.Client.Interface
          Dim obj_tp As New HTMLTemplate("/cms/plugins/admin_site/template.htm")
          Dim obj_plugins As New ArrayList(Plugins.Client.List)
          obj_plugins.Sort()
          For Each str_plugin As String In obj_plugins
            Try
              obj_plugin = Plugins.Client.Get(str_plugin)
              If Not obj_plugin Is Nothing Then
                If Not obj_plugin.GetType Is GetType(Client.Error) Then
                  If obj_plugin.Visible Or Security.IsInRoles(GlobalRole) Then
                    obj_tp("plugins").AddItem("plugin", str_plugin)
                    obj_tp("plugins").SetItem("name", obj_plugin.Name)
                  End If
                Else
                  obj_tp("plugins").AddItem("plugin", str_plugin)
                  obj_tp("plugins").SetItem("name", str_plugin)
                End If

              End If
            Catch ex As Exception
              'mobj_tp.SetItem("name", obj_plugins(i))
            End Try
            obj_plugin = Nothing
          Next
          Return obj_tp.Render
        Else
          If Args.Count = 4 Then
            Dim int_parent As Integer = GetQuery("parent", 0)
            Dim int_page As Integer = GetQuery("page", 0)
            Dim int_id As Integer = GetQuery("id", 0)
            Select Case Args(2)
              Case "pages"
                Select Case Args(3)
                  'Case "templates" : Return ListTemplates()
                  Case "list" : Return ListPages(0, 0)
                  Case "add", "edit" : Return EditPage(int_id, int_parent)
                  Case "del" : Return Pages.Delete(int_id)
                  Case "reorder"
                    Dim int_ord As Integer
                    Dim obj_ord As New List(Of Integer)
                    For Each str_id As String In GetQuery("ord").Split("|")
                      int_ord = ToInt(str_id)
                      If int_ord > 0 And Not obj_ord.Contains(int_ord) Then obj_ord.Add(int_ord)
                    Next
                    Return Pages.ReOrder(int_parent, obj_ord)
                End Select
              Case "mods"
                Dim obj_mod As Modules
                If int_id > 0 Then
                  obj_mod = New Modules(int_id)
                  Select Case Args(3)
                    Case "edit" : Return EditMod(obj_mod)
                    Case "del" : Return obj_mod.Delete
                  End Select
                Else
                  Select Case Args(3)
                    Case "add"
                      obj_mod = New Modules(GetQuery("plugin"), int_page)
                      If obj_mod.Save Then
                        Return obj_mod.Render
                      Else
                        Return False
                      End If
                    Case "reorder"
                      Dim int_ord As Integer
                      Dim obj_ord As New List(Of Integer)
                      For Each str_id As String In GetQuery("ord").Split("|")
                        int_ord = ToInt(str_id)
                        If int_ord > 0 And Not obj_ord.Contains(int_ord) Then obj_ord.Add(int_ord)
                      Next

                      Dim str_pos As String = GetQuery("position")
                      Return Modules.ReOrder(str_pos, obj_ord)
                  End Select
                End If
            End Select

          End If
        End If

      End If

      Return HTMLTemplate.QuickRender("/cms/templates/permissions.htm")
    End Function

    Public Function EditMod(ByVal vobj_mod As Modules) As String
      Dim obj_tp As New HTMLTemplate("/cms/plugins/admin_site/dialog.htm")
      Try
        If Request.Form.Count > 0 Then

          With vobj_mod
            .PageID = GetRequest("pageid", .PageID)
            .Position = GetRequest("position", .Position)
            .Order = GetRequest("order", .Order)
            .Read = GetRequest("read", "")
          End With

          If vobj_mod.Save Then
            Dim str_action As String = GetForm("action")
            If obj_tp.Sections.ContainsKey(str_action) Then
              obj_tp.Sections(str_action).AddNew()
            End If
          End If

        End If

        obj_tp.SetItem("head", vobj_mod.Plugin.Head)
        obj_tp.SetItem("data", vobj_mod.Plugin.Edit)
        ListPermissions(obj_tp, vobj_mod.Read)
      Catch ex As Exception
        obj_tp.SetItem("data", "ERROR: " & ex.Message)
      End Try
      Return obj_tp.Render
    End Function

    Public Function EditPage(ByVal int_id As Integer, ByVal vint_parent As Integer) As String
      Try  
        Dim b_edit As Boolean = (int_id > 0) And Pages.Contains(int_id)
        Dim obj_tp As New HTMLTemplate("/cms/plugins/admin_site/dialog.htm")

        If Request.Form.Count > 0 Then
          Dim obj_page As New Page
          If b_edit Then obj_page = Pages.Get(int_id)

          With obj_page

            Dim str_name As String = GetForm("name").Trim()
            If String.IsNullOrEmpty(str_name) Then
              .Title = "Untitled " & (Pages.Children(GetForm("parent", 0)).Count + 1)
              .Name = .Title.ToLower.Replace(" ", "-")
            Else
              .Name = str_name
              .Title = GetForm("title")
            End If

            .Parent = GetForm("parent", 0)
            .Template = GetForm("template", 0)
            .Visible = GetForm("visible", 0)
            .Landing = GetForm("landing", 0)
            .Application = GetForm("app", 0)
            .Order = GetForm("order", 0)
            .Read = GetForm("read")
            .Redirect = GetForm("redirect")

            .Meta_Title = GetForm("meta_title")
            .Meta_Description = GetForm("meta_description")
            .Meta_Keywords = GetForm("meta_keywords")

            For Each str_col As String In obj_page.Custom.Keys
              obj_page.Custom(str_col) = GetForm("custom_" & str_col)
            Next

          End With

          If obj_page.Save() Then
            obj_tp("redirect").AddItem("url", obj_page.Url)
          Else
            Return FormatError("Unable to update.", Data.Current.LastError)
          End If

        Else
          'reload pages
          Pages.Load()

          Dim obj_page As New Page
          If Pages.Contains(vint_parent) Then
            Dim obj_parent As Page = Pages.Get(vint_parent)
            obj_page.Parent = vint_parent
            obj_page.Template = obj_parent.Template
            obj_page.Read = obj_parent.Read
            obj_page.Visible = obj_parent.Visible
            obj_page.Meta_Title = obj_parent.Meta_Title
            obj_page.Meta_Description = obj_parent.Meta_Description
            obj_page.Meta_Keywords = obj_parent.Meta_Keywords
          End If
          If b_edit Then obj_page = Pages.Get(int_id)

          ListPermissions(obj_tp, obj_page.Read)
          Dim obj_tp2 As New HTMLTemplate("/cms/plugins/admin_site/page.htm")

          ListTemplates(obj_tp2, "templates", "page", obj_page.Template)
          If b_edit Then
            ListPages(obj_tp2, "parents", 0, 1, obj_page.Parent, obj_page.ID)
            ListPriorities(obj_tp2, "ordering", obj_page.Parent, obj_page.Order)
          Else
            ListPages(obj_tp2, "parents", 0, 1, vint_parent, 0)
            obj_tp2.CurrentPart = "ordering"
            obj_tp2.AddNew()
            obj_tp2.SetItem("value", -1)
            obj_tp2.SetItem("name", "Last")
            obj_tp2.CurrentPart = ""

            obj_page.Title = "Untitled " & (Pages.Children(vint_parent).Count + 1)
            obj_page.Name = obj_page.Title.ToLower.Replace(" ", "-")
          End If

          obj_tp2.SetItem("id", obj_page.ID)
          obj_tp2.SetItem("parent", obj_page.Parent)
          obj_tp2.SetItem("name", obj_page.Name)
          obj_tp2.SetItem("title", obj_page.Title)
          obj_tp2.SetItem("template", obj_page.Template)
          obj_tp2.SetItem("redirect", obj_page.Redirect)
          obj_tp2.SetItem("visible_" & obj_page.Visible, "selected=""selected""")
          obj_tp2.SetItem("landing_" & obj_page.Landing, "selected=""selected""")
          obj_tp2.SetItem("args_" & obj_page.Application, "selected=""selected""")

          obj_tp2.SetItem("meta_title", obj_page.Meta_Title)
          obj_tp2.SetItem("meta_description", obj_page.Meta_Description)
          obj_tp2.SetItem("meta_keywords", obj_page.Meta_Keywords)

          If Security.IsGlobal Then obj_tp2.SetItem("global", 1)

          obj_tp2.CurrentPart = ""

          For Each str_col As String In obj_page.Custom.Keys
            obj_tp2.SetItem("custom", 1)
            obj_tp2("custom").AddItem("title", StrConv(str_col, VbStrConv.ProperCase))
            obj_tp2("custom").SetItem("name", "custom_" & str_col)
            Select Case obj_page.Custom.DataType(str_col).ToString.ToLower
              Case "system.byte", "system.integer"
                obj_tp2("custom").SetItem("bool", 1)
                obj_tp2("custom").SetItem(IIf(ToInt(obj_page.Custom(str_col), 0) = 1, "true", "false"), 1)
            End Select
            If b_edit Then obj_tp2("custom").SetItem("value", obj_page.Custom(str_col))
          Next
          obj_tp2.CurrentPart = ""
          obj_tp.SetItem("data", obj_tp2.Render)
        End If
        Return obj_tp.Render
      Catch ex As Exception
        Return "A problem occured. Please try again." & ex.ToString & ""
      End Try
    End Function

    Public Sub ListPages(ByVal vobj_tp As HTMLTemplate, ByVal vstr_part As String, ByVal vint_parent As Integer, ByVal vint_depth As Integer, ByVal vint_compare As Integer, ByVal vint_id As Integer)
      If (Not vobj_tp Is Nothing) Then
        vobj_tp.CurrentPart = vstr_part
        ListPages(vobj_tp, vint_parent, vint_depth, vint_compare, vint_id)
        vobj_tp.CurrentPart = ""
      End If
    End Sub

    Public Sub ListPages(ByVal vobj_tp As HTMLTemplate, ByVal vint_parent As Integer, ByVal vint_depth As Integer, ByVal vint_compare As Integer, ByVal vint_id As Integer)
      On Error Resume Next
      If (Not vobj_tp Is Nothing) Then
        If (vint_depth = 1) Then
          vobj_tp.AddNew()
          vobj_tp.SetItem("value", 0)
          vobj_tp.SetItem("name", "ROOT")
          vobj_tp.SetItem("selected", IIf((vint_compare = 0), "selected", ""))
        End If
        Dim int_repeat As Integer = (vint_depth * 2)
        For Each obj_page As Page In Pages.Children(vint_parent)
          If (obj_page.ID <> vint_id) Then
            vobj_tp.AddNew()
            vobj_tp.SetItem("value", obj_page.ID)
            Dim str_name As String = obj_page.Name
            Dim int_len As Integer = (50 - int_repeat)
            If (((str_name.Length + int_repeat) > int_len) AndAlso (int_len < str_name.Length)) Then
              str_name = str_name.Substring(0, int_len)
            End If
            vobj_tp.SetItem("name", StrRepeat(".", int_repeat) & str_name)
            vobj_tp.SetItem("selected", IIf((obj_page.ID = vint_compare), "selected", ""))
            ListPages(vobj_tp, obj_page.ID, (vint_depth + 1), vint_compare, vint_id)
          End If
        Next
      End If
    End Sub

    Public Sub ListPriorities(ByVal vobj_tp As HTMLTemplate, ByVal vstr_part As String, ByVal vint_parent As Integer, ByVal vint_id As Integer)
      If (Not vobj_tp Is Nothing) Then
        vobj_tp.CurrentPart = vstr_part
        ListPriorities(vobj_tp, vint_parent, vint_id)
        vobj_tp.CurrentPart = ""
      End If
    End Sub

    Public Sub ListPriorities(ByVal vobj_tp As HTMLTemplate, ByVal vint_parent As Integer, ByVal vint_id As Integer)
      If (Not vobj_tp Is Nothing) Then
        Dim obj_pages As List(Of Page) = Pages.Children(vint_parent)
        Dim VBt_i4L0 As Integer = obj_pages.Count
        Dim i As Integer = 1
        Do While (i <= VBt_i4L0)
          vobj_tp.AddNew()
          vobj_tp.SetItem("value", i)
          vobj_tp.SetItem("name", i)
          If (i = vint_id) Then
            vobj_tp.SetItem("selected", "selected")
          End If
          i += 1
        Loop
      End If
    End Sub

    Public Sub ListTemplates(ByVal vobj_tp As HTMLTemplate, ByVal vstr_part As String, ByVal vstr_type As String, ByVal vint_selected As Integer)
      If (Not vobj_tp Is Nothing) Then
        vobj_tp.CurrentPart = vstr_part
        ListTemplates(vobj_tp, vstr_type, vint_selected)
        vobj_tp.CurrentPart = ""
      End If
    End Sub

    Public Sub ListTemplates(ByVal vobj_tp As HTMLTemplate, ByVal vstr_type As String, ByVal vint_selected As Integer)
      On Error Resume Next
      For Each obj_row As DataRow In Data.Current.Rows("select * from cms_templates where [type]='" & vstr_type & "' or [type]='' order by [name]")
        vobj_tp.AddNew()
        vobj_tp.SetItem("value", obj_row.Item("id"))
        vobj_tp.SetItem("name", NullSafe(obj_row.Item("name")).ToLower)
        If obj_row.Item("id") = vint_selected Then
          vobj_tp.SetItem("selected", "selected")
        End If
      Next
    End Sub

    'Private Function ListTemplates() As String
    '  On Error Resume Next
    '  Dim obj_sb As New Text.StringBuilder(String.Empty)
    '  obj_sb.AppendLine("{templates:[")
    '  Dim obj_table As DataTable = Data.Current.Table("select * from cms_templates where [type]='page' or [type]='' order by [name]")
    '  If obj_table.Rows.Count > 0 Then
    '    For Each obj_row As DataRow In obj_table.Rows
    '      obj_sb.Append("{")
    '      obj_sb.Append("id:" & JSONEncode(obj_row("id")) & ",")
    '      obj_sb.Append("text:" & JSONEncode(NullSafe(obj_row("name")).ToLower))
    '      obj_sb.AppendLine("},")
    '    Next
    '    obj_sb.Remove(obj_sb.Length - 3, 1)
    '  End If
    '  obj_sb.AppendLine("]}")
    '  Return obj_sb.ToString
    'End Function

    Private Function ListPages(ByVal vint_parent As Integer, ByVal vint_depth As Integer) As String
      On Error Resume Next
      Dim obj_sb As New Text.StringBuilder(String.Empty)
      Dim obj_pages As List(Of Page) = Global.Pages.Children(vint_parent)
      If obj_pages.Count > 0 Then
        obj_sb.AppendLine("[{")
        Dim str_children As String = String.Empty
        Dim str_item As String = StrRepeat("  ", vint_depth) & "{0}:{1}," & vbNewLine
        Dim b_first As Boolean = True
        For Each obj_page As Page In obj_pages
          If b_first Then
            b_first = False
          Else
            obj_sb.AppendLine("},")
            obj_sb.AppendLine("{")
          End If
          obj_sb.AppendFormat(str_item, "id", JSONEncode(obj_page.ID))
          obj_sb.AppendFormat(str_item, "text", JSONEncode(obj_page.Name))
          str_children = ListPages(obj_page.ID, vint_depth + 1)
          If NotEmpty(str_children) Then
            obj_sb.AppendFormat(str_item, "children", str_children)
          Else
            'obj_sb.AppendFormat(str_item, "leaf", "true")
            obj_sb.AppendFormat(str_item, "expanded", "true")
            obj_sb.AppendFormat(str_item, "children", "[]")
          End If
          obj_sb.AppendFormat(str_item, "url", JSONEncode(obj_page.Url))
          obj_sb.Remove(obj_sb.Length - 3, 1)
        Next
        obj_sb.Append("}]")
      End If
      If obj_sb.Length = 0 And vint_depth = 0 Then obj_sb.Append("{}")
      Return obj_sb.ToString
    End Function

    Public Sub ListPermissions(ByVal vobj_tp As HTMLTemplate, ByVal vstr_read As String)
      On Error Resume Next
      If Not vobj_tp Is Nothing Then
        If Security.IsInRoles(vstr_read, EveryRole) Then vobj_tp.SetItem("every_checked", "checked")
        If Security.IsInRoles(vstr_read, GuestRole) Then vobj_tp.SetItem("guest_checked", "checked")
        For Each obj_row As DataRow In Data.Current.Table("select * from [cms_roles]").Rows
          vobj_tp("permissions").AddItem("name", obj_row("name"))
          vobj_tp("permissions").SetItem("read_id", obj_row("id"))
          If Security.IsInRoles(vstr_read, obj_row("id")) Then vobj_tp("permissions").SetItem("read_checked", "checked")
        Next
      End If
    End Sub

  End Class

End Namespace