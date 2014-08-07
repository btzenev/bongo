Namespace Plugins.Admin

  Public Class User
    Inherits Admin.Base

    Public Overrides ReadOnly Property Name() As String
      Get
        Return "User Management"
      End Get
    End Property

#Region " Common "

    Public Overrides Function Render() As String
      Dim int_id As Integer = GetQuery("id", 0)
      Select Case Args(2)
        Case "roles"
          Select Case Args(3)
            Case "edit", "add", "save" : Return Role_Edit(int_id)
            Case "del" : Return Role_Del(int_id)
            Case "clear" : Return Role_Clear(int_id)
            Case Else : Return Roles()
          End Select
        Case "edit", "add", "save" : Return Edit(int_id)
        Case "del" : Return Del(int_id)
        Case Else : Return List()
      End Select
    End Function

    Public Overrides Sub Menu(ByRef vobj_tp As HTMLTemplate)
      vobj_tp("sections").AddItem("name", "User Management")
      vobj_tp("sections.links").AddItem("name", "Add User ")
      vobj_tp("sections.links").SetItem("link", "/cms/user/add.aspx")
      vobj_tp("sections.links").AddItem("name", "List Users ")
      vobj_tp("sections.links").SetItem("link", "/cms/user/default.aspx")
      vobj_tp("sections").AddItem("name", "Role Management")
      vobj_tp("sections.links").AddItem("name", "Add Role ")
      vobj_tp("sections.links").SetItem("link", "/cms/user/roles/add.aspx")
      vobj_tp("sections.links").AddItem("name", "List Roles ")
      vobj_tp("sections.links").SetItem("link", "/cms/user/roles.aspx")
    End Sub

#End Region

#Region " User "

    Private Function Del(ByVal vint_id As Integer) As String
      If vint_id > 0 Then
        If Data.Current.Delete("cms_users", "id", vint_id) Then
          Response.Redirect("/cms/user/default.aspx", True)
          Return String.Empty
        End If
      End If
      Return "Unable to delete user"
    End Function

    Private Function List()
      Dim str_links As String = "<a href=""/cms/user/edit.aspx?id={id}"">edit</a>&nbsp;<a href=""/cms/user/del.aspx?id={id}"" onclick=""return confirm('Are you sure?')"">delete</a>"
      Dim obj_grid As New Grid()
      obj_grid.Title = "List Users"
      obj_grid.AddColumn("id", "ID", True)
      obj_grid.AddColumn("userid", "UserID", True)
      obj_grid.AddColumn("created", "Created", True, AddressOf Grid.Formatters.ShortDate)
      obj_grid.AddHTMLColumn("addnew", "<a href=""/cms/user/add.aspx"">add new</a>", str_links)
      Return obj_grid.Render("cms_users")
    End Function

    Private Function Edit(ByVal vint_id As Integer) As String
      'on error resume next
      Dim obj_schema As DataColumnCollection = Data.Current.Schema("cms_users")
      Dim obj_edit As New Editor("cms_users", True)
      obj_edit.Title = IIf(vint_id > 0, "Edit", "Add") & " User"
      obj_edit.Redirect = "/cms/user/edit.aspx?id={0}"
      obj_edit.AddView("id", "ID")
      Dim obj_table As DataTable = Data.Current.Table("select * from [cms_roles]")
      If obj_schema.Contains("roles") Then
        Dim obj_roles As New RolesColumn()
        obj_roles.Add("Administrator", "-1")
        obj_roles.Add("Guest", "0")
        For Each obj_row As DataRow In obj_table.Rows
          obj_roles.Add(obj_row("name"), obj_row("id"))
        Next
        obj_edit.AddColumn("roles", "Roles", obj_roles)
        obj_edit.Skip("roleid")
      Else 'roleid
        Dim obj_drop As New Editor.DropColumn()
        obj_drop.Add("Administrator", "-1")
        obj_drop.Add("Guest", "0")
        For Each obj_row As DataRow In obj_table.Rows
          obj_drop.Add(obj_row("name"), obj_row("id"))
        Next
        obj_edit.AddColumn("roleid", "RoleID", obj_drop)
      End If
      obj_edit.AddInput("userid", "Username")
      obj_edit.AddPassword("password", "Password")
      If obj_schema.Contains("enabled") Then obj_edit.AddByte("enabled", "Enabled")
      obj_edit.AddButton("Cancel", "document.location='/cms/user/default.aspx'")
      Return obj_edit.Render(vint_id)
    End Function

#End Region

#Region " Role "

    Private Function Role_Del(ByVal vint_id As Integer) As String
      If Data.Current.Delete("cms_roles", "id", vint_id) Then
        RemoveRole("cms_pages", vint_id)
        RemoveRole("cms_modules", vint_id)
        Response.Redirect("/cms/user/roles.aspx", True)
        Return String.Empty
      Else
        Return "Unable to delete role"
      End If
    End Function

    Public Function Role_Clear(ByVal vint_id As Integer) As String
      If vint_id > 0 Then
        RemoveRole("cms_pages", vint_id)
        RemoveRole("cms_modules", vint_id)
        Pages.Load()
        Response.Redirect("/cms/user/roles/edit.aspx?id=" & vint_id, True)
      End If
      Return "Unable to clear role"
    End Function

    Private Function Roles()
      Dim str_links As String = "<a href=""/cms/user/roles/edit.aspx?id={id}"">edit</a>&nbsp;<a href=""/cms/user/roles/del.aspx?id={id}"" onclick=""return confirm('Are you sure?')"">delete</a>"
      Dim obj_grid As New Grid()
      obj_grid.Title = "List Roles"
      obj_grid.AddColumn("id", "ID", True)
      obj_grid.AddColumn("name", "Name", True)
      obj_grid.AddHTMLColumn("addnew", "<a href=""/cms/user/roles/add.aspx"">add new</a>", str_links)
      Return obj_grid.Render("cms_roles")
    End Function

    Private Function Role_Edit(ByVal vint_id As Integer) As String
      Dim obj_schema As DataColumnCollection = Data.Current.Schema("cms_roles")
      Dim obj_edit As New Editor("cms_roles", AddressOf CopyRoles)
      obj_edit.Title = IIf(vint_id > 0, "Edit", "Add") & " Role"
      If obj_schema.Contains("access") Then
        obj_edit.AddColumn("access", "Admin Privileges", New PluginColumn)
      End If
      obj_edit.AddHTML("copyperms", "Copy Permissions", ListRoles(vint_id))
      obj_edit.Redirect = "/cms/user/roles/edit.aspx?id={0}"
      If vint_id > 0 Then
        obj_edit.AddButton("Clear Permissions", "if (confirm('Are you sure?')) document.location='/cms/user/roles/clear.aspx?id=" & vint_id & "'")
      End If
      obj_edit.AddButton("Cancel", "document.location='/cms/user/roles.aspx'")
      Return obj_edit.Render(vint_id)
    End Function

    Private Function ListRoles(ByVal vint_id As Integer) As String
      On Error Resume Next
      Dim obj_sb As New System.Text.StringBuilder
      obj_sb.Append("<script type=""text/javascript"">" & vbNewLine)
      obj_sb.Append("function copyperms(el){" & vbNewLine)
      obj_sb.Append("if (el.selectedIndex > 0) {" & vbNewLine)
      obj_sb.Append("  if (confirm('Are you sure?')) {" & vbNewLine)
      obj_sb.Append("    return true;" & vbNewLine)
      obj_sb.Append("  }" & vbNewLine)
      obj_sb.Append("}" & vbNewLine)
      obj_sb.Append("el.selectedIndex=0;" & vbNewLine)
      obj_sb.Append("}" & vbNewLine)
      obj_sb.Append("</script>" & vbNewLine)
      obj_sb.Append("<select name=""copypermissions"" onchange=""return copyperms(this)"">" & vbNewLine)
      obj_sb.Append("<option value=""-1"" selected>Select to copy another roles permissions</option>" & vbNewLine)
      obj_sb.Append("<option value=""0"">Guest</option>" & vbNewLine)
      Dim obj_table As DataTable = Data.Current.Table("select * from [cms_roles] where [id]<>" & vint_id)
      For Each obj_row As DataRow In obj_table.Rows
        obj_sb.Append("<option value=""" & obj_row("id") & """>" & obj_row("name") & "</option>" & vbNewLine)
      Next
      obj_sb.Append("</select>" & vbNewLine)
      obj_sb.Append("<script type=""text/javascript"">" & vbNewLine)
      obj_sb.Append("document.forms[0].copypermissions.selectedIndex=0;" & vbNewLine)
      obj_sb.Append("</script>" & vbNewLine)
      Return obj_sb.ToString
    End Function

    Private Function CopyRoles(ByVal vint_id As Integer) As String
      Dim int_role As Integer = GetForm("copypermissions", -5555)
      If Not int_role = -5555 Then
        AppendRole("cms_pages", int_role, vint_id)
        AppendRole("cms_modules", int_role, vint_id)
        Pages.Load()
      End If
    End Function

    Private Sub AppendRole(ByVal vstr_table As String, ByVal vint_find As Integer, ByVal vint_add As Integer)
      'On Error Resume Next
      Dim obj_table As DataTable = Data.Current.Table("select [id],[read] from [" & vstr_table & "]")
      Dim str_read As String
      For Each obj_row As DataRow In obj_table.Rows
        str_read = "" & obj_row("read")
        If str_read.StartsWith(vint_find & ";") Or str_read.IndexOf(";" & vint_find & ";") >= 0 Then
          If Not str_read.StartsWith(vint_add & ";") And str_read.IndexOf(";" & vint_add & ";") = -1 Then
            str_read = str_read.Trim(";") & ";" & vint_add & ";"
            Data.Current.Begin()
            Data.Current.Execute("update [" & vstr_table & "] set [read]='" & str_read & "' where [id]=" & obj_row("id"))
            Data.Current.Commit()
          End If
        End If
      Next
    End Sub

    Private Sub RemoveRole(ByVal vstr_table As String, ByVal vint_remove As Integer)
      'On Error Resume Next
      Dim obj_table As DataTable = Data.Current.Table("select [id],[read] from [" & vstr_table & "]")
      Dim str_read As String
      For Each obj_row As DataRow In obj_table.Rows
        str_read = "" & obj_row("read")
        If str_read.StartsWith(vint_remove & ";") Or str_read.IndexOf(";" & vint_remove & ";") >= 0 Then
          If str_read.StartsWith(vint_remove & ";") Then
            str_read = str_read.Substring((vint_remove & ";").Length)
          Else
            str_read = str_read.Replace(";" & vint_remove & ";", ";")
          End If
          str_read = ";" & str_read.Trim(";") & ";"
          'pages(obj_row("id")).read = str_read
          Data.Current.Begin()
          Data.Current.Execute("update [" & vstr_table & "] set [read]='" & str_read & "' where [id]=" & obj_row("id"))
          Data.Current.Commit()
        End If
      Next
    End Sub

#End Region

#Region " Editor Plugins "

#Region " RoleColumn "

    Public Class RolesColumn
      Inherits Editor.BaseColumn

#Region " Declarations "

      Private mobj_items As New ArrayList

      Private Class Item
        Public Name As String
        Public Value As String
      End Class

#End Region

#Region " Add "

      Public Sub Add(ByVal vstr_value As String)
        Dim obj_item As New Item
        obj_item.Name = vstr_value
        obj_item.Value = vstr_value
        mobj_items.Add(obj_item)
      End Sub

      Public Sub Add(ByVal vstr_name As String, ByVal vstr_value As String)
        Dim obj_item As New Item
        obj_item.Name = vstr_name
        obj_item.Value = vstr_value
        mobj_items.Add(obj_item)
      End Sub

#End Region

      Public Overrides Function Render(ByVal vstr_value As String) As String
        Try
          Dim obj_sb As New StringBuilder
          obj_sb.AppendLine("<select id=""{field}"" name=""{field}"" size=""10"" style=""width:100%;"" multiple class=""multisel""><%options%>")
          obj_sb.AppendLine("<option value=""{{value}}""{{selected}}>{{name}}</option><%options%>")
          obj_sb.AppendLine("</select>")
          Dim obj_tp As New HTMLTemplate
          obj_tp.HTML = obj_sb.ToString
          vstr_value = vstr_value.TrimStart(":"c).TrimEnd(":"c)
          Dim obj_refs As New ArrayList(vstr_value.Split(":"))
          For Each obj_item As Item In mobj_items
            obj_tp("options").AddItem("value", obj_item.Value)
            obj_tp("options").SetItem("name", obj_item.Name)
            If obj_refs.Contains(obj_item.Value) Then
              obj_tp("options").SetItem("selected", " selected=""selected""")
            End If
          Next
          'obj_tp.SetItem("options", Join(obj_refs.ToArray, ":"))
          Return obj_tp.Render
        Catch ex As Exception
          Return ex.ToString
        End Try
      End Function

      Public Overrides Function Save(ByVal vstr_value As String) As String
        vstr_value = vstr_value.Replace(",", ":").Replace(";", ":").Trim(":")
        If Not String.IsNullOrEmpty(vstr_value) Then vstr_value = ":" & vstr_value & ":"
        Return vstr_value
      End Function

    End Class

#End Region

#Region " PluginColumn "

    Public Class PluginColumn
      Inherits Editor.BaseColumn
      Private mstr_sql As String

      Public Overrides Function Render(ByVal vstr_value As String) As String
        Try
          Dim obj_sb As New StringBuilder
          obj_sb.AppendLine("<select id=""{field}"" name=""{field}"" size=""10"" style=""width:100%;"" multiple class=""multisel""><%options%>")
          obj_sb.AppendLine("<option value=""{{value}}""{{selected}}>{{name}}</option><%options%>")
          obj_sb.AppendLine("</select>")

          Dim obj_tp As New HTMLTemplate
          obj_tp.HTML = obj_sb.ToString
          vstr_value = vstr_value.TrimStart(":"c).TrimEnd(":"c)
          Dim obj_refs As New ArrayList(vstr_value.Split(":"))

          Dim obj_plugins As New List(Of String)(Plugins.Admin.List)
          obj_plugins.Sort()
          If obj_plugins.Contains("admin_data") And Not Security.IsGlobal Then obj_plugins.Remove("admin_data")
          If obj_plugins.Contains("admin_tools") And Not Security.IsGlobal Then obj_plugins.Remove("admin_tools")
          If obj_plugins.Contains("admin_user") And Not Security.IsGlobal Then obj_plugins.Remove("admin_user")

          For Each str_plugin As String In obj_plugins
            obj_tp("options").AddItem("value", "" & str_plugin)
            obj_tp("options").SetItem("name", "" & str_plugin)
            If obj_refs.Contains("" & str_plugin) Then
              obj_tp("options").SetItem("selected", " selected=""selected""")
            End If
          Next
          obj_tp.SetItem("options", Join(obj_refs.ToArray, ":"))
          Return obj_tp.Render
        Catch ex As Exception
          Return ex.ToString
        End Try
      End Function

      Public Overrides Function Save(ByVal vstr_value As String) As String
        vstr_value = vstr_value.Replace(",", ":").Replace(";", ":").Trim(":")
        If Not String.IsNullOrEmpty(vstr_value) Then vstr_value = ":" & vstr_value & ":"
        Return vstr_value
      End Function

    End Class

#End Region

#End Region

  End Class

End Namespace