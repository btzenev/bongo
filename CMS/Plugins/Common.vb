Imports System.Web.Compilation

Namespace Plugins

  Friend Class PluginInfo
    Public Path As String
    Public File As String
    Public FullName As String
    Public Modified As Date
    Public Integrated As Boolean = False
    Public Type As Type
  End Class

  Public Module Common

    Private mobj_plugins As New Dictionary(Of String, PluginInfo)

    Public Function Load()
      Add(Plugins.Admin.List, "admin_site", GetType(Admin.Site))
      Add(Plugins.Admin.List, "admin_user", GetType(Admin.User))
      'Add(Plugins.Admin.List, "admin_data", GetType(Admin.DataMan))
      Add(Plugins.Admin.List, "admin_asset", GetType(Admin.Asset))
      'Add(Plugins.Admin.List, "admin_tools", GetType(Admin.Tools))
      Plugins.Load(Plugins.Admin.List, "admin")

      Add(Plugins.Client.List, "client_content", GetType(Client.Content))
      Add(Plugins.Client.List, "client_sitemap", GetType(Client.Sitemap))
      Add(Plugins.Client.List, "client_subtree", GetType(Client.Subtree))
      Add(Plugins.Client.List, "client_include", GetType(Client.Include))
      Add(Plugins.Client.List, "client_login", GetType(Client.Login))
      Plugins.Load(Plugins.Client.List, "client")

      Add(Plugins.Hidden.List, "hidden_subtree", GetType(Hidden.Subtree))
      Add(Plugins.Hidden.List, "hidden_crumbs", GetType(Hidden.Crumbs))
      Add(Plugins.Hidden.List, "hidden_root", GetType(Hidden.Root))
      Plugins.Load(Plugins.Hidden.List, "hidden")
    End Function

    Private Sub Add(ByRef vobj_plugins As List(Of String), ByVal vstr_name As String, ByVal vobj_type As Type)
      SyncLock mobj_plugins
        If mobj_plugins.ContainsKey(vstr_name) Then mobj_plugins.Remove(vstr_name)
        Dim obj_info As New PluginInfo With {.Integrated = True, .Type = vobj_type}
        mobj_plugins.Add(vstr_name, obj_info)
      End SyncLock

      SyncLock vobj_plugins
        If Not vobj_plugins.Contains(vstr_name) Then vobj_plugins.Add(vstr_name)
      End SyncLock
    End Sub

    Public Function [Get](ByVal vstr_name As String) As Type
      Dim obj_ex As System.Exception = Nothing
      Return [Get](vstr_name, obj_ex)
    End Function

    Public Function [Get](ByVal vstr_name As String, ByRef vobj_ex As System.Exception) As Type
      Try
        If mobj_plugins.ContainsKey(vstr_name) Then
          Dim obj_info As PluginInfo = mobj_plugins(vstr_name)
          If Not obj_info.Integrated Then
            Dim b_build As Boolean = obj_info.Type Is Nothing
            Dim obj_date As Date = IO.File.GetLastWriteTime(obj_info.File)
            If Not b_build Then b_build = Not obj_date = obj_info.Modified
            If b_build Then
              Try
                Dim obj_asm As Reflection.Assembly = Compilation.BuildManager.GetCompiledAssembly(obj_info.Path)
                obj_info.Type = obj_asm.GetType(obj_info.FullName, True, True)
                obj_info.Modified = obj_date
                mobj_plugins(vstr_name) = obj_info
              Catch ex As Exception
                vobj_ex = ex
                Return Nothing
              End Try
            End If
          End If
          Return obj_info.Type
        Else
          Throw New System.Exception(vstr_name & " - no plugin by that name")
        End If
      Catch ex As Exception
        vobj_ex = ex
      End Try
      Return Nothing
    End Function

    Public Function Load(ByRef vobj_plugins As List(Of String), ByVal vstr_type As String) As Boolean
      On Error Resume Next

      Dim str_dir As String = MapPath("/App_Plugins/")
      If IO.Directory.Exists(str_dir) Then
        For Each str_file As String In IO.Directory.GetFiles(str_dir, vstr_type & "_*.vb")
          Dim str_name As String = IO.Path.GetFileNameWithoutExtension(str_file).ToLower
          If Not str_name = vstr_type Then
            Dim str_path As String = UnMappath(str_file)

            SyncLock mobj_plugins
              If mobj_plugins.ContainsKey(str_name) Then mobj_plugins.Remove(str_name)
              Dim obj_info As New PluginInfo With {
                .FullName = "plugins." & vstr_type & "." & str_name,
                .Path = str_path,
                .File = str_file,
                .Modified = IO.File.GetLastWriteTime(str_file)
              }
              mobj_plugins.Add(str_name, obj_info)
            End SyncLock

            SyncLock vobj_plugins
              If Not vobj_plugins.Contains(str_name) Then vobj_plugins.Add(str_name)
            End SyncLock
          End If
        Next
      End If
    End Function


  End Module

End Namespace