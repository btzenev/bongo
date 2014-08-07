Namespace Plugins.Admin

    Public Class Asset
        Inherits Admin.Base

        Public Overrides ReadOnly Property Name() As String
            Get
                Return "Asset Management"
            End Get
        End Property

        Public Overrides ReadOnly Property Full() As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides Function Render() As String
            If Security.IsAdmin Then
                If Args.Count > 2 Then
                    Dim str_msg As String = String.Empty
                    Dim b_success As Boolean = False
                    Select Case Args(2)
                        Case "create" : b_success = Create(str_msg)
                        Case "move" : b_success = Move(str_msg)
                        Case "rename" : b_success = Rename(str_msg)
                        Case "delete" : b_success = Delete(str_msg)
                        Case "upload" : b_success = Upload(str_msg)
                        Case "save" : b_success = Save(str_msg)
                        Case "download" : Download() : Return String.Empty
                        Case "read" : Download(True) : Return String.Empty
                        Case "folders" : Return Folders()
                        Case "files" : Return Files()
                        Case Else : Return "Bad Request"
                    End Select
                    Return "{""message"":""" & str_msg & """,""success"":" & b_success.ToString.ToLower & "}"
                Else
                    Return HTMLTemplate.QuickRender("/cms/plugins/admin_asset/template.htm")
                End If
            End If
            Return HTMLTemplate.QuickRender("/cms/templates/permissions.htm")
        End Function

        Private Function Rename(ByRef vstr_msg As String) As Boolean
            Dim str_new As String = GetQuery("new")
            Dim str_file As String = GetQuery("file")
            Dim str_dir As String = GetQuery("dir")
            If Not String.IsNullOrEmpty(str_new) Then
                If Not String.IsNullOrEmpty(str_file) Then
                    Dim obj_file As New FileInfo(MapPath(str_file))
                    If obj_file.Exists Then
                        str_new = IO.Path.Combine(obj_file.DirectoryName, str_new)
                        Try
                            obj_file.MoveTo(str_new)
                            vstr_msg = "File renamed"
                            Return True
                        Catch ex As Exception
                            vstr_msg = "Unable to rename the file"
                        End Try
                    Else
                        vstr_msg = "File does not exist"
                    End If
                ElseIf Not String.IsNullOrEmpty(str_dir) Then
                    Dim obj_dir As New DirectoryInfo(MapPath(str_dir))
                    If obj_dir.Exists Then
                        str_new = IO.Path.Combine(obj_dir.Parent.FullName, str_new)
                        Try
                            obj_dir.MoveTo(str_new)
                            vstr_msg = "Directory renamed"
                            Return True
                        Catch ex As Exception
                            vstr_msg = "Unable to rename the directory"
                        End Try
                    Else
                        vstr_msg = "Directory does not exist"
                    End If
                End If
            End If
            Return False
        End Function

        Private Function Delete(ByRef vstr_msg As String) As Boolean
            Dim str_file As String = GetQuery("file")
            Dim str_dir As String = GetQuery("dir")
            If Not String.IsNullOrEmpty(str_file) Then
                str_file = MapPath(str_file)
                If IO.File.Exists(str_file) Then
                    Try
                        IO.File.Delete(str_file)
                        vstr_msg = "File deleted"
                        Return True
                    Catch ex As Exception
                        vstr_msg = "Unable to delete the file"
                    End Try
                Else
                    vstr_msg = "File does not exist"
                End If
            ElseIf Not String.IsNullOrEmpty(str_dir) Then
                str_dir = MapPath(str_dir)
                Try
                    IO.Directory.Delete(str_dir, True)
                    vstr_msg = "Directory deleted"
                    Return True
                Catch ex As Exception
                    vstr_msg = "Unable to delete the file"
                End Try
            Else
                vstr_msg = "Directory does not exist"
            End If
            Return False
        End Function

        Private Function Move(ByRef vstr_msg As String) As Boolean
            Dim str_new As String = GetQuery("new")
            Dim str_file As String = GetQuery("file")
            Dim str_dir As String = GetQuery("dir")
            If Not String.IsNullOrEmpty(str_new) Then
                str_new = MapPath(str_new)
                If Not String.IsNullOrEmpty(str_file) Then
                    Dim obj_file As New FileInfo(MapPath(str_file))
                    If obj_file.Exists Then
                        str_new = IO.Path.Combine(str_new, obj_file.Name)
                        Try
                            obj_file.MoveTo(str_new)
                            vstr_msg = "File move"
                            Return True
                        Catch ex As Exception
                            vstr_msg = "Unable to move the file"
                        End Try
                    Else
                        vstr_msg = "File does not exist"
                    End If
                ElseIf Not String.IsNullOrEmpty(str_dir) Then
                    Dim obj_dir As New DirectoryInfo(MapPath(str_dir))
                    If obj_dir.Exists Then
                        str_new = IO.Path.Combine(str_new, obj_dir.Name)
                        Try
                            obj_dir.MoveTo(str_new)
                            vstr_msg = "Directory move"
                            Return True
                        Catch ex As Exception
                            vstr_msg = "Unable to move the directory"
                        End Try
                    Else
                        vstr_msg = "Directory does not exist"
                    End If
                End If
            End If
            Return False
        End Function

        Private Function Create(ByRef vstr_msg As String) As Boolean
            Dim str_dir As String = MapPath(GetQuery("dir", "/"))
            Dim str_new As String = GetQuery("new")
            If Not String.IsNullOrEmpty(str_new) Then
                If IO.Directory.Exists(str_dir) Then
                    str_new = IO.Path.Combine(str_dir, str_new)
                    If Not IO.Directory.Exists(str_new) Then
                        Try
                            IO.Directory.CreateDirectory(str_new)
                            vstr_msg = "Directory created"
                            Return True
                        Catch ex As Exception
                            vstr_msg = "Unable to created directory"
                        End Try
                    Else
                        vstr_msg = "Directory already exists"
                    End If
                Else
                    vstr_msg = "Directory does not exist"
                End If
            End If
            Return False
        End Function

        Private Function Upload(ByRef vstr_msg As String) As Boolean
            If Request.Files.Count > 0 Then
                Dim obj_file As HttpPostedFile = Request.Files(0)
                Dim str_path As String = MapPath(GetRequest("url", "/uploads/"))
                Dim str_file As String = IO.Path.Combine(str_path, IO.Path.GetFileName(obj_file.FileName))
                If IO.File.Exists(str_file) Then IO.File.Delete(str_file)
                Try
                    obj_file.SaveAs(str_file)
                    vstr_msg = "File uploaded"
                    Return True
                Catch ex As Exception
                    vstr_msg = IIf(GetQuery("debug", 0) = 1, ex.Message, "Unable to upload file")
                End Try
            Else
                vstr_msg = "No files uploaded"
            End If
            Return False
        End Function

        Private Function Save(ByRef vstr_msg As String) As Boolean
            Dim str_file As String = GetQuery("url")
            If Not String.IsNullOrEmpty(str_file) Then
                str_file = MapPath(str_file)
                If IO.File.Exists(str_file) Then
                    Try
                        Using obj_file As New StreamWriter(str_file)
                            obj_file.Write(GetForm("value"))
                        End Using
                        vstr_msg = "File saved"
                        Return True
                    Catch ex As Exception
                        vstr_msg = "Unable to save file"
                    End Try
                Else
                    vstr_msg = "File does not exist"
                End If
            End If
            Return False
        End Function

        Private Sub Download(Optional ByVal vb_read As Boolean = False)
            On Error Resume Next
            Dim str_file As String = GetQuery("url")
            If Not String.IsNullOrEmpty(str_file) Then
                str_file = MapPath(str_file)
                If IO.File.Exists(str_file) Then
                    If vb_read Then
                        Response.ContentType = "text/plain"
                    Else
                        Response.ContentType = "application/octet-stream"
                        Response.AddHeader("content-disposition", "attachment; filename=" & IO.Path.GetFileName(str_file))
                    End If
                    Response.WriteFile(str_file)
                    'Response.End()
                End If
            End If
        End Sub

        Private Function Folders() As String
            On Error Resume Next
            Dim str_name As String
            Dim obj_sb As New Text.StringBuilder()
            Dim str_node As String = GetRequest("node", "/")
            obj_sb.Append("[")
            For Each str_folder As String In IO.Directory.GetDirectories(MapPath(str_node))
                str_name = IO.Path.GetFileName(str_folder)
                If Not str_name.ToLower = "cms" Or Security.IsGlobal Then
                    obj_sb.AppendLine("{")
                    obj_sb.AppendLine("  id:'" & JSEncode(UnMappath(str_folder)) & "',")
                    obj_sb.AppendLine("  text:'" & JSEncode(str_name) & "',")
                    obj_sb.AppendLine("  cls:'folder'")
                    obj_sb.Append("},")
                End If
            Next
            If obj_sb.Length > 1 Then obj_sb.Remove(obj_sb.Length - 1, 1)
            obj_sb.AppendLine("]")
            Return obj_sb.ToString
        End Function

        Private Function Files() As String
            On Error Resume Next
            Dim str_name As String
            Dim obj_sb As New Text.StringBuilder()
            Dim str_type As String = GetRequest("type", "")
            Dim str_node As String = GetRequest("folder", "/")
            obj_sb.AppendLine("{")
            obj_sb.AppendLine("files: [")
            Dim str_ext As String
            Dim b_add As Boolean
            Dim obj_file As IO.FileInfo
            For Each str_file As String In IO.Directory.GetFiles(MapPath(str_node))
                str_name = IO.Path.GetFileName(str_file)
                str_ext = IO.Path.GetExtension(str_file)
                b_add = False
                Select Case str_type
                    Case "image"
                        Select Case str_ext
                            Case ".jpg", ".jpeg", ".png", ".bmp", ".gif"
                                b_add = True
                        End Select
                    Case Else
                        b_add = True
                End Select
                If b_add Then
                    obj_file = New IO.FileInfo(str_file)
                    obj_sb.AppendLine("{")
                    obj_sb.AppendLine("  name:'" & JSEncode(str_name) & "',")
                    obj_sb.AppendLine("  size:'" & JSEncode(FileSize(obj_file)) & "',")
                    obj_sb.AppendLine("  modified:'" & JSEncode(obj_file.LastWriteTime.ToString) & "',")
                    obj_sb.AppendLine("  url:'" & JSEncode(UnMappath(str_file)) & "'")
                    obj_sb.Append("},")
                End If
            Next
            obj_sb.Remove(obj_sb.Length - 1, 1)
            obj_sb.AppendLine("]")
            obj_sb.AppendLine("}")
            Return obj_sb.ToString
        End Function


    End Class

End Namespace