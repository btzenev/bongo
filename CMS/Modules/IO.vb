Public Module _IO

#Region " ReadFile "

    Public Function ReadFile(ByVal vstr_file As String) As String
        On Error Resume Next
        Using obj_file As TextReader = File.OpenText(MapPath(vstr_file))
            Return obj_file.ReadToEnd
        End Using
    End Function

#End Region

#Region " FileLength "

    Public Function FileLength(ByVal vstr_file As String) As Integer
        On Error Resume Next
        Dim obj_fileinfo As New System.IO.FileInfo(MapPath(vstr_file))
        Return obj_fileinfo.Length
    End Function

#End Region

#Region " FileSize "

    Public Function FileSize(ByVal vstr_file As String) As String
        On Error Resume Next
        Return FileSize(FileLength(vstr_file))
    End Function

    Public Function FileSize(ByVal vobj_file As IO.FileInfo) As String
        On Error Resume Next
        Return FileSize(vobj_file.Length)
    End Function

    Public Function FileSize(ByVal vlng_length As Long) As String
        ' if lenght>1MB, show size in MB
        Const ONE_MB As Long = 1024L * 1024L
        If (vlng_length > ONE_MB) Then
            Return ((((vlng_length * 100) / ONE_MB)) / 100).ToString("#,##0.0") + " MB"
        End If
        ' if length>1KB, show size in KB
        Const ONE_KB As Long = 1024L
        If (vlng_length > ONE_KB) Then
            Return ((((vlng_length * 100) / ONE_KB)) / 100).ToString("#,##0.0") + " KB"
        End If
        ' show size in bytes
        Return vlng_length.ToString("#,###,##0") + " bytes"
    End Function

#End Region

#Region " PathExists "

    Public Function PathExists(ByVal vstr_path As String) As Boolean
        On Error Resume Next
        vstr_path = MapPath(vstr_path)
        Return System.IO.Directory.Exists(vstr_path) Or System.IO.File.Exists(vstr_path)
    End Function

#End Region

#Region " Log "

    Public Sub Log(ByVal vobj_ex As System.Exception)
        Log(vobj_ex.Message, vobj_ex)
    End Sub

    Public Sub Log(ByVal vstr_msg As String, ByVal vobj_ex As System.Exception)
        On Error Resume Next
        Dim obj_sb As New StringBuilder
        If vobj_ex IsNot Nothing Then
            obj_sb.AppendLine("Source : " & vobj_ex.Source.ToString().Trim())
            obj_sb.AppendLine("Method : " & vobj_ex.TargetSite.Name.ToString())
            If vstr_msg <> vobj_ex.Message Then
                obj_sb.AppendLine("Error  : " & vobj_ex.Message.ToString().Trim())
            End If
            obj_sb.AppendLine("Stack  : " & vobj_ex.StackTrace.ToString().Trim())
        End If
        Log(vstr_msg, obj_sb.ToString)
    End Sub

    Public Sub Log(ByVal vstr_msg As String)
        Log(vstr_msg, String.Empty)
    End Sub

    Public Sub Log(ByVal vstr_msg As String, ByVal vstr_details As String)
        Try
            If ToBoolean(Config("cms.log"), True) Then
                Dim str_file As String = MapPath(AppPath, "cms.log")
                Using obj_stream As New FileStream(str_file, FileMode.Append, FileAccess.Write)
                    'obj_stream.Position = 0
                    Using obj_writer As New StreamWriter(obj_stream)
                        obj_writer.WriteLine(Date.Now.ToString & " - " & vstr_msg)
                        If Not String.IsNullOrEmpty(vstr_details) Then
                            vstr_details = "    " & vstr_details.Trim.Replace(vbNewLine, vbNewLine & "    ")
                            obj_writer.WriteLine(vstr_details.TrimEnd)
                        End If
                    End Using
                End Using
            End If
        Catch ex As Exception
            Throw New System.SystemException("Unable to write to log file")
        End Try
    End Sub

#End Region

End Module
