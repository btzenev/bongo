Imports System
Imports System.text

Namespace Context.CMS

  Public Class [Select]
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      SetCache(HttpCacheability.NoCache)
      Dim str_action As String = GetQuery("action")
      If Not String.IsNullOrEmpty(str_action) Then
        Dim str_name As String
        Select Case str_action
          Case "files"
            Dim obj_sb As New Text.StringBuilder()
            Dim str_type As String = GetRequest("type", "")
            Dim str_node As String = GetRequest("folder", "/")
            Dim str_filter As String = GetRequest("filter").ToLower
            obj_sb.Append("{")
            obj_sb.Append("""files"":[")

            Dim str_ext As String
            Dim b_add As Boolean
            Dim obj_file As IO.FileInfo
            For Each str_file As String In IO.Directory.GetFiles(MapPath(str_node))
              str_name = IO.Path.GetFileName(str_file)
              str_ext = IO.Path.GetExtension(str_file).ToLower
              b_add = False
              Select Case str_type
                Case "image"
                  Select Case str_ext
                    Case ".jpg", ".jpeg", ".png", ".bmp", ".gif"
                      b_add = True
                  End Select
                Case Else
                  Select Case str_filter
                    Case "images"
                      Select Case str_ext
                        Case ".jpg", ".jpeg", ".png", ".bmp", ".gif"
                          b_add = True
                      End Select
                    Case "documents"
                      Select Case str_ext
                        Case ".doc", ".docx", ".xls", ".xlsx", ".pdf", ".txt", ".rtf", ".pps"
                          b_add = True
                      End Select
                    Case Else : b_add = True
                  End Select
              End Select
              If b_add Then
                obj_file = New IO.FileInfo(str_file)
                obj_sb.Append("{")
                obj_sb.Append("""name"":""" & JSEncode(str_name) & """,")
                obj_sb.Append("""size"":""" & JSEncode(ToFileSize(obj_file.Length)) & """,")
                obj_sb.Append("""modified"":""" & JSEncode(obj_file.LastWriteTime.ToString) & """,")
                obj_sb.Append("""url"":""" & JSEncode(UnMappath(str_file)) & """")
                obj_sb.Append("},")
              End If
            Next
            If obj_sb.ToString.EndsWith(",") Then obj_sb.Remove(obj_sb.Length - 1, 1)
            obj_sb.Append("]")
            obj_sb.Append("}")
            Response.Write(obj_sb.ToString)
          Case "folders"
            Dim obj_sb As New System.Text.StringBuilder()
            Dim str_node As String = GetForm("node", "/")
            obj_sb.Append("[")
            For Each str_folder As String In IO.Directory.GetDirectories(MapPath(str_node))
              str_name = IO.Path.GetFileName(str_folder)
              If Not (str_node = "/" And str_name.ToLower = "cms") Then
                obj_sb.AppendLine("{")
                obj_sb.Append("""id"":""" & JSEncode(UnMappath(str_folder)) & """,")
                obj_sb.Append("""text"":""" & JSEncode(str_name) & """,")
                obj_sb.AppendLine("""cls"":""folder""")
                obj_sb.Append("},")
              End If
            Next
            obj_sb.Remove(obj_sb.Length - 1, 1)
            obj_sb.Append("]")
            Response.Write(obj_sb.ToString)
        End Select
        Response.End()
      ElseIf Security.IsAdmin Then
        Dim obj_tp As New HTMLTemplate("/cms/templates/select.htm")
        obj_tp.SetItem("type", GetQuery("type"))
        obj_tp.Render(Response.OutputStream)
      End If
    End Sub

    Private Function ToFileSize(ByVal vlng_length As Long) As String
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

  End Class

End Namespace