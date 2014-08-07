Public Module _Web

#Region " MapPath "

  Public Function MapPath() As String
    Return MapPath("/")
  End Function

  Public Function MapPath(ByVal vstr_path As String, ByVal vstr_file As String) As String
    Return MapPath(Path.Combine(vstr_path, vstr_file))
  End Function

  Public Function MapPath(ByVal vstr_path As String) As String
    On Error Resume Next
    If Not vstr_path.IndexOf(":\") > 0 And Not vstr_path.IndexOf("\\") >= 0 Then
      If HttpContext.Current IsNot Nothing Then
        vstr_path = HttpContext.Current.Server.MapPath(vstr_path)
      Else
        vstr_path = Hosting.HostingEnvironment.MapPath(vstr_path)
        'vstr_path = HttpRuntime.AppDomainAppPath.TrimEnd("\") & "\" & vstr_path.Replace("~", String.Empty).Replace("/", "\").TrimStart("\")
      End If
    End If
    Return vstr_path.ToLower
  End Function

  Public Function UnMappath(ByVal vstr_path As String) As String
    On Error Resume Next
    Return "/" & vstr_path.ToLower.Substring(HttpRuntime.AppDomainAppPath.ToLower.TrimEnd("/").Length).Replace("\", "/").TrimStart("/")
  End Function

#End Region

#Region " StripHTML "

  Private mobj_opts As RegexOptions = RegexOptions.IgnoreCase Or RegexOptions.Multiline Or RegexOptions.Compiled
  Private mreg_head As New Regex("(?is)<head>.*?<\/head>", mobj_opts)
  Private mreg_script As New Regex("(?is)<script.*?>.*?<\/script>", mobj_opts)
  Private mreg_style As New Regex("(?is)<style.*?>.*?<\/style>", mobj_opts)
  Private mreg_comm As New Regex("(?is)<!--.*?-->", mobj_opts)
  Private mreg_html As New Regex("(?is)<(.|\n)*?>", mobj_opts)
  Private mreg_br As New Regex("(?is)<br([^>]*)>", mobj_opts)

  Public Function StripHtml(ByVal vstr_html As String, Optional ByVal vb_full As Boolean = False) As String
    On Error Resume Next
    vstr_html = mreg_head.Replace(vstr_html, String.Empty)
    vstr_html = mreg_script.Replace(vstr_html, String.Empty)
    vstr_html = mreg_style.Replace(vstr_html, String.Empty)
    vstr_html = mreg_comm.Replace(vstr_html, String.Empty)
    vstr_html = mreg_br.Replace(vstr_html, vbNewLine)
    vstr_html = mreg_html.Replace(vstr_html, String.Empty)
    If vb_full Then
      vstr_html = Regex.Replace(vstr_html, "\t", " ", mobj_opts)
      vstr_html = Regex.Replace(vstr_html, "&nbsp;", " ", mobj_opts)
      vstr_html = Regex.Replace(vstr_html, "&bull;", " * ", mobj_opts)
      vstr_html = Regex.Replace(vstr_html, "&lsaquo;", "<", mobj_opts)
      vstr_html = Regex.Replace(vstr_html, "&rsaquo;", ">", mobj_opts)
      vstr_html = Regex.Replace(vstr_html, "&trade;", "(tm)", mobj_opts)
      vstr_html = Regex.Replace(vstr_html, "&frasl;", "/", mobj_opts)
      vstr_html = Regex.Replace(vstr_html, "&lt;", "<", mobj_opts)
      vstr_html = Regex.Replace(vstr_html, "&gt;", ">", mobj_opts)
      vstr_html = Regex.Replace(vstr_html, "&copy;", "(c)", mobj_opts)
      vstr_html = Regex.Replace(vstr_html, "&reg;", "(r)", mobj_opts)
      vstr_html = Regex.Replace(vstr_html, "&(.{2,6});", String.Empty, mobj_opts)
      vstr_html = Regex.Replace(vstr_html, "\n+", vbNewLine, mobj_opts)
      vstr_html = Regex.Replace(vstr_html, "\s+", " ", mobj_opts)
    End If
    Return vstr_html
  End Function

#End Region

#Region " GetUrlName "

  Public Function GetUrlName(ByVal vstr_value As String) As String
    On Error Resume Next
    vstr_value = ToStr(vstr_value)
    vstr_value = Normalize(vstr_value)
    vstr_value = vstr_value.Replace("&", " and ")
    vstr_value = vstr_value.Replace("\", " and ")
    vstr_value = vstr_value.Replace("/", " and ")
    vstr_value = Regex.Replace(vstr_value, "\s+", " ", RegexOptions.Compiled).Trim
    vstr_value = Regex.Replace(vstr_value.ToLower, "[^a-z0-9\s\-]", "", RegexOptions.Compiled)
    vstr_value = Regex.Replace(vstr_value, "\s", "-", RegexOptions.Compiled)
    vstr_value = Regex.Replace(vstr_value, "\-+", "-", RegexOptions.Compiled)
    Return Server.UrlEncode(vstr_value)
  End Function

#End Region

#Region " SendEmail "

  Public Function SendEmail(ByVal vobj_mail As Net.Mail.MailMessage, Optional ByVal vb_throw As Boolean = False) As Boolean
    Try
      Dim obj_smtp As New SmtpClient(Config("email.host"), ToInt(Config("email.port"), 25))
      If (Config("email.username") & Config("email.password")) <> "" Then
        obj_smtp.Credentials = New Net.NetworkCredential(Config("email.username"), Config("email.password"))
      End If
      obj_smtp.Send(vobj_mail)
      Return True
    Catch ex As Exception
      If vb_throw Then Throw ex
      Return False
    End Try
  End Function

#End Region

#Region " JSEncode "

  Public Function JSEncode(ByVal vstr_value As String) As String
    On Error Resume Next
    If String.IsNullOrEmpty(vstr_value) Then Return String.Empty
    Dim sb As New StringBuilder()
    vstr_value = vstr_value.Trim
    'vstr_value = HTMLEncode(vstr_value)
    Dim t As String
    For Each chr As Char In vstr_value
      Select Case chr
        Case "\"c, """"c, ">"c, "'"c, "&"c
          sb.Append("\"c)
          sb.Append(chr)
        Case ControlChars.Back
          sb.Append("\b")
        Case ControlChars.Tab
          sb.Append("\t")
        Case ControlChars.Lf
          sb.Append("\n")
        Case ControlChars.FormFeed
          sb.Append("\f")
        Case ControlChars.Cr
          sb.Append("\r")
        Case Else
          If chr < " "c Then
            Dim tmp As New String(chr, 1)
            t = "000" & Integer.Parse(tmp, System.Globalization.NumberStyles.HexNumber)
            sb.Append("\u" & t.Substring(t.Length - 4))
          Else
            sb.Append(chr)
          End If
      End Select
    Next
    Return sb.ToString()
  End Function

#End Region

#Region " HTMLEncode "

  Public Function ConvertUnicode(ByVal vstr_value As String) As String
    On Error Resume Next
    Dim int_code As Integer
    Dim str_rep As String
    For i As Integer = vstr_value.Length To 1 Step -1
      int_code = Convert.ToInt32(vstr_value(i - 1))
      If int_code > 127 Then
        str_rep = "&#" + System.Convert.ToString(int_code) + ";"
      Else
        str_rep = vstr_value(i - 1)
      End If
      vstr_value = vstr_value.Substring(0, i - 1) & str_rep & vstr_value.Substring(i)
    Next
    Return vstr_value
  End Function

  Public Function HTMLEncode(ByVal vstr_value As String) As String
    On Error Resume Next
    Return Server.HtmlEncode(ConvertUnicode(vstr_value))
  End Function

  Public Function HTMLDecode(ByVal vstr_html As String) As String
    On Error Resume Next
    Return Server.HtmlDecode(vstr_html)
  End Function

#End Region

#Region " GetMIME "

  Public Function GetMIME(ByVal vstr_file As String) As String
    Select Case IO.Path.GetExtension(vstr_file)
      Case ".xls" : Return "application/vnd.ms-excel"
      Case ".xml" : Return "text/xml"
      Case ".csv" : Return "text/csv"
      Case ".html", ".htm" : Return "text/html"
      Case ".css" : Return "text/css"
      Case ".js" : Return "text/javascript"
      Case ".txt" : Return "text/plain"
      Case ".pdf" : Return "application/pdf"
      Case ".doc" : Return "application/msword"
      Case ".ppt", ".pps" : Return "application/mspowerpoint"
      Case ".gif" : Return "image/gif"
      Case ".png" : Return "image/png"
      Case ".bmp" : Return "image/bmp"
      Case ".jpg", ".jpeg" : Return "image/jpeg"
      Case ".mp3" : Return "audio/mp3"
      Case Else : Return "application/octet-stream"
    End Select
  End Function

#End Region

#Region " GetUrlArgs "

  Public Function GetUrlArgs() As ArgList
    Return GetUrlArgs(Request.RawUrl)
  End Function

  Public Function GetUrlArgs(ByVal vstr_url As String) As ArgList
    On Error Resume Next
    Dim obj_args As New ArgList()
    vstr_url = vstr_url.ToLower
    Dim int_index As Integer = vstr_url.LastIndexOf("?")
    If int_index = -1 Then int_index = vstr_url.Length
    vstr_url = vstr_url.Substring(0, int_index)
    Dim str_ext As String = IO.Path.GetExtension(vstr_url).TrimStart("."c)
    vstr_url = vstr_url.Replace("." & str_ext, String.Empty).Trim("/"c)
    If Not String.IsNullOrEmpty(vstr_url) Then
      obj_args.AddRange(vstr_url.Split("/"c))
    End If
    Return obj_args
  End Function

#End Region

  Public Function JSONEncode(ByVal sIn As String) As String
    On Error Resume Next
    Dim sbOut As New StringBuilder("'")
    For Each ch As Char In sIn
      If Char.IsControl(ch) OrElse ch = "'"c Then
        sbOut.AppendFormat("\u" + Asc(ch).ToString("x4"))
      Else
        If ch = """"c OrElse ch = "\"c OrElse ch = "/"c Then sbOut.Append("\"c)
        sbOut.Append(ch)
      End If
    Next
    sbOut.Append("'")
    Return sbOut.ToString
  End Function

End Module
