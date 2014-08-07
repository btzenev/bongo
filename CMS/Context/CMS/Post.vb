Imports System.Drawing
Imports System.Drawing.Imaging

Namespace Context.CMS

  Public Class Post
    Inherits System.Web.UI.Page

    'Using any content based form you can save, upload and email data
    '  legend: name: description (value: default)

    'web.config - <add key="post.test.email" value="1"/>\
    'post.[name]:			process post. (1|0: 0)
    '  redirect:								redirects to provided page if no errors have occurrer. (string: site root)

    '  spam:							process spam check. (1|0: 0)
    '  spam.field:				field to perform spam check on. (string: "school")
    '  spam.msg:					field to perform spam check on. (string: none)


    '  to:								recipient of email. (string: form.email|cms.email.address)
    '  from:							send of email. (string: cms.email.address)
    '  cc:								copy of email. (string: none)
    '  bcc:								blind copy of email. (string: none)
    '  subject:						email subject. (string: none)
    '  template:					custom html template. (string: intergrated html table)
    '  submitted:					Include submitted date and time. (1|0: 0)
    '  client:						Send a seperate email to client submitting form. (1|0: 0)
    '  client.field:			form field containing destination email. (string: html form field)
    '  client.template:		custom html template. (string: intergrated html table)
    '  save:							save info in database. (1|0: 1)
    '  client.save:				save seperate client info in database. (1|0: 0)
    '  error:							stop process due to error. (1|0: 0)


    'html - using any html form you can save, upload and email data
    '  <form method="post" action="/post.aspx">
    '    <input type="hidden" name="post" value="test" />
    '  </form>

    Private mstr_name As String
    Private mobj_form As New ArrayList

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

      If GetQuery("clear", 0) = 1 Then
        Data.Current.Execute("delete from [cms_post]")
        Response.Redirect("/", True)
        Return
      End If

      If GetQuery("export", 0) = 1 Then
        Dim obj_table As DataTable = Data.Current.Table("select * from [cms_post] order by [created] desc")
        Response.Write("<table border=1><tr>")
        For Each obj_col As DataColumn In obj_table.Columns
          Response.Write("<td>" & obj_col.ColumnName.ToUpper & "</td>")
        Next
        Response.Write("</tr>")
        For Each obj_row As DataRow In obj_table.Rows
          Response.Write("<tr>")
          For Each obj_col As DataColumn In obj_table.Columns
            Response.Write("<td>" & obj_row(obj_col) & "</td>")
          Next
          Response.Write("</tr>")
        Next
        Response.Write("</table>")
        Return
      End If

      Dim b_save As Boolean = (GetSetting("save", 1) = 1) Or (GetSetting("client.save", 0) = 1)
      mstr_name = GetForm("post")
      Dim str_url As String = GetSetting("redirect", "/")
      Dim b_redirect As Boolean = True
      Try
        If b_save Then Data.Current.Open()
        If Request.Form.Count > 0 Then
          mstr_name = GetForm("post")
          If NullSafe(Config("post." & mstr_name), 0) = 1 Then
            Dim b_ok As Boolean = True
            Dim str_spam As String = GetSetting("spam.field", "school")
            If GetSetting("spam", 0) = 1 Then b_ok = (GetForm(str_spam).Trim() = "")
            If b_ok Then
              mobj_form = GetFormList(str_spam & ";X;Y;post;")
              If Not Process_Email() Then
                Throw New System.Exception("Unable to send email")
              End If
            End If
          End If
        End If
      Catch ex As Exception
        If GetSetting("error", 0) = 1 Then
          Response.Write(Server.HtmlEncode(ex.ToString))
          b_redirect = False
        End If
        'If Request.UrlReferrer IsNot Nothing Then str_url = NullSafe(Request.UrlReferrer.ToString, "/")
      Finally
        If b_save Then Data.Current.Close()
        If b_redirect Then Response.Redirect(str_url, True)
      End Try
    End Sub

    Private Function Save(ByVal vobj_mail As Net.Mail.MailMessage) As Boolean
      Dim obj_query As New QueryBuilder("cms_post", QueryType.Insert)
      obj_query.Add("post", mstr_name)
      obj_query.Add("to", vobj_mail.To.ToString)
      obj_query.Add("from", vobj_mail.From.Address)
      obj_query.Add("subject", vobj_mail.Subject)
      obj_query.Add("body", vobj_mail.Body)
      If Data.Current.Execute(obj_query) Then
        Return True
      Else
        Throw New System.Exception("Unable to save data")
        Return False
      End If
    End Function

    Private Function Process_Email() As Boolean
      Dim obj_mail As New Net.Mail.MailMessage
      With obj_mail
        .To.Add(GetSetting("to", NullSafe(GetForm("email"), Config("email.address"))))
        .From = New Net.Mail.MailAddress(GetSetting("from", Config("email.address")))
        Dim str_reply As String = NullSafe(GetForm(GetSetting("replyto", "email")))
        If Not String.IsNullOrEmpty(GetSetting("cc")) Then .ReplyToList.Add(str_reply)
        .Subject = GetSetting("subject")
        If Not String.IsNullOrEmpty(GetSetting("cc")) Then .CC.Add(GetSetting("cc"))
        If Not String.IsNullOrEmpty(GetSetting("bcc")) Then .Bcc.Add(GetSetting("bcc"))
        .IsBodyHtml = True

        Dim b_sub As Boolean = GetSetting("submitted", 0) = 1
        Dim str_sub As String = Date.Now.ToString
        Dim str_file As String = GetSetting("template")
        If str_file <> "" Then
          Dim obj_tp As New HTMLTemplate(str_file)
          For Each str_key As String In mobj_form
            obj_tp.CurrentPart = "fields"
            obj_tp.AddNew()
            obj_tp.SetItem("field", str_key)
            obj_tp.SetItem("label", StrConv(str_key, VbStrConv.ProperCase))
            obj_tp.SetItem("value", GetForm(str_key))
            obj_tp.CurrentPart = ""
            obj_tp.SetItem(str_key & "_label", StrConv(str_key, VbStrConv.ProperCase))
            obj_tp.SetItem(str_key & "_value", GetForm(str_key))
            obj_tp.SetItem(str_key, GetForm(str_key))
          Next
          If b_sub Then
            obj_tp.SetItem("submitted_label", "Submitted")
            obj_tp.SetItem("submitted_value", str_sub)
            obj_tp.SetItem("submitted", str_sub)
          End If
          .Body = obj_tp.Render
        Else
          Dim obj_sb As New StringBuilder
          obj_sb.Append("<html><body>")
          obj_sb.Append("<table border=""0"" cellspacing=""2"" cellpadding=""2"">" & vbNewLine)
          For Each str_key As String In mobj_form
            obj_sb.Append("<tr>" & vbNewLine)
            obj_sb.Append("<th align=""right"">" & StrConv(str_key, VbStrConv.ProperCase) & "</th>" & vbNewLine)
            obj_sb.Append("<td>" & GetForm(str_key) & "</td>" & vbNewLine)
            obj_sb.Append("</tr>" & vbNewLine)
          Next
          If b_sub Then
            obj_sb.Append("<tr>" & vbNewLine)
            obj_sb.Append("<th align=""right"">Submitted</th>" & vbNewLine)
            obj_sb.Append("<td>" & str_sub & "</td>" & vbNewLine)
            obj_sb.Append("</tr>" & vbNewLine)
          End If
          obj_sb.Append("</table>" & vbNewLine)
          obj_sb.Append("</body></html>")
          .Body = obj_sb.ToString
        End If
      End With

      If GetSetting("save", 1) = 1 Then Save(obj_mail)

      If SendEmail(obj_mail) Then
        If GetSetting("client", 0) = 1 Then
          Dim obj_client As New Net.Mail.MailMessage
          With obj_client
            .To.Add(GetForm(GetSetting("client.field", "email")))
            .From = New Net.Mail.MailAddress(GetSetting("from", Config("email.address")))
            .Subject = GetSetting("client.subject", GetSetting("subject"))
            .IsBodyHtml = True
          End With
          Dim str_file As String = GetSetting("client.template")
          If str_file <> "" Then
            Dim obj_tp As New HTMLTemplate(str_file)
            For Each str_key As String In mobj_form
              obj_tp.CurrentPart = "fields"
              obj_tp.AddNew()
              obj_tp.SetItem("field", str_key)
              obj_tp.SetItem("label", StrConv(str_key, VbStrConv.ProperCase))
              obj_tp.SetItem("value", GetForm(str_key))
              obj_tp.CurrentPart = ""
              obj_tp.SetItem(str_key & "_label", StrConv(str_key, VbStrConv.ProperCase))
              obj_tp.SetItem(str_key & "_value", GetForm(str_key))
              obj_tp.SetItem(str_key, GetForm(str_key))
            Next
            obj_client.Body = obj_tp.Render
            If GetSetting("client.save", 0) = 1 Then Save(obj_client)
            SendEmail(obj_client)
          End If
        End If
        Return True
      End If
      Return False
    End Function

    Private Function GetSetting(ByVal vstr_name As String) As String
      Return GetSetting(vstr_name, String.Empty)
    End Function

    Private Function GetSetting(ByVal vstr_name As String, ByVal vobj_default As Object) As String
      Return NullSafe(Config("post." & mstr_name & "." & vstr_name), vobj_default)
    End Function

    Private Function GetSetting(ByVal vstr_name As String, ByVal vobj_default As String) As String
      Return NullSafe(Config("post." & mstr_name & "." & vstr_name), vobj_default)
    End Function

    Private Function GetSetting(ByVal vstr_name As String, ByVal vobj_default As Integer) As Integer
      Return NullSafe(Config("post." & mstr_name & "." & vstr_name), vobj_default)
    End Function

  End Class

End Namespace
