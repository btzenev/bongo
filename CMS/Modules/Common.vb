Public Module _Common



#Region " Exception "

    Public Sub Exception(ByVal vstr_msg As String)
        Log(vstr_msg)
        Throw New System.Exception(vstr_msg)
    End Sub

    Public Sub Exception(ByVal vstr_context As String, ByVal vstr_msg As String, Optional ByVal vb_throw As Boolean = False)
        vstr_msg = vstr_context.ToUpper & ": " & vstr_msg
        Log(vstr_msg)
        If Debug Or vb_throw Then
            Throw New System.Exception(vstr_msg)
        End If
    End Sub

    Public Sub Exception(ByVal vstr_context As String, ByVal vstr_msg As String, ByVal vobj_ex As Exception, Optional ByVal vb_throw As Boolean = False)
        vstr_msg = vstr_context.ToUpper & ": " & vstr_msg
        Log(vstr_msg, vobj_ex)
        If Debug Or vb_throw Then
            Throw New System.Exception(vstr_msg, vobj_ex)
        End If
    End Sub

    Public Sub Exception(ByVal vstr_context As String, ByVal vobj_ex As Exception, Optional ByVal vb_throw As Boolean = False)
        Log(vobj_ex)
        If Debug Or vb_throw Then
            If vobj_ex IsNot Nothing Then
                Throw vobj_ex
            Else
                Throw New System.Exception(vstr_context.ToUpper & ": An unknown exception occurred.")
            End If
        End If
    End Sub

#End Region


    Public Function HashJoin(ByVal vobj_hash As ICollection, Optional ByVal vstr_delimiter As String = "") As String
        If Not vobj_hash Is Nothing Then
            Dim str_hash(vobj_hash.Count - 1) As String
            vobj_hash.CopyTo(str_hash, 0)
            Return Join(str_hash, vstr_delimiter)
        End If
        Return String.Empty
    End Function

    Public Function HashToArray(ByVal vobj_hash As ICollection) As Object()
        If Not vobj_hash Is Nothing Then
            Dim obj_hash(vobj_hash.Count - 1) As Object
            vobj_hash.CopyTo(obj_hash, 0)
            Return obj_hash
        End If
        Return Nothing
    End Function

    Public Function GetFormList(Optional ByVal vstr_exclude As String = "", Optional ByVal vstr_prefix As String = "") As ArrayList
        Dim obj_form As New ArrayList
        vstr_exclude += "post_redirect;"
        Dim obj_exclude As New ArrayList(vstr_exclude.Split(";"))
        For Each str_item As String In Request.Form.Keys
            If vstr_prefix = "" Then
                If Not obj_exclude.Contains(str_item) Then
                    obj_form.Add(str_item)
                End If
            Else
                If str_item.StartsWith(vstr_prefix & "_") Then
                    If Not obj_exclude.Contains(str_item) Then
                        obj_form.Add(str_item)
                    End If
                End If
            End If
        Next
        Return obj_form
    End Function

    Public Function GetQueryList(Optional ByVal vstr_exclude As String = "", Optional ByVal vstr_prefix As String = "") As ArrayList
        Dim obj_form As New ArrayList
        Dim obj_exclude As New ArrayList(vstr_exclude.Split(";"))
        For Each str_item As String In Request.QueryString.Keys
            If vstr_prefix = "" Then
                If Not obj_exclude.Contains(str_item) Then
                    obj_form.Add(str_item)
                End If
            Else
                If str_item.StartsWith(vstr_prefix & "_") Then
                    If Not obj_exclude.Contains(str_item) Then
                        obj_form.Add(str_item)
                    End If
                End If
            End If
        Next
        Return obj_form
    End Function

#Region " NullSafe "

    'DEFAULT / OBJECT

    Public Function GetCookie(ByVal vstr_name As String, ByVal vobj_default As Object) As Object
        If Not Request.Cookies(vstr_name) Is Nothing Then
            Response.Cookies(vstr_name).Value = NullSafe(Request.Cookies(vstr_name).Value, vobj_default)
            Return Request.Cookies(vstr_name).Value
        Else
            Return vobj_default
        End If
    End Function

    Public Function GetForm(ByVal vstr_name As String, ByVal vobj_default As Object) As Object
        Return NullSafe(Request.Form(vstr_name), vobj_default)
    End Function

    Public Function GetQuery(ByVal vstr_name As String, ByVal vobj_default As Object) As Object
        Return NullSafe(Request.Form(vstr_name), vobj_default)
    End Function

    Public Function GetRequest(ByVal vstr_name As String, ByVal vobj_default As Object) As Object
        Return NullSafe(Request(vstr_name), vobj_default)
    End Function

    Public Function NullSafe(ByVal vobj_value As Object) As String
        Return NullSafe(vobj_value, String.Empty)
    End Function

    Public Function NullSafe(ByVal vobj_value As Object, ByVal vobj_default As Object) As Object
        Try
            If Not vobj_value Is Nothing Then
                If Not vobj_value Is DBNull.Value Then
                    Return vobj_value
                End If
            End If
        Catch
        End Try
        Return vobj_default
    End Function

    'STRING

    Public Function GetCookie(ByVal vstr_name As String) As String
        Return GetCookie(vstr_name, "")
    End Function

    Public Function GetCookie(ByVal vstr_name As String, ByVal vstr_default As String) As String
        If Not Request.Cookies(vstr_name) Is Nothing Then
            Return NullSafe(Request.Cookies(vstr_name).Value, vstr_default)
        Else
            Return vstr_default
        End If
    End Function

    Public Function GetForm(ByVal vstr_name As String) As String
        Return GetForm(vstr_name, "")
    End Function

    Public Function GetForm(ByVal vstr_name As String, ByVal vstr_default As String) As String
        Return NullSafe(Request.Form(vstr_name), vstr_default)
    End Function

    Public Function GetQuery(ByVal vstr_name As String) As String
        Return GetQuery(vstr_name, "")
    End Function

    Public Function GetQuery(ByVal vstr_name As String, ByVal vstr_default As String) As String
        Return NullSafe(Request.QueryString(vstr_name), vstr_default)
    End Function

    Public Function GetRequest(ByVal vstr_name As String) As String
        Return GetRequest(vstr_name, "")
    End Function

    Public Function GetRequest(ByVal vstr_name As String, ByVal vstr_default As String) As String
        Return NullSafe(Request(vstr_name), vstr_default)
    End Function

    Public Function NullSafe(ByVal vobj_value As Object, ByVal vstr_default As String) As String
        Try
            If Not vobj_value Is Nothing Then
                If Not vobj_value Is DBNull.Value Then
                    Dim str_value As String = vobj_value.ToString.Trim
                    If Not str_value = String.Empty Then
                        Return str_value
                    End If
                End If
            End If
        Catch
        End Try
        Return vstr_default
    End Function

    'INTEGER

    Public Function GetCookie(ByVal vstr_name As String, ByVal vint_default As Integer) As Integer
        If Not Request.Cookies(vstr_name) Is Nothing Then
            Return NullSafe(Request.Cookies(vstr_name).Value, vint_default)
        Else
            Return vint_default
        End If
    End Function

    Public Function GetForm(ByVal vstr_name As String, ByVal vint_default As Integer) As Integer
        Return NullSafe(Request.Form(vstr_name), vint_default)
    End Function

    Public Function GetQuery(ByVal vstr_name As String, ByVal vint_default As Integer) As Integer
        Return NullSafe(Request.QueryString(vstr_name), vint_default)
    End Function

    Public Function GetRequest(ByVal vstr_name As String, ByVal vint_default As Integer) As Integer
        Return NullSafe(Request(vstr_name), vint_default)
    End Function

    Public Function NullSafe(ByVal vobj_value As Object, ByVal vint_default As Integer) As Integer
        Try
            If Not vobj_value Is Nothing Then
                If Not vobj_value Is DBNull.Value Then
                    Dim str_value As String = vobj_value.ToString
                    If IsNumeric(str_value) Then
                        Return CInt(str_value)
                    End If
                End If
            End If
        Catch
        End Try
        Return vint_default
    End Function

    'BOOLEAN

    Public Function GetCookie(ByVal vstr_name As String, ByVal vb_default As Boolean) As Boolean
        If Not Request.Cookies(vstr_name) Is Nothing Then
            Return NullSafe(Request.Cookies(vstr_name).Value, vb_default)
        Else
            Return vb_default
        End If
    End Function

    Public Function GetForm(ByVal vstr_name As String, ByVal vb_default As Boolean) As Boolean
        Return NullSafe(Request.Form(vstr_name), vb_default)
    End Function

    Public Function GetQuery(ByVal vstr_name As String, ByVal vb_default As Boolean) As Boolean
        Return NullSafe(Request.QueryString(vstr_name), vb_default)
    End Function

    Public Function GetRequest(ByVal vstr_name As String, ByVal vb_default As Boolean) As Boolean
        Return NullSafe(Request(vstr_name), vb_default)
    End Function

    Public Function NullSafe(ByVal vobj_value As Object, ByVal vb_default As Boolean) As Boolean
        Try
            If Not vobj_value Is Nothing Then
                If Not vobj_value Is DBNull.Value Then
                    Dim str_value As String = vobj_value.ToString
                    If IsNumeric(str_value) Then
                        Return NullSafe(str_value, 0) = 1
                    Else
                        Return Boolean.Parse(str_value)
                    End If
                End If
            End If
        Catch
        End Try
        Return vb_default
    End Function

    'DATE

    Public Function GetCookie(ByVal vstr_name As String, ByVal vdat_default As Date) As Date
        If Not Request.Cookies(vstr_name) Is Nothing Then
            Return NullSafe(Request.Cookies(vstr_name).Value, vdat_default)
        Else
            Return vdat_default
        End If
    End Function

    Public Function GetForm(ByVal vstr_name As String, ByVal vdat_default As Date) As Date
        Return NullSafe(Request.Form(vstr_name), vdat_default)
    End Function

    Public Function GetQuery(ByVal vstr_name As String, ByVal vdat_default As Date) As Date
        Return NullSafe(Request.QueryString(vstr_name), vdat_default)
    End Function

    Public Function GetRequest(ByVal vstr_name As String, ByVal vdat_default As Date) As Date
        Return NullSafe(Request(vstr_name), vdat_default)
    End Function

    Public Function NullSafe(ByVal vobj_value As Object, ByVal vdat_default As Date) As Date
        Try
            If Not vobj_value Is Nothing Then
                If Not vobj_value Is DBNull.Value Then
                    Dim str_value As String = vobj_value.ToString
                    If IsDate(str_value) Then
                        Return Date.Parse(str_value)
                    End If
                End If
            End If
        Catch
        End Try
        Return vdat_default
    End Function

#End Region

#Region " FormatError "

    Public Function FormatError(ByVal vobj_ex As System.Exception) As String
        Return FormatError("Not Available At This Time.", vobj_ex)
    End Function

  Public Function FormatError(ByVal vstr_msg As String, ByVal vobj_ex As System.Exception) As String
    On Error Resume Next
    Dim obj_sb As New StringBuilder(vstr_msg)
    If vobj_ex IsNot Nothing Then obj_sb.Append("<br><!--" & Server.HtmlEncode(vobj_ex.ToString.Replace("--", "__")) & "-->")
    Return obj_sb.ToString
  End Function

#End Region

#Region " GetLang "

  Public Function GetLang() As String
    On Error Resume Next
    Dim int_len As Integer = Config("site.language")
    If Args.Count > 0 Then
      If Args(0).Length = int_len Then
        Return Args(0).ToLower
      End If
    End If
    Return IIf(int_len = 3, "eng", "en")
  End Function

  Public Function GetLang(ByVal vstr_key As String, Optional ByVal vstr_default As String = "") As String
    On Error Resume Next
    If Languages.ContainsKey(vstr_key) Then
      Dim str_lang As String = GetLang()
      If Languages(vstr_key).ContainsKey(str_lang) Then
        Return NullSafe(Languages(vstr_key)(str_lang), vstr_default)
      End If
    End If
    Return vstr_default
  End Function

#End Region

#Region " GetParents "

  Private Sub GetParents(ByRef vobj_list As List(Of Integer), ByVal vint_root As Integer, ByVal vint_id As Integer)
    On Error Resume Next
    If vint_id = 0 Or vint_id = vint_root Then Return
    If Pages.List.ContainsKey(vint_id) Then
      vobj_list.Add(vint_id)
      GetParents(vobj_list, vint_root, Pages.Get(vint_id).Parent)
    End If
  End Sub

  Public Function GetParents(ByVal vint_id As Integer, Optional ByVal vb_sub As Boolean = False) As List(Of Integer)
    On Error Resume Next
    Dim obj_list As New List(Of Integer)
    GetParents(obj_list, GetRoot(vint_id, vb_sub), vint_id)
    Return obj_list
  End Function

#End Region

  Public Function GetRoot(ByVal int_id As Integer, Optional ByVal vb_sub As Boolean = False) As Integer
    On Error Resume Next
    If int_id = 0 Then Return int_id
    Dim obj_page As Page = Pages.Get(int_id)
    If vb_sub Then
      If obj_page.Parent = 0 Then
        Return int_id
      Else
        obj_page = Pages.Get(obj_page.Parent)
        If (MultiLingual And obj_page.Url = "/" & GetLang()) Then
          Return int_id
        Else
          Return GetRoot(obj_page.ID, vb_sub)
        End If
      End If
    Else
      If (MultiLingual And obj_page.Url = "/" & GetLang()) Then
        Return int_id
      Else
        Return GetRoot(obj_page.Parent, vb_sub)
      End If
    End If
  End Function

  Public Function BreadCrumbs(ByVal int_page As Integer) As String
    Dim obj_stack As New System.Collections.Stack
    Dim obj_page As Page
    Dim b_current As Boolean = True
    While int_page > 0
      obj_page = Pages.Get(int_page)
      If b_current Then
        b_current = False
        obj_stack.Push("<span class=""crumb last"">" & obj_page.Title & "</span>")
      Else
        obj_stack.Push("<a class=""crumb"" href=""" & obj_page.Url & """>" & obj_page.Title & "</a>")
      End If
      int_page = obj_page.Parent
    End While
    obj_stack.Push("<a class=""crumb first"" href=""/"">Home</a>")
    Return Join(obj_stack.ToArray, ":")
  End Function

  Public Sub SetCache(ByVal vobj_cache As System.Web.HttpCacheability)
    If vobj_cache = System.Web.HttpCacheability.NoCache Then
      Response.AddHeader("pragma", "no-cache")
      Response.AddHeader("cache-control", "private, no-cache, must-revalidate")
      Response.Expires = -1
      Response.ExpiresAbsolute = DateTime.Now.AddDays(-1)
      Response.CacheControl = "no-cache"
    End If
    Response.Cache.SetCacheability(vobj_cache)
  End Sub

End Module
