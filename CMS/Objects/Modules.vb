Public Class Modules
  Public ID As Integer
  Public PageID As Integer
  Public Position As String
  Public Order As Integer
  Public Read As String = ""

  Private mstr_plugin As String
  Private mb_plugin As Boolean = False
  Private mobj_plugin As Client.Interface

  Public ReadOnly Property Plugin() As Client.Interface
    Get
      On Error Resume Next
      If mobj_plugin Is Nothing Then
        If Not String.IsNullOrEmpty(mstr_plugin) Then
          mobj_plugin = Plugins.Client.Get(mstr_plugin)
          mobj_plugin.ModuleID = ID
                    mobj_plugin.Page = Pages.Get(PageID)
        End If
      End If
      mb_plugin = mobj_plugin IsNot Nothing
      Return mobj_plugin
    End Get
  End Property

#Region " New "

  Public Sub New(ByVal vstr_plugin As String, ByVal vint_page As Integer)
    MyBase.New()
    mstr_plugin = vstr_plugin
    PageID = vint_page
  End Sub

  Public Sub New(ByVal vint_id As Integer)
    MyBase.New()
    On Error Resume Next
    If vint_id > 0 Then
      ID = vint_id
      Dim obj_row As DataRow = Data.Current.Row("SELECT * FROM [cms_modules] WHERE [id]=" & vint_id & " ORDER BY [order]")
      If obj_row IsNot Nothing Then
        PageID = ToInt(obj_row("pageid"), 0)
        mstr_plugin = "" & obj_row("plugin")
        Position = ToStr(obj_row("position"), "content")
        Order = ToInt(obj_row("order"), 0)
        Read = "" & obj_row("read")
      End If
    End If
  End Sub

  Public Sub New(ByVal vobj_row As DataRow)
    MyBase.New()
    On Error Resume Next
    If vobj_row IsNot Nothing Then
      ID = ToInt(vobj_row("id"), 0)
      PageID = ToInt(vobj_row("pageid"), 0)
      mstr_plugin = "" & vobj_row("plugin")
      Position = ToStr(vobj_row("position"), "content")
      Order = ToInt(vobj_row("order"), 0)
      Read = "" & vobj_row("read")
    End If
  End Sub

#End Region

  Public Function Render() As String
    Dim obj_buf As New Text.StringBuilder
    If Security.HasPermissions(Read) Then
      obj_buf.Append("<div id=""mod_" & ID & """")
      If Security.IsAdmin() Then
        obj_buf.Append(" title=""")
        If Plugin IsNot Nothing Then
          obj_buf.Append(Plugin.Name)
        Else
          obj_buf.Append(mstr_plugin)
        End If
        obj_buf.Append("""")
      End If
      obj_buf.Append(" class=""cms_module"">" & vbNewLine)
      Try
        If Plugin IsNot Nothing Then
          obj_buf.Append(Plugin.Render & vbNewLine)
        Else
          obj_buf.Append("Plugin [" & mstr_plugin & "] is not available.")
        End If
      Catch ex As Exception
        obj_buf.Append("Plugin [" & mstr_plugin & "] is not available.<!--" & Server.HtmlEncode(ex.ToString) & "-->")
      End Try
      obj_buf.Append("</div>" & vbNewLine)
    End If
    Return obj_buf.ToString
  End Function

  Public Function Save() As Boolean
    Dim b_edit As Boolean = (ID > 0)

    If Not b_edit Then
      Dim obj_row As DataRow = Data.Current.Row("select count(id),max(order) from [cms_modules] where [pageid]=" & PageID & " and [position]='" & Position & "'")
      If obj_row IsNot Nothing Then
        If ToInt(obj_row(0), 0) > 0 Then Order = ToInt(obj_row(1), 0)
      End If

      If Pages.Contains(PageID) Then
        Dim obj_page As Page = Pages.Get(PageID)
        Read = obj_page.Read
      End If
    End If

    Read = Read.Replace(",", ";").TrimEnd(";") & ";"
    Read = Read.Replace(";;", ";")
    If Read = ";" Then Read = String.Empty

    Dim obj_query As New QueryBuilder("cms_modules", QueryType.Insert)
    If b_edit Then
      obj_query.Where("id", ID)
      obj_query.QueryType = QueryType.Update
    Else
      obj_query.Add("pageid", PageID)
      obj_query.Add("plugin", mstr_plugin)
    End If
    obj_query.Add("position", ToStr(Position, "content"))
    obj_query.Add("order", Order)
    obj_query.Add("read", Read)
    If Not b_edit Then
      If Data.Current.Insert(obj_query, ID) Then
        If Plugin IsNot Nothing Then Plugin.Add()
        Return True
      End If
    Else
      Return Data.Current.Update(obj_query)
    End If
  End Function

  Public Function Delete() As Boolean
    Try
      Data.Current.Delete("cms_modules", "id", ID)
      If Plugin IsNot Nothing Then Plugin.Delete()
      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function

  Public Shared Function ReOrder(ByVal vstr_position As String, ByVal vobj_order As List(Of Integer)) As Boolean
    On Error Resume Next
    If Not IsEmpty(vstr_position) Then
      Dim obj_query As QueryBuilder
      For i As Integer = 0 To vobj_order.Count - 1
        obj_query = New QueryBuilder("cms_modules", QueryType.Update)
        obj_query.Add("order", i)
        obj_query.Add("position", vstr_position)
        obj_query.Where("id", vobj_order.Item(i))
        Data.Current.Execute(obj_query)
      Next
      Return True
    End If
    Return False
  End Function

End Class
