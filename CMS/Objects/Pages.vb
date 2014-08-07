Imports System.Text
Imports System.Text.RegularExpressions

Public Class Page
  Public ID As Integer
  Public Parent As Integer
  Public Name As String
  Public Title As String
  Public Visible As Integer
  Public Landing As Integer
  Public Order As Integer
  Public Template As Integer
  Public Read As String
  Public Redirect As String
  Public Meta_Title As String
  Public Meta_Description As String
  Public Meta_Keywords As String
  Public Custom As New CustomCollection
  Public Application As Integer


  Private mobj_styles As New List(Of String)
  Public ReadOnly Property Styles As List(Of String)
    Get
      mobj_styles = mobj_styles.Distinct.ToList
      Return mobj_styles
    End Get
  End Property

  Private mobj_scripts As New List(Of String)
  Public ReadOnly Property Scripts As List(Of String)
    Get
      mobj_scripts = mobj_scripts.Distinct.ToList
      Return mobj_scripts
    End Get
  End Property

  Private mstr_url As String

  Public Property Url As String
    Get
      If Landing = 1 Then Return "/"
      Return mstr_url
    End Get
    Set(ByVal value As String)
      mstr_url = value
    End Set
  End Property

  Public ReadOnly Property Args As ArgList
    Get
      On Error Resume Next
      Dim obj_temp As New ArgList(Global.RawUrl)
      Dim obj_args As New ArgList(MyClass.Url)
      obj_temp.RemoveRange(0, obj_args.Count)
      Return obj_temp
    End Get
  End Property

  Public Class CustomCollection
    Inherits Dictionary(Of String, CustomItem)

    Public Class CustomItem
      Public Value As Object
      Public Type As System.Type
      Public Sub New(ByVal vobj_value As Object, ByVal vobj_type As System.Type)
        Value = vobj_value
        Type = vobj_type
      End Sub
    End Class

    Default Public Overloads Property Item(ByVal vstr_key As String) As Object
      Get
        If MyBase.ContainsKey(vstr_key) Then Return MyBase.Item(vstr_key).Value
      End Get
      Set(ByVal value As Object)
        If MyBase.ContainsKey(vstr_key) Then MyBase.Item(vstr_key).Value = value
      End Set
    End Property

    Public Overloads ReadOnly Property DataType(ByVal vstr_key As String) As System.Type
      Get
        If MyBase.ContainsKey(vstr_key) Then Return MyBase.Item(vstr_key).Type
      End Get
    End Property

    Public Overloads Sub Add(ByVal vstr_key As String, ByVal vobj_value As Object, ByVal vobj_type As System.Type)
      MyBase.Add(vstr_key, New CustomItem(vobj_value, vobj_type))
    End Sub

  End Class

#Region " New "

  Public Sub New()
    MyBase.New()
  End Sub

  Public Sub New(ByVal vobj_row As DataRow)
    MyBase.New()
    On Error Resume Next

    ID = ToInt(vobj_row("id"), 0)
    Parent = ToInt(vobj_row("parent"), 0)
    Name = ToStr(vobj_row("name")).Trim.ToLower
    Title = ToStr(vobj_row("title"))
    Url = "/" & Name
    Visible = ToInt(vobj_row("visible"), 0)
    Landing = ToInt(vobj_row("landing"), 0)
    Order = ToInt(vobj_row("order"), 0)
    Template = ToInt(vobj_row("template"), 0)
    Redirect = ToStr(vobj_row("redirect"))
    Read = ToStr(vobj_row("read"))
    Application = ToStr(vobj_row("app"))

    Meta_Title = ToStr(vobj_row("meta_title"))
    Meta_Description = ToStr(vobj_row("meta_description"))
    Meta_Keywords = ToStr(vobj_row("meta_keywords"))

    Dim str_col As String
    For Each obj_col As DataColumn In vobj_row.Table.Columns
      str_col = obj_col.ColumnName.ToLower
      If str_col.StartsWith("custom_") Then
        Custom.Add(str_col.Substring(7), vobj_row(obj_col), obj_col.DataType)
      End If
    Next
  End Sub

#End Region

  Public Function Delete() As Boolean
    Return Pages.Delete(Me.ID)
  End Function

  Public Function Save() As Boolean
    Dim b_edit As Boolean = (ID > 0)
    Title = NullSafe(Title).Trim
    Name = GetUrlName(Name)
    For Each obj_page As Page In Pages.Children(Parent)
      If obj_page.Name = Name And Not obj_page.ID = ID Then
        For i As Integer = 2 To 1000
          If Not obj_page.Name = Name & "_" & i Then
            Name &= "_" & i
            Exit For
          End If
        Next
        Exit For
      End If
    Next

    Url = "/" & Name
    If Parent > 0 Then Url = Pages.Get(Parent).Url & Url
    Url = Url.ToLower

    Dim obj_query As New QueryBuilder("cms_pages", QueryType.Insert)
    If b_edit Then
      obj_query.QueryType = QueryType.Update
      obj_query.Where("id", ID)
    End If

    obj_query.Add("parent", Parent)
    obj_query.Add("title", Title)
    obj_query.Add("name", Name)
    obj_query.Add("template", Template)
    obj_query.Add("redirect", Redirect)
    obj_query.Add("order", Order)
    obj_query.Add("visible", Visible)
    obj_query.Add("landing", Landing)
    obj_query.Add("app", Application)

    obj_query.Add("meta_title", ToStr(Meta_Title))
    obj_query.Add("meta_description", ToStr(Meta_Description))
    obj_query.Add("meta_keywords", ToStr(Meta_Keywords))

    Read = Read.Replace(",", ";").TrimEnd(";") & ";"
    obj_query.Add("read", Read.Replace(";;", ";"))
    If Not b_edit Then obj_query.Add("order", Pages.Children(Parent).Count + 1)

    For Each str_col As String In Custom.Keys
      obj_query.Add("custom_" & str_col, Custom(str_col))
    Next

    If Landing = 1 Then Data.Current.ExecuteNonQuery("update [cms_pages] set [landing]=0")
    '  Log(obj_query.ToString)

    Dim b_saved As Boolean = False
    If Not b_edit Then
      b_saved = Data.Current.Insert(obj_query, ID)
    Else
      Dim int_oldOrder As Integer = Data.Current.Get("cms_pages", "order", "id", ID)
      b_saved = Data.Current.Update(obj_query)
      If Not int_oldOrder = Order Then Pages.ReOrder(Order, Parent, ID)
    End If

    If b_saved Then
      'If Landing = 1 Then Config("page.landing") = ID
      If Pages.Load() Then
        Return True
      End If
    End If
    Return False
  End Function

#Region " Render "

  Public Function Render() As String
    Dim obj_sb As New Text.StringBuilder
    Using obj_writer As New IO.StringWriter(obj_sb)
      Render(obj_writer)
    End Using
    Return obj_sb.ToString
  End Function

  Public Sub Render(ByVal vobj_stream As System.IO.Stream)
    Using obj_writer As New IO.StreamWriter(vobj_stream)
      Render(obj_writer)
    End Using
  End Sub

  Private Shared mreg_tags As New Regex("\[\[(.*?)\]\]", RegexOptions.Singleline Or RegexOptions.IgnoreCase Or RegexOptions.Compiled)

  Public Sub Render(ByVal vobj_writer As System.IO.TextWriter)
    On Error Resume Next
    Dim str_data As String
    Dim obj_tp As HTMLTemplate
    Dim obj_mod As Modules

    mobj_styles.Clear()
    mobj_scripts.Clear()

    If Not Security.IsAdmin() Then
      If Not String.IsNullOrEmpty(Redirect) Then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", Redirect)
        Response.End()
        Return
      End If
    End If

    Dim obj_temp As DataRow = Data.Current.Row("select * from [cms_templates] where id=" & Template)
    If obj_temp Is Nothing Then
      Throw New System.Exception("No template available. ")
    End If
    obj_tp = New HTMLTemplate(NullSafe(obj_temp("src")))
    If obj_tp Is Nothing Then
      Throw New System.Exception("No template available. ")
    End If
    Dim str_tppos As String = "" & obj_temp("position")


    Dim obj_hidden As Hidden.Interface
    For Each str_plugin As String In Plugins.Hidden.List
      If obj_tp.GlobalTagExists(str_plugin) Then
        obj_hidden = Plugins.Hidden.Get(str_plugin)
        If obj_hidden IsNot Nothing Then
          obj_hidden.Page = Me
          str_data = obj_hidden.Render
          If Not String.IsNullOrEmpty(str_data) Then
            str_data = mreg_tags.Replace(str_data, "{{$1}}")
            obj_tp.SetGlobalItem(str_plugin, str_data)
          End If
        End If
      End If
    Next

    Dim obj_client As Client.Interface
    For Each str_plugin As String In Plugins.Client.List
      If obj_tp.GlobalTagExists(str_plugin) Then
        obj_client = Plugins.Client.Get(str_plugin)
        If obj_client IsNot Nothing Then
          obj_client.Page = Me
          obj_client.ModuleID = -1
          str_data = obj_client.Render
          If Not String.IsNullOrEmpty(str_data) Then
            str_data = mreg_tags.Replace(str_data, "{{$1}}")
            obj_tp.SetGlobalItem(str_plugin, str_data)
          End If
        End If
      End If
    Next

    Dim obj_table As DataTable = Data.Current.Table("SELECT * FROM [cms_modules] WHERE [pageid]=" & ID & " ORDER BY [position], [order] ASC")
    For Each obj_row As DataRow In obj_table.Rows
      obj_mod = New Modules(obj_row)
      If obj_tp.Sections.ContainsKey(obj_mod.Position) Then
        obj_tp.Sections(obj_mod.Position).AddNew()
        str_data = obj_mod.Render()
        If Not String.IsNullOrEmpty(str_data) Then
          str_data = mreg_tags.Replace(str_data, "{{$1}}")
          obj_tp.Sections(obj_mod.Position).SetItem("data", str_data)
        End If
      End If
    Next

    obj_tp.CurrentPart = ""

    For Each str_col As String In Custom.Keys
      obj_tp.SetGlobalItem("page.custom_" & str_col, "" & Custom(str_col))
    Next

    obj_tp.SetGlobalItem("page.title", Title)
    obj_tp.SetGlobalItem("page.name", Name)

    obj_tp.SetGlobalItem("page.meta_title", NullSafe(Meta_Title, Title))
    obj_tp.SetGlobalItem("page.meta_description", Meta_Description)
    obj_tp.SetGlobalItem("page.meta_keywords", Meta_Keywords)

    obj_tp.SetGlobalItem("page.url", Url)
    obj_tp.SetGlobalItem("page.id", ID)
    obj_tp.SetGlobalItem("crumbs", BreadCrumbs(ID))
    Dim int_root As Integer = GetRoot(ID)
    obj_tp.SetGlobalItem("page.root", int_root)
    Dim obj_page As Page = Pages.Get(int_root)
    obj_tp.SetGlobalItem("root.title", obj_page.Title)
    obj_tp.SetGlobalItem("root.name", obj_page.Name)
    obj_tp.SetGlobalItem("root.url", obj_page.Url)
    obj_tp.SetGlobalItem("root.id", int_root)

    Dim int_subroot As Integer = GetRoot(ID, True)
    If int_subroot <> int_root Then
      obj_tp.SetGlobalItem("page.subroot", int_subroot)
      Dim obj_subpage As Page = Pages.Get(int_subroot)
      obj_tp.SetGlobalItem("subroot.title", obj_subpage.Title)
      obj_tp.SetGlobalItem("subroot.name", obj_subpage.Name)
      obj_tp.SetGlobalItem("subroot.url", obj_subpage.Url)
      obj_tp.SetGlobalItem("subroot.id", int_subroot)
    End If

    Dim int_userid As Integer = Security.UserID
    obj_tp.SetGlobalItem("userid", int_userid)
    obj_tp.SetGlobalItem("isauthenticated", IIf(Request.IsAuthenticated, 1, 0))

    For Each str_script As String In MyClass.Scripts
      obj_tp("page.scripts").AddItem("src", str_script)
    Next
    mobj_scripts.Clear()

    For Each str_style As String In MyClass.Styles
      obj_tp("page.styles").AddItem("src", str_style)
    Next
    mobj_styles.Clear()

    'If Security.IsAdmin Then obj_tp.SetGlobalItem("cms", HTMLTemplate.QuickRender("/CMS/plugins/admin_site/site.htm"))
    obj_tp.Render(vobj_writer)

  End Sub

#End Region

End Class

Public Class Pages

  Private Shared mobj_children As New Dictionary(Of Integer, List(Of Integer))
  Private Shared mobj_urls As New Dictionary(Of String, Integer)
  Private Shared mobj_pages As New Dictionary(Of Integer, Page)

#Region " Properties "

  Public Shared ReadOnly Property List() As Dictionary(Of Integer, Page)
    Get
      Return mobj_pages
    End Get
  End Property

  Public Shared ReadOnly Property [Get](ByVal vint_id As Integer) As Page
    Get
      If mobj_pages.ContainsKey(vint_id) Then Return mobj_pages(vint_id)
    End Get
  End Property

  Public Shared ReadOnly Property [Get](ByVal vstr_url As String) As Page
    Get
      On Error Resume Next
      Return [Get](ID(vstr_url))
    End Get
  End Property

  Public Shared ReadOnly Property ID(ByVal vstr_url As String) As Integer
    Get
      On Error Resume Next
      vstr_url = vstr_url.ToLower.Trim
      If vstr_url.IndexOf(".aspx") > 0 Then vstr_url = vstr_url.Substring(0, vstr_url.Length - 5)
      vstr_url = vstr_url.TrimEnd("/")
      If mobj_urls.ContainsKey(vstr_url) Then Return mobj_urls(vstr_url)
      Return -1
    End Get
  End Property

  Public Shared ReadOnly Property Children(ByVal vint_id As Integer, Optional ByVal vint_skip As Integer = -1) As List(Of Page)
    Get
      On Error Resume Next
      Dim obj_list As New List(Of Page)
      If mobj_children.ContainsKey(vint_id) Then
        For Each int_id As Integer In mobj_children(vint_id)
          If Not int_id = vint_skip And mobj_pages.ContainsKey(int_id) Then
            obj_list.Add(mobj_pages(int_id))
          End If
        Next
        obj_list.Sort(AddressOf Sort)
      End If
      Return obj_list
    End Get
  End Property

  Public Shared ReadOnly Property Urls() As Dictionary(Of String, Integer)
    Get
      Return mobj_urls
    End Get
  End Property

#End Region

#Region " Functions "

  Public Shared Function Contains(ByVal vint_id As Integer) As Boolean
    Return mobj_pages.ContainsKey(vint_id)
  End Function

  Public Shared Function Sort(ByVal x As Page, ByVal y As Page) As Integer
    On Error Resume Next
    Return x.Order < y.Order
  End Function

  Public Shared Function ReOrder(ByVal vint_parent As Integer, ByVal vobj_order As List(Of Integer)) As Boolean
    On Error Resume Next
    Dim obj_query As QueryBuilder
    For i As Integer = 0 To vobj_order.Count - 1
      obj_query = New QueryBuilder("cms_pages", QueryType.Update)
      obj_query.Add("order", i)
      obj_query.Add("parent", vint_parent)
      obj_query.Where("id", vobj_order.Item(i))
      Data.Current.Execute(obj_query)
    Next
    Return Load()
  End Function

  Public Shared Sub ReOrder(ByVal vint_order As Integer, ByVal vint_parent As Integer, ByVal vint_id As Integer)
    Dim i As Integer
    Dim obj_pages As New List(Of Page)
    obj_pages = Pages.Children(vint_parent)
    obj_pages.Remove(Pages.Get(vint_id))
    'Data.Current.Begin()
    If vint_order = -1 Then
      vint_order = obj_pages.Count + 1
    End If
    Data.Current.Execute("UPDATE [cms_pages] SET [order]=" & vint_order & " WHERE [id]=" & vint_id)
    For Each obj_page As Page In obj_pages
      i += 1
      If vint_order = i Then i += 1
      Data.Current.Execute("UPDATE [cms_pages] SET [order]=" & i & " WHERE [id]=" & obj_page.ID)
    Next
    'Data.Current.Commit()
  End Sub

  Public Shared Function Delete(ByVal vint_id As Integer) As Boolean
    On Error Resume Next
    If vint_id > 0 Then
      If Pages.Contains(vint_id) Then
        Dim obj_page As Page = [Get](vint_id)
        'If obj_page.Locked And Not Security.IsGlobal Then Return False
        Dim obj_mod As Modules
        For Each obj_row As DataRow In Data.Current.Rows("select [id] from [cms_modules] where [pageid]=" & vint_id)
          obj_mod = New Modules(obj_row(0), vint_id)
          obj_mod.Delete()
        Next
        Data.Current.Delete("cms_pages", "id", vint_id)
        Pages.Load()
        Return True
      End If
    End If
    Return False
  End Function

#End Region

#Region " Loaders "

  Public Shared Function Load() As Boolean
    Try
      mobj_pages = New Dictionary(Of Integer, Page)
      Dim obj_page As Page
      For Each obj_row As DataRow In Data.Current.Rows("SELECT * FROM [cms_pages] order by [parent]")
        obj_page = New Page(obj_row)
        If obj_page.Landing = 1 Then Config("page.landing") = obj_page.ID
        mobj_pages.Add(obj_page.ID, obj_page)
      Next
      LoadChildren()
      LoadUrls()
      Return True
    Catch ex As Exception
      Log(ex)
      Return False
    End Try
  End Function

  Public Shared Sub LoadChildren()
    On Error Resume Next
    mobj_children = New Dictionary(Of Integer, List(Of Integer))
    mobj_children.Add(0, New List(Of Integer))
    For Each obj_page As Page In mobj_pages.Values
      If mobj_children.ContainsKey(obj_page.Parent) Then
        If Not mobj_children(obj_page.Parent).Contains(obj_page.ID) Then
          mobj_children(obj_page.Parent).Add(obj_page.ID)
        End If
      Else
        mobj_children.Add(obj_page.Parent, New List(Of Integer))
        mobj_children(obj_page.Parent).Add(obj_page.ID)
      End If
    Next
  End Sub

  Private Shared Sub LoadUrls(ByVal vint_parent As Integer)
    Dim obj_parent As Page
    For Each obj_page As Page In Children(vint_parent)
      If obj_page.Parent > 0 Then
        If mobj_pages.ContainsKey(obj_page.Parent) Then
          obj_parent = mobj_pages(obj_page.Parent)
          obj_page.Url = obj_parent.Url & "/" & obj_page.Name
        End If
      End If
      mobj_urls.Add(obj_page.Url, obj_page.ID)
      LoadUrls(obj_page.ID)
    Next
  End Sub

  Public Shared Sub LoadUrls()
    mobj_urls = New Dictionary(Of String, Integer)
    LoadUrls(0)
  End Sub

#End Region

End Class