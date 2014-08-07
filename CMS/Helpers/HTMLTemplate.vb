Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Globalization

Public Class HTMLTemplate

#Region " HTMLTemplateData "

  Friend Class HTMLTemplateData
    Public HTML As String
    Public Sections As New Dictionary(Of String, HTMLTemplate)
    Public Blocks As New Hashtable

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal vstr_html As String, ByVal vobj_sections As Dictionary(Of String, HTMLTemplate))
      MyBase.New()
      Me.HTML = vstr_html
      Me.Sections = New Dictionary(Of String, HTMLTemplate)(vobj_sections)
    End Sub

    Public Function Clone() As HTMLTemplateData
      Dim obj_tp As HTMLTemplate
      Dim obj_return As New HTMLTemplateData
      obj_return.HTML = String.Empty & HTML
      For Each str_key As String In Sections.Keys
        obj_tp = Sections(str_key)
        obj_return.Sections.Add(str_key, obj_tp.Clone)
      Next
      Return obj_return
    End Function

    Public Function Contains(ByVal vstr_name As String) As Boolean
      If Sections.ContainsKey(vstr_name) Then
        Return True
      Else
        Dim int_start As Integer = vstr_name.IndexOf(".")
        If int_start > 0 Then
          Dim str_name As String = vstr_name.Substring(0, int_start)
          vstr_name = vstr_name.Substring(int_start + 1)
          If Sections.ContainsKey(str_name) Then
            Return Sections(str_name).Contains(vstr_name)
          End If
        End If
      End If
    End Function
  End Class

#End Region

#Region " CacheData "

  Private Class CacheData
    Public CRC As String
    Public TemplateData As HTMLTemplateData
  End Class

#End Region

#Region " Tag Evaluator "

  Friend Class TagEval
    Public Value As String
    Private mobj_tp As HTMLTemplate
    'new: \.(?=([^"]*"[^"]*"[^"]*)*$|[^"]*$)
    'old: (\w*(?:\((?:\".*?\"|.*?)\))?).?|.
    Private Shared mreg_methods As New Regex("\.(?=([^""]*""[^""]*""[^""]*)*$|[^""]*$)", RegexOptions.Singleline Or RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    Public Shared Function Process(ByRef vobj_tp As HTMLTemplate, ByRef vobj_val As Object, ByVal vstr_html As String, ByVal vstr_tag As String) As String
      Dim obj_eval As New TagEval(vobj_tp, vobj_val)
      Return obj_eval.Replace(vstr_html, vstr_tag)
    End Function

    Public Sub New(ByRef vobj_tp As HTMLTemplate)
      mobj_tp = vobj_tp
    End Sub

    Public Sub New(ByRef vobj_tp As HTMLTemplate, ByRef vobj_val As Object)
      mobj_tp = vobj_tp
      Value = vobj_val
    End Sub

    Private Function ProcessMatch(ByVal vobj_match As Match) As String
      On Error Resume Next
      Dim obj_proc As New Processor(mobj_tp)
      Dim obj_value As Object = Value
      Dim str_methods As String = vobj_match.Groups("MeThOdS").Value
      For Each str_method As String In mreg_methods.Split(str_methods)
        If String.IsNullOrEmpty(str_method) Then Exit For
        obj_value = obj_proc.Process(str_method, obj_value)
      Next
      Return DirectCast(obj_value, String)
    End Function

    Public Function Replace(ByVal vstr_html As String, ByVal vstr_tag As String) As String
      Return Regex.Replace(vstr_html, vstr_tag, New MatchEvaluator(AddressOf ProcessMatch), RegexOptions.IgnoreCase Or RegexOptions.Singleline)
    End Function

  End Class

#End Region

#Region " Declarations "

  Private Shared mobj_cache As New Hashtable

  Public Delegate Function Callback(ByVal vobj_args As List(Of String)) As String

  Private mobj_root As New HTMLTemplateData
  Private mobj_data As New Dictionary(Of Integer, HTMLTemplateData)

  Public Class Referrer
    Public Globals As New Dictionary(Of String, Object)
    Public Functions As New Dictionary(Of String, Callback)
    Public References As New Dictionary(Of String, Object)
    Public DataTables As New Dictionary(Of String, DataTable)
  End Class

  Public Refer As New Referrer
  Public CurrentPart As String
  Public Depth As Integer = 0

  Private Shared mreg_opts As RegexOptions = RegexOptions.Compiled Or RegexOptions.IgnoreCase Or RegexOptions.Singleline
  Private Shared mreg_incs As New Regex("<!--#include.*?file=\""(?<FiLe>[^\""]*)\"".*?-->", mreg_opts)
  Private Shared mreg_secs_find As New Regex("<%(?<SecName>.*?)%>(?<CoDe>.*?)<%\k<SecName>%>", mreg_opts Or RegexOptions.Multiline)
  Private Shared mreg_secs_clean As New Regex("<%.*?%>", mreg_opts Or RegexOptions.Singleline)
  Private Shared mreg_secs_set As New Regex("\[\[SECTION\:(?<SeCtIoN>.*?)\]\]", RegexOptions.Singleline Or RegexOptions.Compiled)
  Private Shared mreg_tags As New Regex("{{(?<TaG>.*?)}}", mreg_opts)
  Private Shared mreg_blocks As New Regex("<(?<TyPe>[\?|\!])(?<BlOcK>.*?)\k<TyPe>>(?<CoDe>.*?)<\k<TyPe>\k<BlOcK>\k<TyPe>>", mreg_opts Or RegexOptions.Multiline)

  Private Shared mreg_set As New Regex("<(?<TyPe>set):(?<NaMe>.*?)\s+value=\""(?<VaLuE>[^\""]*)\""\s*/?>", mreg_opts)

  Private Shared mreg_post As New Regex("<(?<TyPe>(if|ifnot|rep|data)):(?<NaMe>.*?)\s+(?<AttRiBs>.*?)>(?<CoDe>.*?)</\k<TyPe>:\k<NaMe>>", mreg_opts Or RegexOptions.Multiline)
  Private Shared mreg_attribs As New Regex("(?<NaMe>\b\w+\b)\s*=\s*(?:""(?<VaLuE>[^""]*)""|'(?<VaLuE>[^']*)'|(?<VaLuE>[^""'<>\s]+))", mreg_opts)

  Private mstr_template As String
  Private mint_item As Integer

#End Region

#Region " Properties "

  Friend Property Data() As HTMLTemplateData
    Get
      AutoAdd()
      Return mobj_data(mint_item)
    End Get
    Set(ByVal Value As HTMLTemplateData)
      AutoAdd()
      mobj_data(mint_item) = Value
    End Set
  End Property

  Public ReadOnly Property Sections() As Dictionary(Of String, HTMLTemplate)
    Get
      Return Data.Sections
    End Get
  End Property

  Default Public ReadOnly Property Item(ByVal vstr_name As String) As HTMLTemplate
    Get
      On Error Resume Next
      Dim obj_tp As New HTMLTemplate
      obj_tp.Refer = Refer
      If Data.Sections.ContainsKey(vstr_name) Then
        obj_tp = Data.Sections(vstr_name)
      Else
        Dim int_start As Integer = vstr_name.IndexOf(".")
        If int_start > 0 Then
          Dim str_name As String = vstr_name.Substring(0, int_start)
          vstr_name = vstr_name.Substring(int_start + 1)
          If Data.Sections.ContainsKey(str_name) Then
            obj_tp = Data.Sections(str_name).Item(vstr_name)
          End If
        End If
      End If
      Return obj_tp
    End Get
  End Property

  Public Property Template() As String
    Get
      Return mstr_template
    End Get
    Set(ByVal Value As String)
      Value = MapPath(Value)
      If Not mstr_template = Value Then
        ParseFile(Value)
      End If
    End Set
  End Property

  Public Property HTML() As String
    Get
      Return mobj_root.HTML
    End Get
    Set(ByVal Value As String)
      If Not mobj_root.HTML = Value Then
        mobj_root.HTML = Value
        Parse()
      End If
    End Set
  End Property

#End Region

#Region " New "

  Public Sub New()
    MyBase.New()
    AddRef("now", Date.Now.ToString)
    AddRef("today", Date.Today.ToString)
    AddFunction("rnd", AddressOf GetRnd)
    AddFunction("guid", AddressOf GetGUID)
    AddFunction("query", AddressOf GetQuery)
    AddFunction("form", AddressOf GetForm)
    AddFunction("config", AddressOf GetConfig)
    AddFunction("args", AddressOf GetArgs)
    AddFunction("data", AddressOf GetData)
    AddFunction("lang", AddressOf GetLang)
    AddFunction("str", AddressOf GetString)
    AddFunction("plugin", AddressOf GetPlugin)
  End Sub

  Public Sub New(ByVal vstr_file As String)
    MyClass.New()
    Me.Template = vstr_file
  End Sub

  Public Sub New(ByVal vstr_path As String, ByVal vstr_file As String)
    MyClass.New()
    Me.Template = IO.Path.Combine(vstr_path, vstr_file)
  End Sub

  Friend Sub New(ByVal vobj_root As HTMLTemplateData)
    MyClass.New()
    mobj_root = vobj_root
  End Sub

  Friend Sub New(ByVal vobj_root As HTMLTemplateData, ByVal vobj_data As Dictionary(Of Integer, HTMLTemplateData), ByRef vobj_refer As Referrer, ByVal vint_depth As Integer)
    MyClass.New()
    mobj_root = vobj_root
    mobj_data = New Dictionary(Of Integer, HTMLTemplateData)(vobj_data)
    Refer = vobj_refer
    Depth = vint_depth
  End Sub

#End Region

#Region " Common Methods "

  Public Function Contains(ByVal vstr_name As String) As Boolean
    If Data.Sections.ContainsKey(vstr_name) Then
      Return True
    Else
      Dim int_start As Integer = vstr_name.IndexOf(".")
      If int_start > 0 Then
        Dim str_name As String = vstr_name.Substring(0, int_start)
        vstr_name = vstr_name.Substring(int_start + 1)
        If Data.Sections.ContainsKey(str_name) Then
          Return Data.Sections(str_name).Contains(vstr_name)
        End If
      End If
    End If
  End Function

  Friend Function Clone() As HTMLTemplate
    Return New HTMLTemplate(mobj_root, mobj_data, Refer, Depth)
  End Function

#End Region

#Region " Function Methods "

  Public Sub AddFunc(ByVal vstr_tag As String, ByRef vobj_del As Callback)
    AddFunction(vstr_tag, vobj_del)
  End Sub

  Public Sub AddFunction(ByVal vstr_tag As String, ByRef vobj_del As Callback)
    If Not Refer.Functions.ContainsKey(vstr_tag) Then Refer.Functions.Add(vstr_tag, vobj_del)
  End Sub

#End Region

#Region " Reference Methods "

  Public Sub AddRef(ByVal vstr_tag As String, ByRef vobj_del As Object)
    AddReference(vstr_tag, vobj_del)
  End Sub

  Public Sub AddReference(ByVal vstr_tag As String, ByRef vobj_del As Object)
    If Not Refer.References.ContainsKey(vstr_tag) Then Refer.References.Add(vstr_tag, vobj_del)
  End Sub

#End Region

#Region " DataTable Methods "

  Public Sub AddDataTable(ByVal vstr_tag As String, ByRef vobj_table As DataTable)
    If Refer.DataTables.ContainsKey(vstr_tag) Then
      Refer.DataTables(vstr_tag) = vobj_table
    Else
      Refer.DataTables.Add(vstr_tag, vobj_table)
    End If
  End Sub

#End Region

#Region " Set Methods "

  Public Sub Add(ByVal vstr_tag As String, ByVal vobj_value As Object)
    Me.AddNew()
    Me.SetItem(vstr_tag, vobj_value)
  End Sub

  Public Sub Add(ByVal vobj_tags As IDictionary)
    Me.AddNew()
    Me.SetItems(vobj_tags)
  End Sub

  Public Sub AddNew()
    If Not String.IsNullOrEmpty(CurrentPart) Then
      Item(CurrentPart).AddNew()
    Else
      mint_item = mobj_data.Count
      mobj_data.Add(mint_item, mobj_root.Clone)
    End If
  End Sub

  Public Sub AddItem(ByVal vstr_tag As String, ByVal vobj_value As Object)
    AddNew()
    SetItem(vstr_tag, vobj_value)
  End Sub

  Public Sub SetItems(ByVal vobj_tags As IDictionary)
    For Each vstr_key As String In vobj_tags.Keys
      Me.SetItem(vstr_key, vobj_tags(vstr_key))
    Next
  End Sub

  Public Sub SetItem(ByVal vstr_tag As String, ByVal vobj_value As Object)
    If CurrentPart <> "" Then
      Item(CurrentPart).SetItem(vstr_tag, vobj_value)
    Else
      If Not NullSafe(Data.HTML) = "" Then
        If Not vobj_value Is Nothing Then
          vobj_value = NullSafe(vobj_value)
          If Not vobj_value = "" Then
            SyncLock Data.Blocks
              If Not Data.Blocks.ContainsKey(vstr_tag) Then
                Data.Blocks.Add(vstr_tag, 1)
              End If
            End SyncLock
          End If
          Data.HTML = Data.HTML.Replace("{{" & vstr_tag & "}}", vobj_value)
          Data.HTML = TagEval.Process(Me, vobj_value, Data.HTML, "{{" & vstr_tag & "(?:\.)(?<MeThOdS>(.*?))?}}")
        End If
      End If
    End If
  End Sub

#End Region

#Region " Set Custom Methods "

  Public Sub SetCustomItems(ByVal vobj_tags As IDictionary)
    For Each vstr_key As String In vobj_tags.Keys
      Me.SetCustomItem(vstr_key, vobj_tags(vstr_key).ToString)
    Next
  End Sub

  Public Sub SetCustomItem(ByVal vstr_tag As String, ByVal vstr_value As String)
    If CurrentPart <> "" Then
      Item(CurrentPart).SetCustomItem(vstr_tag, vstr_value)
    Else
      Data.HTML = Data.HTML.Replace(vstr_tag, vstr_value)
    End If
  End Sub

#End Region

#Region " Set Global Methods "

  Public Sub SetGlobalItem(ByVal vstr_tag As String, ByVal vstr_value As String)
    If Not Refer.Globals.ContainsKey(vstr_tag) Then
      Refer.Globals.Add(vstr_tag, vstr_value)
    End If

    For i As Integer = 0 To mobj_data.Values.Count - 1
      For Each obj_tp As HTMLTemplate In mobj_data(i).Sections.Values
        obj_tp.SetGlobalItem(vstr_tag, vstr_value)
      Next
    Next
  End Sub

  Public Function GlobalTagExists(ByVal vstr_tag As String) As Boolean
    On Error Resume Next
    AutoAdd()
    If Data.HTML.IndexOf("{{" & vstr_tag & "}}") >= 0 Then
      Return True
    End If
    For Each obj_data As HTMLTemplateData In mobj_data.Values
      For Each obj_tp As HTMLTemplate In obj_data.Sections.Values
        obj_tp.GlobalTagExists(vstr_tag)
      Next
    Next
  End Function

#End Region

#Region " QuickRender "

  Public Shared Function QuickRender(ByVal vstr_tp As String) As String
    Dim obj_tp As New HTMLTemplate(vstr_tp)
    Return obj_tp.Render
  End Function

  Public Shared Sub QuickRender(ByVal vobj_stream As System.IO.Stream, ByVal vstr_tp As String)
    Dim obj_tp As New HTMLTemplate(vstr_tp)
    obj_tp.Render(vobj_stream)
  End Sub

  Public Shared Sub QuickRender(ByVal vobj_writer As System.IO.TextWriter, ByVal vstr_tp As String)
    Dim obj_tp As New HTMLTemplate(vstr_tp)
    obj_tp.Render(vobj_writer)
  End Sub

#End Region

#Region " Render "

  Public Function Render(Optional ByVal vb_raw As Boolean = False) As String
    On Error Resume Next
    Dim obj_sb As New Text.StringBuilder
    Using obj_writer As New IO.StringWriter(obj_sb)
      Render(obj_writer, vb_raw)
    End Using
    Return obj_sb.ToString
  End Function

  Public Sub Render(ByVal vobj_stream As System.IO.Stream, Optional ByVal vb_raw As Boolean = False)
    On Error Resume Next
    Using obj_writer As New IO.StreamWriter(vobj_stream)
      Render(obj_writer, vb_raw)
    End Using
  End Sub

  Private mb_rendered As Boolean = False
  Public Sub Render(ByVal vobj_writer As System.IO.TextWriter, Optional ByVal vb_raw As Boolean = False)
    On Error Resume Next
    If mb_rendered Then Return
    mb_rendered = True
    AutoAdd()
    Dim obj_tp As HTMLTemplate
    Dim int_index As Integer
    Dim str_section As String
    Dim int_next As Integer
    If Request.IsAuthenticated Then
      Refer.Globals.Add("auth", 1)
      Refer.Globals.Add("userid", Security.UserID)
    End If

    For i As Integer = 0 To mobj_data.Count - 1
      mint_item = i
      If Not String.IsNullOrEmpty(Data.HTML) Then
        'process globals
        For Each str_tag As String In Refer.Globals.Keys
          SetItem(str_tag, Refer.Globals(str_tag))
        Next
        'process sets
        Data.HTML = mreg_set.Replace(Data.HTML, AddressOf SetProcess)
        'process functions
        For Each str_tag As String In Refer.Functions.Keys
          Data.HTML = TagEval.Process(Me, String.Empty, Data.HTML, "{{(?<MeThOdS>" & str_tag & "(\(.*?\))?(\.(.*?))?)}}")
        Next
        'process references
        For Each str_tag As String In Refer.References.Keys
          Data.HTML = TagEval.Process(Me, Refer.References(str_tag), Data.HTML, "{{" & str_tag & "(?<MeThOdS>\.(.*?))?}}")
        Next
        'post preprocessing
        mobj_post.Add(mint_item, New Dictionary(Of String, PostItem))
        Data.HTML = mreg_post.Replace(Data.HTML, AddressOf PrePostProcess)
        'post processing
        Dim obj_item As PostItem
        Dim str_value As String = String.Empty
        For Each str_key As String In mobj_post(mint_item).Keys
          obj_item = mobj_post(mint_item)(str_key)
          Select Case obj_item.Type.ToLower
            Case "if", "ifnot" : str_value = PostProcessConds(obj_item)
            Case "rep" : str_value = PostProcessReps(obj_item)
            Case "data" : str_value = PostProcessData(obj_item)
            Case Else : str_value = String.Empty
          End Select
          Data.HTML = Data.HTML.Replace("[[POST" & obj_item.Type.ToUpper & ":" & str_key & "]]", str_value)
        Next
        'process blocks
        If Not vb_raw Then Data.HTML = mreg_blocks.Replace(Data.HTML, AddressOf ProcessBlocks)
        'process unused tags
        If Not vb_raw Then Data.HTML = mreg_tags.Replace(Data.HTML, String.Empty)
        'process sections
        If Data.Sections.Count > 0 Then
          int_index = 0
          For Each obj_match As Match In mreg_secs_set.Matches(Data.HTML)
            int_next = obj_match.Index - int_index
            If int_next > 0 Then vobj_writer.Write(Data.HTML.Substring(int_index, int_next))
            str_section = obj_match.Groups("SeCtIoN").Value
            int_index = obj_match.Index + obj_match.Value.Length
            If Data.Sections.ContainsKey(str_section) Then
              obj_tp = Data.Sections(str_section)
              obj_tp.Refer = Refer
              obj_tp.Render(vobj_writer)
            End If
          Next
          int_next = Data.HTML.Length - int_index
          If int_next > 0 Then vobj_writer.Write(Data.HTML.Substring(int_index, int_next))
        Else
          vobj_writer.Write(Data.HTML)
        End If
      End If
    Next
  End Sub

#End Region

#Region " Private Methods "

  Private Function GetPlugin(ByVal vobj_args As List(Of String)) As String
    On Error Resume Next
    Dim str_plugin As String = vobj_args(0)
    If vobj_args.Count > 1 Then str_plugin &= "_" & vobj_args(1)
    Dim int_page As Integer = Global.GetQuery("pageid", NullSafe(Config("page.default"), 0))
    If vobj_args(0).StartsWith("hidden") Then
      Dim obj_plugin As Hidden.Interface

      obj_plugin = Plugins.Hidden.Get(str_plugin)
      If Not obj_plugin Is Nothing Then
        obj_plugin.Page = Pages.Get(int_page)
        Return obj_plugin.Render
      End If
    ElseIf vobj_args(0).StartsWith("client") Then
      Dim obj_plugin As Client.Interface
      obj_plugin = Plugins.Client.Get(str_plugin)
      If Not obj_plugin Is Nothing Then
        obj_plugin.Page = Pages.Get(int_page)
        Return obj_plugin.Render
      End If
    End If
  End Function

  Private Function GetString(ByVal vobj_args As List(Of String)) As String
    On Error Resume Next
    Return vobj_args(0)
  End Function

  Private Function GetLang(ByVal vobj_args As List(Of String)) As String
    On Error Resume Next
    Dim str_name As String = NullSafe(vobj_args(0), String.Empty)
    Dim str_default As String = NullSafe(vobj_args(1), String.Empty)
    If vobj_args.Count > 0 And Not String.IsNullOrEmpty(str_name) Then
      Return Global.GetLang(str_name, str_default)
    Else
      Return Global.GetLang()
    End If
  End Function

  Private Function GetData(ByVal vobj_args As List(Of String)) As String
    On Error Resume Next
    Dim int_start As Integer = 1
    Dim str_sql As String = vobj_args(0)
    If Not str_sql.ToLower.StartsWith("select") Then
      str_sql = vobj_args(1).Replace("{0}", "" & vobj_args(0))
      int_start += 1
    End If
    For i As Integer = int_start To vobj_args.Count - 1
      str_sql = str_sql.Replace("{" & i & "}", vobj_args(i))
    Next
    Return Global.Data.Current.Get(str_sql)
  End Function

  Private Function GetArgs(ByVal vobj_args As List(Of String)) As String
    On Error Resume Next
    Dim str_default As String = NullSafe(vobj_args(1), String.Empty)
    Dim int_num As Integer = NullSafe(vobj_args(0), -1)
    If int_num >= 0 Then
      If int_num <= Global.Args.Count - 1 Then
        Return NullSafe(Global.Args(int_num), str_default)
      End If
    End If
    Return str_default
  End Function

  Private Function GetConfig(ByVal vobj_args As List(Of String)) As String
    On Error Resume Next
    Return Global.Config(vobj_args(0))
  End Function

  Private Function GetGUID(ByVal vobj_args As List(Of String)) As String
    Return Guid.NewGuid.ToString.Trim("{", "}")
  End Function

  Private Function GetRnd(ByVal vobj_args As List(Of String)) As String
    On Error Resume Next
    Dim obj_rnd As New Random(Now.Ticks)
    If vobj_args.Count > 1 Then
      Return obj_rnd.Next(NullSafe(vobj_args(0), 0), NullSafe(vobj_args(1), 0))
    Else
      Return obj_rnd.Next
    End If
    Return Global.GetQuery(vobj_args(0))
  End Function

  Private Function GetQuery(ByVal vobj_args As List(Of String)) As String
    On Error Resume Next
    Return Global.GetQuery(vobj_args(0))
  End Function

  Private Function GetForm(ByVal vobj_args As List(Of String)) As String
    On Error Resume Next
    Return Global.GetForm(vobj_args(0))
  End Function

  Private Sub AutoAdd()
    If Depth = 0 Then
      If mobj_data.Count = 0 Then
        mint_item = mobj_data.Count
        mobj_data.Add(mint_item, mobj_root.Clone)
      End If
    End If
  End Sub

  Private Sub ParseFile(ByVal vstr_file As String)
    mstr_template = MapPath(vstr_file)
    If Not IO.File.Exists(mstr_template) Then
      Throw New System.ArgumentException("TEMPLATE: File doesn't exists. (" & mstr_template & ")")
    End If
    If Depth = 0 Then
      Dim str_key As String = mstr_template & "-" & Global.GetLang()
      mobj_root.HTML = ReadFile(mstr_template)
      mobj_root.HTML = mreg_incs.Replace(mobj_root.HTML, New MatchEvaluator(AddressOf ProcessIncludes))
      Dim str_crc As String = GetCRC(mobj_root.HTML)
      Dim obj_cache As New CacheData
      SyncLock mobj_cache
        If mobj_cache.ContainsKey(str_key) Then
          obj_cache = mobj_cache(str_key)
          If obj_cache.CRC = str_crc Then
            mobj_root = obj_cache.TemplateData
            Return
          Else
            mobj_cache.Remove(str_key)
          End If
        End If
        Parse()
        obj_cache.CRC = str_crc
        obj_cache.TemplateData = mobj_root.Clone
        mobj_cache.Add(str_key, obj_cache)
      End SyncLock
    End If
  End Sub

  Private Sub Parse()
    If Not NullSafe(mobj_root.HTML) = "" Then
      mobj_root.HTML = mreg_secs_find.Replace(mobj_root.HTML, New MatchEvaluator(AddressOf ProcessSections))
      mobj_root.HTML = mreg_secs_clean.Replace(mobj_root.HTML, String.Empty)
    End If
  End Sub

  Private Function ProcessBlocks(ByVal vobj_match As Match) As String
    On Error Resume Next
    Dim b_not As Boolean = (vobj_match.Groups("TyPe").Value = "!")
    Dim b_display As Boolean = b_not
    Dim str_name As String = vobj_match.Groups("BlOcK").Value
    If Data.Blocks.Contains(str_name) Then
      If Data.Blocks(str_name) = 1 Then
        b_display = Not b_not
      End If
    End If
    If b_display Then
      Return mreg_blocks.Replace(vobj_match.Groups("CoDe").Value, New MatchEvaluator(AddressOf ProcessBlocks))
    Else
      Return String.Empty
    End If
  End Function

  Private Class PostItem
    Public Type As String
    Public Name As String
    Public Attribs As MatchCollection
    Public Code As String
  End Class

  Private mobj_post As New Dictionary(Of Integer, Dictionary(Of String, PostItem))
  Private Function PrePostProcess(ByVal vobj_match As Match) As String
    On Error Resume Next
    If vobj_match IsNot Nothing Then
      Dim str_type As String = vobj_match.Groups("TyPe").Value
      Dim str_name As String = vobj_match.Groups("NaMe").Value
      Dim str_key As String = str_type.ToLower & "-" & str_name
      If Not mobj_post(mint_item).ContainsKey(str_key) Then
        Dim obj_item As New PostItem
        obj_item.Type = str_type
        obj_item.Name = str_name
        obj_item.Attribs = mreg_attribs.Matches(vobj_match.Groups("AttRiBs").Value)
        'obj_item.Code = mreg_post.Replace(vobj_match.Groups("CoDe").Value, AddressOf PrePostProcess)
        obj_item.Code = vobj_match.Groups("CoDe").Value
        mobj_post(mint_item).Add(str_key, obj_item)
        Return "[[POST" & str_type.ToUpper & ":" & str_key & "]]"
      End If
    End If
    Return String.Empty
  End Function

  Private Function PostProcessReps(ByVal vobj_item As PostItem) As String
    On Error Resume Next
    If vobj_item.Attribs.Count > 0 Then
      Dim str_type As String = String.Empty
      Dim str_frmt As String = String.Empty
      Dim str_start As String = String.Empty
      Dim str_end As String = String.Empty
      Dim str_step As String = String.Empty
      Dim str_sel As String = String.Empty
      For Each obj_match As Match In vobj_item.Attribs
        Select Case obj_match.Groups("NaMe").ToString.ToLower
          Case "type" : str_type = NullSafe(obj_match.Groups("VaLuE"))
          Case "format", "frmt" : str_frmt = NullSafe(obj_match.Groups("VaLuE"))
          Case "start" : str_start = NullSafe(obj_match.Groups("VaLuE"))
          Case "end", "count" : str_end = NullSafe(obj_match.Groups("VaLuE"))
          Case "step" : str_step = NullSafe(obj_match.Groups("VaLuE"))
          Case "sel" : str_sel = NullSafe(obj_match.Groups("VaLuE"))
        End Select
      Next

      Dim obj_tp As New HTMLTemplate()
      obj_tp.Refer = Refer

      obj_tp.HTML = vobj_item.Code
      Select Case str_type.ToLower
        Case "time"
          str_frmt = NullSafe(str_frmt, "h:mm tt")
          Dim int_start As Integer = NullSafe(str_start, 0)
          Dim int_end As Integer = NullSafe(str_end, 23)
          Dim int_step As Integer = NullSafe(str_step, 60)
          Dim obj_date As New Date
          Dim str_time As String
          Dim int_index As Integer = 0
          If int_end >= int_start Then
            For i As Integer = int_start To int_end 'hours
              For j As Integer = 0 To 59 Step int_step
                int_index += 1
                obj_tp.AddItem("i", int_index)
                str_time = Date.ParseExact(i & ":" & j, "H:m", CultureInfo.InvariantCulture).ToString(str_frmt)
                obj_tp.SetItem("val", str_time)
                If str_sel = str_time Then obj_tp.SetItem("sel", str_time)
              Next
            Next
            Return obj_tp.Render
          End If
          'Case "date"
        Case Else
          Dim int_start As Integer = NullSafe(str_start, 1)
          Dim int_end As Integer = NullSafe(str_end, 1)
          Dim int_sel As Integer = NullSafe(str_sel, -1)
          If int_end >= int_start Then
            For i As Integer = int_start To int_end Step NullSafe(str_step, 1)
              obj_tp.AddItem("i", i)
              If int_sel = i Then obj_tp.SetItem("sel", i)
            Next
            Return obj_tp.Render(True)
          End If
      End Select

    End If
    Return String.Empty
  End Function

  Private Function PostProcessData(ByVal vobj_item As PostItem) As String
    On Error Resume Next
    If vobj_item.Attribs.Count > 0 Then
      Dim str_datatable As String = String.Empty
      Dim str_sql As String = String.Empty
      Dim str_table As String = String.Empty
      Dim str_order As String = String.Empty
      Dim int_count As Integer = 0
      Dim str_step As String = String.Empty
      Dim str_sel As String = String.Empty
      Dim int_size As Integer = 0
      Dim obj_fields As New Dictionary(Of String, Object)

      For Each obj_match As Match In vobj_item.Attribs
        Select Case obj_match.Groups("NaMe").ToString.ToLower
          Case "datatable" : str_datatable = NullSafe(obj_match.Groups("VaLuE"))
          Case "sql" : str_sql = NullSafe(obj_match.Groups("VaLuE"))
          Case "table" : str_table = NullSafe(obj_match.Groups("VaLuE"))
          Case "order" : str_order = NullSafe(obj_match.Groups("VaLuE"))
          Case "count", "top" : int_count = NullSafe(obj_match.Groups("VaLuE"), 0)
          Case "size" : int_size = NullSafe(obj_match.Groups("VaLuE"), 0)
          Case Else : obj_fields.Add(obj_match.Groups("NaMe").ToString.ToLower, NullSafe(obj_match.Groups("VaLuE")))
        End Select
      Next

      Dim obj_table As DataTable = Nothing
      If Not String.IsNullOrEmpty(str_datatable) Then
        If Refer.DataTables.ContainsKey(str_datatable) Then obj_table = Refer.DataTables(str_datatable)
      ElseIf Not String.IsNullOrEmpty(str_sql) Then
        obj_table = Global.Data.Current.Table(str_sql)
      Else
        Dim obj_query As New QueryBuilder(str_table)
        For Each str_field As String In obj_fields.Keys
          obj_query.Where(str_field, obj_fields(str_field))
        Next
        If Not String.IsNullOrEmpty(str_order) Then obj_query.Order(str_order)
        obj_table = Global.Data.Current.Table(obj_query)
      End If

      Dim obj_tp As New HTMLTemplate()
      obj_tp.Refer = Refer

      obj_tp.HTML = vobj_item.Code.Replace("<$data$>", "<%data%>")

      Dim int_pages As Integer
      Dim int_page As Integer = Global.GetQuery("page", 1)
      Dim int_start As Integer
      Dim int_end As Integer
      Dim int_rows As Integer = obj_table.Rows.Count
      Dim obj_data As HTMLTemplate = obj_tp
      If int_rows > 0 Then
        If obj_tp.Sections.ContainsKey("data") Then
          If int_size > 0 Then
            int_pages = int_rows \ int_size
            If (int_rows Mod int_size) > 0 Then
              int_pages += 1
            End If
            If int_page < 1 Or int_page > int_pages Then
              int_page = 1
            End If
            int_start = int_size * (int_page - 1)
            int_end = int_start + int_size - 1
            If int_end > int_rows - 1 Then
              int_end = int_rows - 1
            End If
          Else
            int_pages = 1
            int_page = 1
            int_start = 0
            int_end = int_rows - 1
          End If
          obj_tp.SetItem("pages", int_pages)
          obj_tp.SetItem("page", int_page)
          obj_tp.SetItem("count", int_rows)

          Dim str_url As String = "/" & Join(Args.ToArray, "/")
          Dim str_page As String
          If String.IsNullOrEmpty(Request.Url.Query) Then
            str_url &= "?" & Request.Url.Query
            str_page = "&page="
          Else
            str_page = "?page="
          End If

          Dim obj_sb As New StringBuilder()
          Dim str_link As String = "<a href=""" & str_url & "{2}{0}"">{1}</a>"
          If int_page > 1 Then
            obj_tp.SetItem("first", 1)
            obj_tp.SetItem("first_url", str_url)
            obj_sb.AppendFormat("<a href=""" & str_url & """>{0}</a>", "First")
            obj_sb.Append("&nbsp;|&nbsp;")
            obj_tp.SetItem("prev", int_page - 1)
            If (int_page - 1) = 1 Then
              obj_tp.SetItem("prev_url", str_url)
              obj_sb.AppendFormat("<a href=""" & str_url & """>{0}</a>", "Prev")
            Else
              obj_tp.SetItem("prev_url", str_url & str_page & (int_page - 1))
              obj_sb.AppendFormat(str_link, int_page - 1, "Prev", str_page)
            End If
            obj_sb.Append("&nbsp;|&nbsp;")
          End If
          obj_sb.AppendFormat(" Page {0} of {1} ", int_page, int_pages)
          If int_page < int_pages Then
            obj_sb.Append("&nbsp;|&nbsp;")
            obj_tp.SetItem("next", int_page + 1)
            obj_tp.SetItem("next_url", str_url & str_page & (int_page + 1))
            obj_sb.AppendFormat(str_link, int_page + 1, "Next", str_page)
            obj_sb.Append("&nbsp;|&nbsp;")
            obj_tp.SetItem("last_url", str_url & str_page & int_pages)
            obj_sb.AppendFormat(str_link, int_pages, "Last", str_page)
          End If
          obj_tp.SetItem("paging", obj_sb.ToString)
          obj_data = obj_tp.Sections("data")
        Else
          If int_count = 0 Or int_rows < int_count Then int_count = int_rows
          int_start = 0
          int_end = int_count - 1
        End If
        Dim str_val As String
        For i As Integer = int_start To int_end
          If i > int_start Or obj_tp.Sections.ContainsKey("data") Then obj_data.AddNew()
          For Each obj_col As DataColumn In obj_table.Columns
            str_val = "" & obj_table.Rows(i)(obj_col)
            obj_data.SetItem(vobj_item.Name & "_" & obj_col.ColumnName.ToLower, str_val)
            obj_data.SetItem(obj_col.ColumnName.ToLower, str_val)
            If IsNumeric(str_val) Then
              obj_data.SetItem(vobj_item.Name & "_" & obj_col.ColumnName.ToLower & "_" & str_val, str_val)
              obj_data.SetItem(obj_col.ColumnName.ToLower & "_" & str_val, str_val)
            End If
          Next
        Next
      ElseIf obj_tp.Sections.ContainsKey("data") Then
        obj_tp.SetItem("none", 1)
        obj_tp.SetItem("pages", 1)
        obj_tp.SetItem("page", 1)
        obj_tp.SetItem("count", 0)
      Else
        Return String.Empty
      End If
      Return obj_tp.Render(True)
    End If
    Return String.Empty
  End Function

  Private Function PostProcessConds(ByVal vobj_item As PostItem) As String
    On Error Resume Next
    Dim b_not As Boolean = vobj_item.Type.ToLower.Equals("ifnot")
    If vobj_item.Attribs.Count > 0 Then
      Dim str_code As String = vobj_item.Code
      Dim str_else As String = String.Empty
      Dim int_index As Integer = str_code.IndexOf("<else:" & vobj_item.Name, StringComparison.OrdinalIgnoreCase)
      If int_index > 0 Then
        Dim int_else As Integer = str_code.IndexOf(">", int_index) + 1
        If int_else > int_index Then str_else = str_code.Substring(int_else)
        str_code = str_code.Substring(0, int_index)
      End If
      Dim b_show As Boolean = False
      For Each obj_match As Match In vobj_item.Attribs
        b_show = obj_match.Groups("VaLuE").Value.ToLower.Equals("true")
        If Not b_show Then Exit For
      Next
      Dim obj_tp As New HTMLTemplate()
      obj_tp.Refer = Refer
      obj_tp.HTML = IIf((b_not And Not b_show) Or (Not b_not And b_show), str_code, str_else)
      Return obj_tp.Render(True)
    End If
    Return String.Empty
  End Function

  Private Function SetProcess(ByVal vobj_match As Match) As String
    On Error Resume Next
    If vobj_match IsNot Nothing Then
      Dim str_name As String = vobj_match.Groups("NaMe").Value
      Dim str_value As String = vobj_match.Groups("VaLuE").Value
      If Refer.References.ContainsKey(str_name) Then
        If str_value = "++" Then str_value = NullSafe(Refer.References(str_name), 0) + 1
        If str_value = "--" Then str_value = NullSafe(Refer.References(str_name), 0) - 1
        Refer.References(str_name) = str_value
      Else
        If str_value = "++" Then str_value = 1
        If str_value = "--" Then str_value = 0
        Refer.References.Add(str_name, str_value)
      End If
    End If
    Return String.Empty
  End Function

  Private Function ProcessSections(ByVal vobj_match As Match) As String
    Dim str_name As String = vobj_match.Groups("SecName").Value
    Dim str_code As String = vobj_match.Groups("CoDe").Value
    If mobj_root.Sections.ContainsKey(str_name) Then
      Throw New System.Exception("TEMPLATE: Section [" & str_name & "] already exists.")
    End If
    Dim obj_tp As New HTMLTemplate
    obj_tp.Refer = Refer
    obj_tp.Depth = Depth + 1
    obj_tp.HTML = str_code
    mobj_root.Sections.Add(str_name, obj_tp)
    Return "[[SECTION:" & str_name & "]]"
  End Function

  Private Function ProcessIncludes(ByVal vobj_match As Match) As String
    On Error Resume Next
    Dim str_file As String = vobj_match.Groups("FiLe").Value
    Dim str_lang As String = Global.GetLang()
    str_file = str_file.Replace("[[lang]]", str_lang)
    str_file = str_file.Replace("{{lang}}", str_lang)
    str_file = MapPath(str_file)
    If IO.File.Exists(str_file) Then
      Return mreg_incs.Replace(ReadFile(str_file), New MatchEvaluator(AddressOf ProcessIncludes))
    Else
      Return "<!-- FILE INCLUDE NOT FOUND (" & vobj_match.Groups("FiLe").Value & ") -->"
    End If
  End Function

#End Region

#Region " Processor "

  Friend Class Processor
    Private mobj_tp As HTMLTemplate
    Private Shared mreg_method As New Regex("(?<MeThOd>[\w]*)(\((?<PaRaMs>.*?)\)$)?", RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.Compiled)
    Private Shared mreg_args As New Regex("""([^\""\\]*(?:\\.[^\""\\]*)*)"",?|,", RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.Compiled)
    Private Shared mreg_blanks As New Regex("((\r|\n)\s*(\r|\n))", RegexOptions.Singleline Or RegexOptions.Multiline Or RegexOptions.Compiled)

    Public Sub New(ByVal vobj_tp As HTMLTemplate)
      mobj_tp = vobj_tp
    End Sub

    Public Function Process(ByVal vstr_call As String, ByVal vstr_value As String) As String
      On Error Resume Next
      Dim obj_match As Match = mreg_method.Match(vstr_call)
      If Not obj_match Is Nothing Then
        Dim str_method As String = obj_match.Groups("MeThOd").Value
        If Not String.IsNullOrEmpty(str_method.Trim) Then
          Dim obj_args As New List(Of String)
          If Not String.IsNullOrEmpty(vstr_value) Then obj_args.Add(vstr_value)
          For Each str_arg As String In mreg_args.Split(obj_match.Groups("PaRaMs").Value)
            If Not String.IsNullOrEmpty(str_arg) Then
              obj_args.Add(str_arg)
            End If
          Next
          If mobj_tp.Refer.References.ContainsKey(str_method) Then
            Return DirectCast(mobj_tp.Refer.References(str_method), String)
          ElseIf mobj_tp.Refer.Functions.ContainsKey(str_method) Then
            Dim obj_call As Callback = mobj_tp.Refer.Functions(str_method)
            Return obj_call.Invoke(obj_args)
          Else
            Return Process(str_method, obj_args)
          End If
        End If
      End If
      Return vstr_value
    End Function

    Private Function Process(ByVal vstr_method As String, ByVal vobj_args As List(Of String)) As String
      On Error Resume Next
      Dim str_value As String = vobj_args(0)
      If Not String.IsNullOrEmpty(str_value) Then
        Select Case vstr_method
          Case "set" : str_value = 1
          Case "true", "false"
            Return String.Equals(str_value, vstr_method, StringComparison.OrdinalIgnoreCase)
          Case "exists"
            Dim str_file As String = str_value
            If vobj_args.Count > 1 Then str_file = String.Format(vobj_args(1), str_file)
            str_file = MapPath(str_file)
            Return IO.File.Exists(str_file) Or IO.Directory.Exists(str_file)
          Case "equals", "eq", "is"
            Return String.Equals(str_value, vobj_args(1), StringComparison.OrdinalIgnoreCase)
          Case "notequal", "not"
            Return Not String.Equals(str_value, vobj_args(1), StringComparison.OrdinalIgnoreCase)
          Case "lessthan", "lt"
            Return NullSafe(str_value, 0) < NullSafe(vobj_args(1), 0)
          Case "greaterthan", "gt"
            Return NullSafe(str_value, 0) > NullSafe(vobj_args(1), 0)
          Case "lessthanorequals", "lte"
            Return NullSafe(str_value, 0) <= NullSafe(vobj_args(1), 0)
          Case "greaterthanorequals", "gte"
            Return NullSafe(str_value, 0) >= NullSafe(vobj_args(1), 0)
          Case "length", "len" : Return str_value.Length
          Case "tolower", "lcase" : str_value = str_value.ToLower
          Case "toupper", "ucase" : str_value = str_value.ToUpper
          Case "toproper", "pcase" : str_value = StrConv(str_value, VbStrConv.ProperCase)
          Case "select" : Return IIf(String.Equals(str_value, vobj_args(1), StringComparison.OrdinalIgnoreCase), " selected", "")
          Case "check" : Return IIf(String.Equals(str_value, vobj_args(1), StringComparison.OrdinalIgnoreCase), " checked", "")
          Case "replace"
            str_value = str_value.Replace(vbNewLine, "\n").Replace(vbTab, "\t")
            str_value = str_value.Replace(vobj_args(1), vobj_args(2))
            str_value = str_value.Replace("\n", vbNewLine).Replace("\t", vbTab)
          Case "left" : str_value = Left(str_value, NullSafe(vobj_args(1), str_value.Length))
          Case "right" : str_value = Right(str_value, NullSafe(vobj_args(1), str_value.Length))
          Case "remove"
            If vobj_args.Count > 1 Then
              For i As Integer = 1 To vobj_args.Count - 1
                str_value = str_value.Replace(vobj_args(i), String.Empty)
              Next
            End If
          Case "iif" : str_value = IIf(str_value = vobj_args(1), vobj_args(2), vobj_args(3))
          Case "ifset" : str_value = vobj_args(1)
          Case "trim" : str_value = str_value.Trim()
          Case "trimstart" : str_value = str_value.TrimStart()
          Case "trimend" : str_value = str_value.TrimEnd()
          Case "insert" : str_value = str_value.Insert(NullSafe(vobj_args(1), 0), vobj_args(2))
          Case "regex" : str_value = Regex.Replace(str_value, vobj_args(1).ToString, vobj_args(2))
          Case "contains" : str_value = IIf(str_value.IndexOf(vobj_args(1).ToString) >= 0, 1, 0)
          Case "http" : str_value = "http://" & str_value.ToLower.Replace("http://", "").Replace("http//", "")
          Case "bold" : str_value = "<b>" & str_value & "</b>"
          Case "max", "maxlen"
            Dim int_max As Integer = NullSafe(vobj_args(1), str_value.Length)
            If str_value.Length > int_max Then str_value = str_value.Substring(0, int_max)
          Case "summary"
            Dim int_max As Integer = NullSafe(vobj_args(1), str_value.Length)
            str_value = mreg_blanks.Replace(str_value, vbNewLine)
            If str_value.Length > int_max Then str_value = str_value.Substring(0, int_max) & "..."
          Case "format" : str_value = String.Format(vobj_args(1).ToString, str_value)
          Case "currency", "cur" : str_value = FormatCurrency(str_value, NullSafe(vobj_args(1).ToString, 2))
          Case "uni", "unicode" : str_value = ConvertUnicode(str_value)
          Case "date"
            Dim obj_date As Date ' = NullSafe(str_value, Date.Now)
            If Date.TryParse(str_value, obj_date) Then
              Select Case vobj_args(1).ToLower
                Case "short" : str_value = obj_date.ToShortDateString
                Case "long" : str_value = obj_date.ToLongDateString
                Case "year" : str_value = obj_date.Year.ToString
                Case "month" : str_value = obj_date.Month.ToString
                Case "day" : str_value = obj_date.Day.ToString
                Case "weekday" : str_value = obj_date.DayOfWeek.ToString
                Case "monthname" : str_value = obj_date.ToString("MMMM")
                Case Else : str_value = obj_date.ToString(vobj_args(1))
              End Select
            Else
              str_value = String.Empty
            End If
          Case "clean"
            str_value = Regex.Replace(str_value, "\t", "  ", RegexOptions.Compiled)
            str_value = mreg_blanks.Replace(str_value, vbNewLine)
          Case "html" : str_value = Server.HtmlEncode(str_value)
          Case "escape" : str_value = Regex.Escape(str_value)
          Case "javascript", "js" : Return JSEncode(str_value)
          Case "strip", "striphtml" : str_value = StripHtml(str_value)
          Case "url" : str_value = Server.UrlEncode(str_value)
          Case "urlname" : str_value = GetUrlName(str_value)
          Case "wrap" : str_value = vobj_args(1).ToString.Replace("{0}", str_value)
          Case "start", "prefix" : str_value = vobj_args(1).ToString & str_value
          Case "end", "suffix" : str_value = str_value & vobj_args(1).ToString
          Case "first" : str_value = str_value.Substring(0, NullSafe(vobj_args(1), str_value.Length))
          Case "last" : str_value = str_value.Substring(str_value.Length - NullSafe(vobj_args(1), 0))
          Case "concat" : str_value = String.Concat(vobj_args.ToArray)
          Case "join" : str_value = Join(str_value.Split(vobj_args(1)), vobj_args(1))
          Case "mask"
            Dim int_mask As Integer = NullSafe(vobj_args(1), 0)
            Dim int_num As Integer = Math.Abs(int_mask)
            Dim int_len As Integer = str_value.Length
            If int_num > int_len Then int_num = int_len
            Dim str_mask As String = String.Empty
            For i As Integer = 1 To int_len
              str_mask &= "*"
            Next
            If int_mask > 0 Then
              str_mask = str_mask.Substring(int_num)
              str_mask = str_value.Substring(0, int_num) & str_mask
            ElseIf int_mask < 0 Then
              str_mask = str_mask.Substring(int_num)
              str_mask = str_mask & str_value.Substring(int_len - int_num)
            End If
            str_value = str_mask
        End Select
      Else
        Select Case vstr_method
          Case "set" : str_value = 0
          Case "ifset" : str_value = vobj_args(2)
          Case "default" : str_value = vobj_args(1)
        End Select
      End If
      Return str_value
    End Function

  End Class

#End Region

End Class