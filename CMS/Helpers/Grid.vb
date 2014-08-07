Imports System.Text.RegularExpressions

Public Class Grid

#Region " Formatters "

  Public Class Formatters

    Public Shared Function YesNo(ByVal vstr_value As String) As String
      Select Case vstr_value.ToLower
        Case "yes", "y", "1", "true" : Return "Yes"
        Case Else : Return "No"
      End Select
    End Function

    Public Shared Function TrueFalse(ByVal vstr_value As String) As String
      Select Case vstr_value.ToLower
        Case "yes", "y", "1", "true" : Return "True"
        Case Else : Return "False"
      End Select
    End Function

    Public Shared Function ShortDate(ByVal vstr_value As String) As String
      On Error Resume Next
      Dim dat_value As Date
      If Date.TryParse(vstr_value, dat_value) Then
        Return dat_value.ToShortDateString
      End If
      Return String.Empty
    End Function

    Public Shared Function LongDate(ByVal vstr_value As String) As String
      On Error Resume Next
      Dim dat_value As Date
      If Date.TryParse(vstr_value, dat_value) Then
        Return dat_value.ToLongDateString
      End If
      Return String.Empty
    End Function

    Public Shared Function Proper(ByVal vstr_value As String) As String
      Return StrConv(vstr_value, VbStrConv.ProperCase)
    End Function

    Public Shared Function Lower(ByVal vstr_value As String) As String
      Return vstr_value.ToLower
    End Function

    Public Shared Function Upper(ByVal vstr_value As String) As String
      Return vstr_value.ToUpper
    End Function

    Public Shared Function Money(ByVal vstr_value As String) As String
      On Error Resume Next
      Return FormatCurrency(vstr_value, 2)
    End Function

    Public Shared Function Summary(ByVal vstr_value As String) As String
      On Error Resume Next
      If vstr_value.Length > 255 Then vstr_value = vstr_value.Substring(0, 255)
      Return vstr_value
    End Function

  End Class

#End Region

#Region " Columns "

  Public Interface IColumn
    Property Name() As String
    Property Title() As String
    Property Data() As Data.Engine
    Property Row() As DataRow
    Property Sortable() As Boolean
    Property Formatter() As Callback
    Property Width() As Integer
    Function Render(ByVal vstr_value As String) As String
  End Interface

#Region " DefaultColumn "
  Public Class DefaultColumn
    Implements IColumn
    Private mstr_name As String
    Private mstr_title As String
    Private mobj_row As DataRow
    Private mb_sortable As Boolean
    Private mobj_format As Callback
    Private mobj_data As Data.Engine
    Private mint_width As Integer

#Region " Properties "
    Public Property Data() As Data.Engine Implements IColumn.Data
      Get
        If mobj_data Is Nothing Then mobj_data = Global.Data.Current
        Return mobj_data
      End Get
      Set(ByVal value As Data.Engine)
        mobj_data = value
      End Set
    End Property
    Public Property Row() As System.Data.DataRow Implements IColumn.Row
      Get
        Return mobj_row
      End Get
      Set(ByVal Value As System.Data.DataRow)
        mobj_row = Value
      End Set
    End Property
    Public Property Sortable() As Boolean Implements IColumn.Sortable
      Get
        Return mb_sortable
      End Get
      Set(ByVal value As Boolean)
        mb_sortable = value
      End Set
    End Property
    Public Property Width() As Integer Implements IColumn.Width
      Get
        Return mint_width
      End Get
      Set(ByVal value As Integer)
        mint_width = value
      End Set
    End Property
    Public Property Formatter() As Callback Implements IColumn.Formatter
      Get
        Return mobj_format
      End Get
      Set(ByVal value As Callback)
        mobj_format = value
      End Set
    End Property
    Public Property Title() As String Implements IColumn.Title
      Get
        Return mstr_title
      End Get
      Set(ByVal Value As String)
        mstr_title = Value
      End Set
    End Property
    Public Property Name() As String Implements IColumn.Name
      Get
        Return mstr_name
      End Get
      Set(ByVal Value As String)
        mstr_name = Value
      End Set
    End Property
#End Region

    Public Overridable Function ProcessFields(ByVal vobj_match As Match) As String
      Dim str_name As String = vobj_match.Groups("fIeLd").Value
      If Row.Table.Columns.Contains(str_name) Then
        Return "" & Row(str_name)
      End If
      Return String.Empty
    End Function

    Public Overridable Function ProcessFields(ByVal vstr_value As String) As String
      On Error Resume Next
      Return Regex.Replace(vstr_value, "{(?<fIeLd>.*?)}", New MatchEvaluator(AddressOf ProcessFields), RegexOptions.Singleline)
    End Function

    Public Overridable Function Render(ByVal vstr_value As String) As String Implements IColumn.Render
      On Error Resume Next
      vstr_value = ProcessFields(vstr_value)
      If mobj_format IsNot Nothing Then vstr_value = mobj_format(vstr_value)
      Return vstr_value
    End Function
  End Class
#End Region

#Region " SQLColumn "
  Public Class SQLColumn
    Inherits DefaultColumn
    Private mstr_sql As String

    Public WriteOnly Property SQL() As String
      Set(ByVal Value As String)
        mstr_sql = Value
      End Set
    End Property

    Public Shadows Function ProcessFields(ByVal vobj_match As Match) As String
      Dim str_name As String = vobj_match.Groups("fIeLd").Value
      If Row.Table.Columns.Contains(str_name) Then
        Return NullSafe(Row(str_name)).Replace("'", "''")
      End If
      Return String.Empty
    End Function

    Public Overrides Function Render(ByVal vstr_value As String) As String
      Try
        Dim str_sql As String = mstr_sql
        str_sql = str_sql.Replace("{0}", vstr_value.Replace("'", "''"))
        str_sql = MyBase.ProcessFields(str_sql)
        vstr_value = NullSafe(Data.Get(str_sql), vstr_value)
      Catch ex As Exception
      End Try
      Return MyBase.Render(vstr_value)
    End Function
  End Class
#End Region

#Region " HTMLColumn "
  Public Class HTMLColumn
    Inherits DefaultColumn
    Private mstr_html As String

    Public WriteOnly Property HTML() As String
      Set(ByVal Value As String)
        mstr_html = Value
      End Set
    End Property

    Public Overrides Function Render(ByVal vstr_value As String) As String
      On Error Resume Next
      vstr_value = mstr_html.Replace("{0}", vstr_value)
      Return MyBase.Render(vstr_value)
    End Function
  End Class
#End Region

#Region " CustomColumn "
  Public Class CustomColumn
    Inherits DefaultColumn

    Private mdel_callback As Callback
    Public Property Callback() As Callback
      Get
        Return mdel_callback
      End Get
      Set(ByVal Value As Callback)
        mdel_callback = Value
      End Set
    End Property

    Public Overrides Function Render(ByVal vstr_value As String) As String
      If Not mdel_callback Is Nothing Then vstr_value = mdel_callback.Invoke(vstr_value)
      Return MyBase.Render(vstr_value)
    End Function
  End Class
#End Region

#End Region

  Private mstr_title As String = ""
  Private mobj_names As New List(Of String)
  Private mobj_columns As New List(Of IColumn)
  Private mstr_template As String
  Private mstr_sql As String
  Private mint_pagesize As Integer = 50
  Private mobj_currentrow As DataRow = Nothing
  Private mstr_current As String
  Private mobj_data As Data.Engine
  Public Delegate Function Callback(ByVal vstr_value As String) As String

#Region " Properties "

  Public Property Data() As Data.Engine
    Get
      If mobj_data Is Nothing Then mobj_data = Global.Data.Current
      Return mobj_data
    End Get
    Set(ByVal value As Data.Engine)
      mobj_data = value
    End Set
  End Property

  Public Property Template() As String
    Get
      If String.IsNullOrEmpty(mstr_template) Then mstr_template = "/cms/templates/grid.htm"
      Return mstr_template
    End Get
    Set(ByVal Value As String)
      mstr_template = Value
    End Set
  End Property

  Public Property PageSize() As Integer
    Get
      Return mint_pagesize
    End Get
    Set(ByVal Value As Integer)
      mint_pagesize = Value
    End Set
  End Property

  Public Property Title() As String
    Get
      Return mstr_title
    End Get
    Set(ByVal Value As String)
      mstr_title = Value
    End Set
  End Property

  Public ReadOnly Property CurrentRow() As DataRow
    Get
      Return mobj_currentrow
    End Get
  End Property

  Public ReadOnly Property CurrentColumnName() As String
    Get
      Return mstr_current
    End Get
  End Property

#End Region

#Region " New "

  Public Sub New()
    MyBase.New()
  End Sub

  Public Sub New(ByVal vstr_sql As String)
    MyBase.New()
    mstr_sql = vstr_sql
  End Sub

#End Region

#Region " Clear "

  Public Sub Clear()
    mobj_names.Clear()
    mobj_columns.Clear()
  End Sub

#End Region

#Region " AddColumn "

  Public Sub AddColumn(ByVal vobj_col As IColumn)
    vobj_col.Data = Data
    mobj_names.Add(vobj_col.Name)
    mobj_columns.Add(vobj_col)
  End Sub

  Public Sub AddColumn(ByVal vstr_name As String, ByVal vstr_title As String)
    Dim obj_col As New DefaultColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    AddColumn(obj_col)
  End Sub

  Public Sub AddColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vb_sortable As Boolean)
    Dim obj_col As New DefaultColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    obj_col.Sortable = vb_sortable
    AddColumn(obj_col)
  End Sub

  Public Sub AddColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vobj_format As Callback)
    Dim obj_col As New DefaultColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    obj_col.Formatter = vobj_format
    AddColumn(obj_col)
  End Sub

  Public Sub AddColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vb_sortable As Boolean, ByVal vobj_format As Callback)
    Dim obj_col As New DefaultColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    obj_col.Sortable = vb_sortable
    obj_col.Formatter = vobj_format
    AddColumn(obj_col)
  End Sub

#End Region

#Region " AddSQLColumn "

  Public Sub AddSQLColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vstr_sql As String)
    Dim obj_col As New SQLColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    obj_col.SQL = vstr_sql
    AddColumn(obj_col)
  End Sub

  Public Sub AddSQLColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vstr_sql As String, ByVal vb_sortable As Boolean)
    Dim obj_col As New SQLColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    obj_col.SQL = vstr_sql
    obj_col.Sortable = vb_sortable
    AddColumn(obj_col)
  End Sub

  Public Sub AddSQLColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vstr_sql As String, ByVal vobj_format As Callback)
    Dim obj_col As New SQLColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    obj_col.SQL = vstr_sql
    obj_col.Formatter = vobj_format
    AddColumn(obj_col)
  End Sub

  Public Sub AddSQLColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vstr_sql As String, ByVal vb_sortable As Boolean, ByVal vobj_format As Callback)
    Dim obj_col As New SQLColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    obj_col.SQL = vstr_sql
    obj_col.Sortable = vb_sortable
    obj_col.Formatter = vobj_format
    AddColumn(obj_col)
  End Sub

#End Region

#Region " AddHTMLColumn "

  Public Sub AddHTMLColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vstr_html As String)
    Dim obj_col As New HTMLColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    obj_col.HTML = vstr_html
    AddColumn(obj_col)
  End Sub

  Public Sub AddHTMLColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vstr_html As String, ByVal vb_sortable As Boolean)
    Dim obj_col As New HTMLColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    obj_col.HTML = vstr_html
    obj_col.Sortable = vb_sortable
    AddColumn(obj_col)
  End Sub

  Public Sub AddHTMLColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vstr_html As String, ByVal vobj_format As Callback)
    Dim obj_col As New HTMLColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    obj_col.HTML = vstr_html
    obj_col.Formatter = vobj_format
    AddColumn(obj_col)
  End Sub

  Public Sub AddHTMLColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vstr_html As String, ByVal vb_sortable As Boolean, ByVal vobj_format As Callback)
    Dim obj_col As New HTMLColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    obj_col.HTML = vstr_html
    obj_col.Sortable = vb_sortable
    obj_col.Formatter = vobj_format
    AddColumn(obj_col)
  End Sub

#End Region

#Region " AddCustomColumn "

  Public Sub AddCustomColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vobj_call As Callback)
    Dim obj_col As New CustomColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    obj_col.Callback = vobj_call
    AddColumn(obj_col)
  End Sub

  Public Sub AddCustomColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vobj_call As Callback, ByVal vb_sortable As Boolean)
    Dim obj_col As New CustomColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    obj_col.Callback = vobj_call
    obj_col.Sortable = vb_sortable
    AddColumn(obj_col)
  End Sub

  Public Sub AddCustomColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vobj_call As Callback, ByVal vobj_format As Callback)
    Dim obj_col As New CustomColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    obj_col.Callback = vobj_call
    obj_col.Formatter = vobj_format
    AddColumn(obj_col)
  End Sub

  Public Sub AddCustomColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vobj_call As Callback, ByVal vb_sortable As Boolean, ByVal vobj_format As Callback)
    Dim obj_col As New CustomColumn
    obj_col.Data = Data
    obj_col.Name = vstr_name
    obj_col.Title = vstr_title
    obj_col.Callback = vobj_call
    obj_col.Sortable = vb_sortable
    obj_col.Formatter = vobj_format
    AddColumn(obj_col)
  End Sub

#End Region

#Region " Render "

  Public Function Render() As String
    Return Render(mstr_sql)
  End Function

  Public Function Render(ByVal vstr_sql As String) As String
    On Error Resume Next
    If Not vstr_sql.ToLower.Trim.StartsWith("select") Then
      Return Render(New QueryBuilder(vstr_sql))
    End If
    If GetQuery("json", 0) = 1 Then
      Dim obj_table As DataTable = Data.Table(vstr_sql)
      Return Render(obj_table)
    Else
      Return Render(False)
    End If
  End Function

  Public Function Render(ByVal vobj_query As QueryBuilder) As String
    On Error Resume Next
    If GetQuery("json", 0) = 1 Then
      Dim str_sort As String = GetForm("sort").ToLower
      Dim str_dir As String = GetForm("dir", "desc").ToLower
      If Not String.IsNullOrEmpty(str_sort) Then
        str_sort = str_sort.ToLower
        Select Case str_dir.ToLower
          Case "desc" : vobj_query.Order(str_sort, Sort.Desc)
          Case Else : vobj_query.Order(str_sort, Sort.Asc)
        End Select
      End If
      Dim obj_table As DataTable = Data.Table(vobj_query)
      Return Render(obj_table)
    Else
      Return Render(True)
    End If
  End Function

  Private Function Render(ByVal vb_sortable As Boolean) As String
    Dim obj_tp As New HTMLTemplate(Template)
    If Not String.IsNullOrEmpty(mstr_title) Then obj_tp.SetItem("title", mstr_title)
    Dim str_url As String = RawUrl.ToLower
    str_url &= IIf(str_url.IndexOf("?") > 0, "&", "?") & "json=1"
    obj_tp.SetItem("url", str_url)
    Dim obj_col As IColumn
    For i As Integer = 0 To mobj_names.Count - 1
      obj_col = mobj_columns(i)
      obj_tp("fields").AddItem("name", obj_col.Name)
      obj_tp("cols").AddItem("field", obj_col.Name)
      obj_tp("cols").SetItem("title", obj_col.Title)
      If vb_sortable Then obj_tp("cols").SetItem("sort", obj_col.Sortable.ToString.ToLower)
      If obj_col.Width > 0 Then obj_tp("cols").SetItem("width", obj_col.Width)
    Next
    Return obj_tp.Render
  End Function

  Private Function Render(ByRef vobj_table As DataTable) As String
    On Error Resume Next
    Dim obj_row As DataRow
    Dim obj_col As IColumn
    Dim obj_value As Object
    Dim int_start As Integer = GetRequest("start", 0)
    Dim int_end As Integer = 0
    Dim int_size As Integer = GetRequest("limit", 0)
    Dim int_rows As Integer = vobj_table.Rows.Count - 1
    If int_size > 0 Then
      int_end = int_start + int_size - 1
      If int_end > int_rows Then int_end = int_rows
    Else
      int_end = int_rows
    End If

    Dim obj_sb As New StringBuilder
    obj_sb.AppendLine("{""totalCount"":""" & int_rows + 1 & """,""records"":[")
    Dim str_fmt As String = """{0}"":""{1}"","

    For i As Integer = int_start To int_end
      obj_row = vobj_table.Rows(i)
      mobj_currentrow = obj_row
      obj_sb.Append("{")
      For j As Integer = 0 To mobj_names.Count - 1
        obj_col = mobj_columns(j)
        mstr_current = obj_col.Name
        obj_col.Row = obj_row
        If vobj_table.Columns.Contains(obj_col.Name) Then
          obj_value = "" & obj_row(obj_col.Name)
        Else
          obj_value = String.Empty
        End If
        obj_sb.AppendFormat(str_fmt, JSEncode(obj_col.Name), JSEncode(obj_col.Render(obj_value)))
      Next
      obj_sb.Remove(obj_sb.Length - 1, 1)
      obj_sb.Append("},")
    Next
    obj_sb.Remove(obj_sb.Length - 1, 1)
    obj_sb.AppendLine("]}")
    Return obj_sb.ToString.Replace("},", "}," & vbNewLine)
  End Function

#End Region

End Class
