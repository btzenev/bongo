Imports System.Text.RegularExpressions

Public Class Editor

#Region " Columns "

  Public Interface IColumn
    Property ID() As Integer
    Property Name() As String
    Property Title() As String
    Property Data() As Data.Engine
    Property Row() As DataRow
    Property Tab() As String
    Property DefaultValue() As Object
    ReadOnly Property Scripts() As String
    ReadOnly Property Hidden() As Boolean
    Function Render(ByVal vstr_value As String) As String
    Function Save(ByVal vstr_value As String) As String
  End Interface

#Region " BaseColumn "
  Public MustInherit Class BaseColumn
    Implements IColumn

    Private mint_id As Integer
    Private mstr_name As String
    Private mstr_title As String
    Private mobj_row As DataRow
    Private mobj_default As Object
    Private mstr_tab As String
    Private mobj_data As Data.Engine

    Public Property Data() As Data.Engine Implements IColumn.Data
      Get
        If mobj_data Is Nothing Then mobj_data = Global.Data.Current
        Return mobj_data
      End Get
      Set(ByVal value As Data.Engine)
        mobj_data = value
      End Set
    End Property

    Public Property ID() As Integer Implements IColumn.ID
      Get
        Return mint_id
      End Get
      Set(ByVal value As Integer)
        mint_id = value
      End Set
    End Property

    Public Property Row() As DataRow Implements IColumn.Row
      Get
        Return mobj_row
      End Get
      Set(ByVal Value As DataRow)
        mobj_row = Value
      End Set
    End Property

    Public Property DefaultValue() As Object Implements IColumn.DefaultValue
      Get
        Return mobj_default
      End Get
      Set(ByVal Value As Object)
        mobj_default = Value
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

    Public Property Title() As String Implements IColumn.Title
      Get
        Return mstr_title
      End Get
      Set(ByVal Value As String)
        mstr_title = Value
      End Set
    End Property

    Public Property Tab() As String Implements IColumn.Tab
      Get
        Return mstr_tab
      End Get
      Set(ByVal Value As String)
        mstr_tab = Value
      End Set
    End Property

    Public Overridable ReadOnly Property Scripts() As String Implements IColumn.Scripts
      Get
        Return String.Empty
      End Get
    End Property

    Public Overridable ReadOnly Property Hidden() As Boolean Implements IColumn.Hidden
      Get
        Return False
      End Get
    End Property

    Public MustOverride Function Render(ByVal vstr_value As String) As String Implements IColumn.Render

    Public Overridable Function Save(ByVal vstr_value As String) As String Implements IColumn.Save
      Return vstr_value
    End Function


    Public Overridable Function ProcessFields(ByVal vobj_match As Match) As String
      Try
        Dim str_name As String = vobj_match.Groups("fIeLd").Value
        If Not Row Is Nothing Then
          If Row.Table.Columns.Contains(str_name) Then
            Return "" & Row(str_name)
          End If
        End If
        Return vobj_match.Value
      Catch ex As Exception
        Return vbNewLine & "<!--" & Server.HtmlEncode(ex.ToString) & "-->"
      End Try
    End Function

    Public Overridable Function ProcessSQLFields(ByVal vobj_match As Match) As String
      Return ProcessFields(vobj_match)
    End Function

    Public Overridable Function ProcessFields(ByVal vstr_value As String, ByRef vobj_match As MatchEvaluator) As String
      On Error Resume Next
      Return Regex.Replace(vstr_value, "(?is){(?<fIeLd>.*?)}", vobj_match, RegexOptions.Singleline Or RegexOptions.Compiled)
    End Function

    Public Overridable Function ProcessFields(ByVal vstr_value As String) As String
      Return ProcessFields(vstr_value, New MatchEvaluator(AddressOf ProcessFields))
    End Function

    Public Sub New()

    End Sub
  End Class
#End Region

#Region " ViewColumn "
  Public Class ViewColumn
    Inherits BaseColumn

    Private mb_auto As Boolean

    Public Sub New()
    End Sub

    Public Sub New(ByVal vb_auto As Boolean)
      mb_auto = vb_auto
    End Sub

    Public Overrides Function Render(ByVal vstr_value As String) As String
      If mb_auto Or vstr_value = "" Then vstr_value = "Auto"
      Return vstr_value
    End Function

  End Class
#End Region

#Region " InputColumn "
  Public Class InputColumn
    Inherits BaseColumn

    Public Overrides Function Render(ByVal vstr_value As String) As String
      Return "<input type=""text"" id=""{field}"" name=""{field}"" maxlength=""{maxlen}"" style=""width:100%;"" value=""" & Server.HtmlEncode(vstr_value) & """ />"
    End Function

  End Class
#End Region

#Region " HiddenColumn "
  Public Class HiddenColumn
    Inherits BaseColumn

    Private mstr_default As String

    Public Sub New()
    End Sub

    Public Sub New(ByVal vstr_default As String)
      mstr_default = vstr_default
    End Sub

    Public Overrides ReadOnly Property Hidden() As Boolean
      Get
        Return True
      End Get
    End Property

    Public Overrides Function Render(ByVal vstr_value As String) As String
      If vstr_value = "" Then vstr_value = mstr_default
      Return "<input type=""hidden"" id=""{field}"" name=""{field}"" value=""" & Server.HtmlEncode(vstr_value) & """ />"
    End Function
  End Class
#End Region

#Region " TextColumn "
  Public Class TextColumn
    Inherits BaseColumn

    Private mint_rows As Integer = 5

    Public Sub New()
    End Sub

    Public Sub New(ByVal vint_rows As Integer)
      mint_rows = vint_rows
    End Sub

    Public Property Rows() As Integer
      Get
        Return mint_rows
      End Get
      Set(ByVal Value As Integer)
        mint_rows = Value
      End Set
    End Property

    Public Overrides Function Save(ByVal vstr_value As String) As String
      Return convertUnicode(vstr_value)
    End Function

    Public Overrides Function Render(ByVal vstr_value As String) As String
      Return "\n<textarea style=""width:100%;"" rows=""" & mint_rows & """ id=""{field}"" name=""{field}"">\n" & Server.HtmlEncode(vstr_value) & "\n</textarea>\n"
    End Function
  End Class
#End Region

#Region " MaskedColumn "
  Public Class MaskedColumn
    Inherits BaseColumn

    Private mstr_mask As String
    Public Property Mask() As String
      Get
        Return mstr_mask
      End Get
      Set(ByVal Value As String)
        mstr_mask = Value
      End Set
    End Property

    Public Sub New()
    End Sub

    Public Sub New(ByVal vstr_mask As String)
      mstr_mask = vstr_mask
    End Sub

    Public Overrides Function Render(ByVal vstr_value As String) As String
      Return "<input type=""text"" id=""{field}"" name=""{field}"" mask=""" & mstr_mask & """ class=""masked"" style=""width:100%;"" value=""" & Server.HtmlEncode(vstr_value) & """ />"
    End Function

  End Class
#End Region

#Region " NumberColumn "
  Public Class NumberColumn
    Inherits MaskedColumn
    Public Overrides Function Render(ByVal vstr_value As String) As String
      Dim obj_col As New MaskedColumn
      obj_col.Mask = "*#"
      Dim int_value As Integer = NullSafe(vstr_value, 0)
      Return obj_col.Render(int_value)
    End Function

    Public Overrides Function Save(ByVal vstr_value As String) As String
      On Error Resume Next
      Return NullSafe(vstr_value, 0)
    End Function

  End Class
#End Region

#Region " BooleanColumn "
  Public Class BooleanColumn
    Inherits BaseColumn

    Public Overrides Function Render(ByVal vstr_value As String) As String
      On Error Resume Next
      Dim b_value As Boolean = Boolean.Parse(vstr_value)
      Return "<input type=""hidden"" id=""{field}"" class=""check"" name=""{field}"" value=""" & IIf(b_value, "1", "0") & """ />"
    End Function
  End Class
#End Region

#Region " ByteColumn "
  Public Class ByteColumn
    Inherits BaseColumn

    Public Overrides Function Render(ByVal vstr_value As String) As String
      On Error Resume Next
      Return "<input type=""hidden"" id=""{field}"" class=""check"" name=""{field}"" value=""" & IIf(vstr_value = "1", "1", "0") & """ />"
    End Function
  End Class
#End Region

#Region " EncryptedColumn "
  Public Class EncryptedColumn
    Inherits BaseColumn

    Public Overrides Function Render(ByVal vstr_value As String) As String
      vstr_value = Decrypt(vstr_value)
      Return "<input style=""width:100%;"" type=""text"" maxlength=""{maxlen}"" name=""{field}"" id=""{field}"" value=""" & Server.HtmlEncode(vstr_value) & """ />"
    End Function

    Public Overrides Function Save(ByVal vstr_value As String) As String
      Return Encrypt(vstr_value)
    End Function

  End Class
#End Region

#Region " PasswordColumn "
  Public Class PasswordColumn
    Inherits BaseColumn

    Public Overrides Function Render(ByVal vstr_value As String) As String
      On Error Resume Next
      vstr_value = Decrypt(vstr_value)
      Return "<input style=""width:100%;"" type=""password"" maxlength=""{maxlen}"" name=""{field}"" id=""{field}"" value=""" & Server.HtmlEncode(vstr_value) & """ />"
    End Function

    Public Overrides Function Save(ByVal vstr_value As String) As String
      On Error Resume Next
      Return Encrypt(vstr_value)
    End Function

  End Class
#End Region

#Region " FileColumn "
  Public Class FileColumn
    Inherits BaseColumn

    Public Overrides Function Render(ByVal vstr_value As String) As String
      Return "<input type=""text"" id=""{field}"" name=""{field}"" class=""file"" style=""width:95%;"" value=""" & Server.HtmlEncode(vstr_value) & """><input style=""width:5%;border-left:0px;"" type=""button"" value=""..."" onclick=""SelectFile('{field}');"" />"
    End Function
  End Class
#End Region

#Region " ReferenceFileColumn "
  Public Class ReferenceFileColumn
    Inherits BaseColumn

    Public Overrides Function Render(ByVal vstr_value As String) As String
      Dim b_set As Boolean = NullSafe(vstr_value, 0) = 1
      Dim obj_sb As New StringBuilder()
      obj_sb.Append("<input type=""file"" id=""{field}"" name=""{field}"" class=""reffile"" style=""width:" & IIf(b_set, 95, 60) & "%;"" />")
      If b_set Then
        Dim str_url As String = RawUrl
        str_url = IIf(str_url.IndexOf("?") > 0, "&", "?")
        obj_sb.Append("<a href=""" & str_url & "clear=1"">Clear</a>")
      End If

      Return "<input type=""text"" id=""{field}"" name=""{field}"" class=""file"" style=""width:95%;"" value=""" & Server.HtmlEncode(vstr_value) & """><input style=""width:5%;border-left:0px;"" type=""button"" value=""..."" onclick=""SelectFile('{field}');"" />"
    End Function

    Public Overrides Function Save(ByVal vstr_value As String) As String

    End Function

  End Class
#End Region

#Region " EditorColumn "
  Public Class EditorColumn
    Inherits TextColumn

    Public Overrides Function Save(ByVal vstr_value As String) As String
      Return convertUnicode(vstr_value)
    End Function

    Public Overrides Function Render(ByVal vstr_value As String) As String
      Return "\n<textarea style=""width:100%;"" class=""editor"" rows=""" & MyClass.Rows & """ id=""{field}"" name=""{field}"">\n" & Server.HtmlEncode(vstr_value) & "\n</textarea>\n"
    End Function

  End Class
#End Region

#Region " DateColumn "
  Public Class DateColumn
    Inherits InputColumn

    Public Overrides Function Render(ByVal vstr_value As String) As String
      Return "<input class=""date"" type=""text"" id=""{field}"" name=""{field}"" maxlength=""{maxlen}"" style=""width:100%;"" value=""" & Server.HtmlEncode(vstr_value) & """ />"
    End Function
  End Class
#End Region

#Region " DropColumn "
  Public Class DropColumn
    Inherits BaseColumn

    Private Class Item
      Public Name As String
      Public Value As String
    End Class

    Private mobj_items As New ArrayList

    Public Sub New()
    End Sub

    Public Sub New(ByVal vstr_values As String())
      Add(vstr_values)
    End Sub

    Public Sub Add(ByVal vstr_value As String)
      Dim obj_item As New Item
      obj_item.Name = vstr_value
      obj_item.Value = vstr_value
      mobj_items.Add(obj_item)
    End Sub

    Public Sub Add(ByVal vstr_values As String())
      Dim obj_item As Item
      For Each str_value As String In vstr_values
        obj_item = New Item
        obj_item.Name = str_value
        obj_item.Value = str_value
        mobj_items.Add(obj_item)
      Next
    End Sub

    Public Sub Add(ByVal vstr_name As String, ByVal vstr_value As String)
      Dim obj_item As New Item
      obj_item.Name = vstr_name
      obj_item.Value = vstr_value
      mobj_items.Add(obj_item)
    End Sub

    Public Overrides Function Render(ByVal vstr_value As String) As String
      Dim obj_sb As New System.Text.StringBuilder
      obj_sb.Append("<select id=""{field}"" name=""{field}"" style=""width:100%;"">" & vbNewLine)
      For Each obj_item As Item In mobj_items
        obj_sb.Append("<option")
        If obj_item.Value = vstr_value Then obj_sb.Append(" selected")
        obj_sb.Append(" value=""" & obj_item.Value & """")
        obj_sb.Append(">" & obj_item.Name)
        obj_sb.Append("</option>" & vbNewLine)
      Next
      obj_sb.Append("</select>" & vbNewLine)
      Return obj_sb.ToString
    End Function
  End Class
#End Region

#Region " SQLDropColumn "
  Public Class SQLDropColumn
    Inherits BaseColumn
    Private mstr_sql As String
    Private mstr_dname As String
    Private mstr_dvalue As String

    Public Property SQL() As String
      Get
        Return mstr_sql
      End Get
      Set(ByVal Value As String)
        mstr_sql = Value
      End Set
    End Property

    Public Sub New()
    End Sub

    Public Sub New(ByVal vstr_sql As String)
      mstr_sql = vstr_sql
    End Sub

    Public Sub SetDefault(ByVal vstr_name As String, ByVal vstr_value As String)
      mstr_dname = vstr_name
      mstr_dvalue = vstr_value
    End Sub

    Public Overrides Function Render(ByVal vstr_value As String) As String
      Dim obj_sb As New System.Text.StringBuilder
      Dim str_sql As String = mstr_sql.Replace("{0}", vstr_value)
      str_sql = MyBase.ProcessFields(str_sql, New MatchEvaluator(AddressOf ProcessSQLFields))
      Dim obj_table As DataTable = Data.Table(str_sql)
      obj_sb.Append("<select id=""{field}"" name=""{field}"" style=""width:100%;"">" & vbNewLine)
      If mstr_dname <> "" Then
        obj_sb.Append("<option value=""" & mstr_dvalue & """")
        If mstr_dvalue = vstr_value Then obj_sb.Append(" selected")
        obj_sb.Append(">" & mstr_dname & "</option>" & vbNewLine)
      End If
      If Not obj_table Is Nothing Then
        For Each obj_row As DataRow In obj_table.Rows
          obj_sb.Append("<option")
          If NullSafe(obj_row(0)) = vstr_value Then obj_sb.Append(" selected")
          If obj_table.Columns.Count = 2 Then
            obj_sb.Append(" value=""" & obj_row(0) & """")
          End If
          obj_sb.Append(">" & obj_row(obj_table.Columns.Count - 1))
          obj_sb.Append("</option>" & vbNewLine)
        Next
      End If
      obj_sb.Append("</select>" & vbNewLine)
      Return obj_sb.ToString
    End Function
  End Class
#End Region

#Region " MultiSelColumn "
  Public Class MultiSelColumn
    Inherits BaseColumn

    Public Shared Sub SaveData(ByVal vstr_table As String, ByVal vstr_ref As String, ByVal vstr_field As String, ByVal vint_id As Integer, ByVal vstr_name As String)
      On Error Resume Next
      Dim obj_vals As New List(Of String)(GetForm(vstr_name).Split(","))
      Global.Data.Current.Execute("delete from [" & vstr_table & "] where [" & vstr_ref & "]=" & vint_id)
      Dim obj_query As QueryBuilder = Nothing
      For Each str_val As String In obj_vals
        obj_query = New QueryBuilder(vstr_table, QueryType.Insert)
        obj_query.Add(vstr_ref, vint_id)
        obj_query.Add(vstr_field, str_val)
        Global.Data.Current.Execute(obj_query)
      Next
    End Sub

    Private mstr_display As String
    Private mstr_data As String

    Public Sub New(ByVal vstr_display As String, ByVal vstr_data As String)
      mstr_display = vstr_display
      mstr_data = vstr_data
    End Sub

    Public Overrides Function Render(ByVal vstr_value As String) As String
      On Error Resume Next
      Dim obj_vals As New List(Of String)
      For Each obj_row As DataRow In Data.Table(mstr_data).Rows
        obj_vals.Add("" & obj_row(0))
      Next
      Dim obj_sb As New System.Text.StringBuilder
      obj_sb.Append("<select multiple=""multiple"" size=""5"" id=""{field}"" name=""{field}"" style=""width:100%;"">" & vbNewLine)
      For Each obj_row As DataRow In Data.Table(mstr_display).Rows
        obj_sb.Append("<option")
        If obj_vals.Contains("" & obj_row(0)) Then obj_sb.Append(" selected=""selected""")
        obj_sb.Append(" value=""" & obj_row(0) & """")
        obj_sb.Append(">" & obj_row(1))
        obj_sb.Append("</option>" & vbNewLine)
      Next
      obj_sb.Append("</select>" & vbNewLine)
      Return obj_sb.ToString
    End Function

    Public Overrides Function Save(ByVal vstr_value As String) As String
      Return String.Empty
    End Function
  End Class
#End Region

#Region " HTMLColumn "
  Public Class HTMLColumn
    Inherits BaseColumn

    Private mstr_html As String
    Public Property HTML() As String
      Get
        Return mstr_html
      End Get
      Set(ByVal Value As String)
        mstr_html = Value
      End Set
    End Property

    Public Sub New()
    End Sub

    Public Sub New(ByVal vstr_html As String)
      mstr_html = vstr_html
    End Sub

    Public Overrides Function Render(ByVal vstr_value As String) As String
      Return MyBase.ProcessFields(mstr_html.Replace("{0}", vstr_value))
    End Function
  End Class
#End Region

#Region " CustomColumn "
  Public Class CustomColumn
    Inherits BaseColumn
    Public Delegate Function Action(ByVal vstr_action As String, ByRef vstr_value As String, ByVal vobj_col As IColumn) As String
    Private mdel_call As Action

    Public Property Callback() As Action
      Get
        Return mdel_call
      End Get
      Set(ByVal Value As Action)
        mdel_call = Value
      End Set
    End Property

    Private mstr_scripts As String
    Public Property CustomScripts() As String
      Get
        Return mstr_scripts
      End Get
      Set(ByVal Value As String)
        mstr_scripts = Value
      End Set
    End Property

    Public Overrides ReadOnly Property Scripts() As String
      Get
        Return mstr_scripts
      End Get
    End Property

    Public Sub New()
    End Sub

    Public Sub New(ByVal vdel_callback As Action)
      mdel_call = vdel_callback
    End Sub

    Public Overrides Function Render(ByVal vstr_value As String) As String
      If Not mdel_call Is Nothing Then
        Return mdel_call.Invoke("render", vstr_value, Me)
      End If
      Return MyBase.ProcessFields(vstr_value)
    End Function

    Public Overrides Function Save(ByVal vstr_value As String) As String
      If Not mdel_call Is Nothing Then
        Return mdel_call.Invoke("save", vstr_value, Me) = "true"
      End If
      Return MyBase.Save(vstr_value)
    End Function
  End Class
#End Region

#End Region

  Protected mstr_table As String
  Protected mb_manual As Boolean
  Protected mstr_template As String
  Protected mstr_redirect As String
  Protected mstr_title As String = ""
  Protected mobj_names As New ArrayList
  Protected mobj_columns As New Hashtable
  Protected mobj_skip As New Hashtable
  Protected mobj_buttons As New Hashtable
  Protected mobj_data As Data.Engine
  Private mobj_scripts As New List(Of String)

  Protected mstr_color As String = "#FFFFFF"

  Public Delegate Function Save(ByVal vint_id As Integer) As String

  Protected mobj_save As Save

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
  Public Property Table() As String
    Get
      Return mstr_table
    End Get
    Set(ByVal Value As String)
      mstr_table = Value
    End Set
  End Property
  Public Property Redirect() As String
    Get
      Return mstr_redirect
    End Get
    Set(ByVal Value As String)
      mstr_redirect = Value
    End Set
  End Property
  Public Property CallBack() As Save
    Get
      Return mobj_save
    End Get
    Set(ByVal Value As Save)
      mobj_save = Value
    End Set
  End Property
  Public Property Manual() As Boolean
    Get
      Return mb_manual
    End Get
    Set(ByVal Value As Boolean)
      mb_manual = Value
    End Set
  End Property
  Public Property Template() As String
    Get
      Return mstr_template
    End Get
    Set(ByVal Value As String)
      mstr_template = Value
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
  Public Property Scripts() As List(Of String)
    Get
      Return mobj_scripts
    End Get
    Set(ByVal Value As List(Of String))
      mobj_scripts = Value
    End Set
  End Property
#End Region

#Region " New "
  Public Sub New()
  End Sub

  Public Sub New(ByVal vobj_call As Save)
    mobj_save = vobj_call
  End Sub

  Public Sub New(ByVal vstr_table As String)
    mstr_table = vstr_table
  End Sub

  Public Sub New(ByVal vstr_table As String, ByVal vobj_call As Save)
    mobj_save = vobj_call
    mstr_table = vstr_table
  End Sub

  Public Sub New(ByVal vb_manual As Boolean)
    mb_manual = vb_manual
  End Sub

  Public Sub New(ByVal vb_manual As Boolean, ByVal vobj_call As Save)
    mobj_save = vobj_call
    mb_manual = vb_manual
  End Sub

  Public Sub New(ByVal vstr_table As String, ByVal vb_manual As Boolean)
    mstr_table = vstr_table
    mb_manual = vb_manual
  End Sub

  Public Sub New(ByVal vstr_table As String, ByVal vb_manual As Boolean, ByVal vobj_call As Save)
    mstr_table = vstr_table
    mb_manual = vb_manual
    mobj_save = vobj_call
  End Sub
#End Region

  Public Sub Clear()
    mobj_names.Clear()
    mobj_columns.Clear()
    mobj_skip.Clear()
  End Sub

  Public Sub Skip(ByVal vstr_name As String)
    mobj_skip.Add(vstr_name, String.Empty)
  End Sub

  Public Sub Skip(ByVal vstr_name As String, ByVal vstr_default As String)
    mobj_skip.Add(vstr_name, vstr_default)
  End Sub

  Public Sub AddButton(ByVal vstr_name As String, ByVal vstr_action As String)
    mobj_buttons.Add(vstr_name, vstr_action)
  End Sub

  Public Sub AddScript(ByVal vstr_file As String)
    On Error Resume Next
    If Not mobj_scripts.Contains(vstr_file) Then mobj_scripts.Add(vstr_file)
  End Sub

#Region " Add Methods "

  Public Function AddColumn(ByVal vstr_name As String, ByVal vobj_col As IColumn) As IColumn
    Return AddColumn(vstr_name, StrConv(vstr_name, VbStrConv.ProperCase), vobj_col)
  End Function

  Public Function AddColumn(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vobj_col As IColumn) As IColumn
    mobj_names.Add(vstr_name)
    vobj_col.Data = Data
    vobj_col.Name = vstr_name
    vobj_col.Title = vstr_title
    mobj_columns.Add(vstr_name, vobj_col)
    Return vobj_col
  End Function

  Public Function AddView(ByVal vstr_name As String) As ViewColumn
    Return AddColumn(vstr_name, New ViewColumn)
  End Function

  Public Function AddView(ByVal vstr_name As String, ByVal vstr_title As String) As ViewColumn
    Return AddColumn(vstr_name, vstr_title, New ViewColumn)
  End Function

  Public Function AddInput(ByVal vstr_name As String) As InputColumn
    Return AddColumn(vstr_name, New InputColumn)
  End Function

  Public Function AddInput(ByVal vstr_name As String, ByVal vstr_title As String) As InputColumn
    Return AddColumn(vstr_name, vstr_title, New InputColumn)
  End Function

  Public Function AddHidden(ByVal vstr_name As String) As HiddenColumn
    Return AddColumn(vstr_name, New HiddenColumn)
  End Function

  Public Function AddHidden(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vstr_default As String) As HiddenColumn
    Return AddColumn(vstr_name, vstr_title, New HiddenColumn(vstr_default))
  End Function

  Public Function AddHidden(ByVal vstr_name As String, ByVal vstr_default As String) As HiddenColumn
    Return AddColumn(vstr_name, New HiddenColumn(vstr_default))
  End Function

  Public Function AddText(ByVal vstr_name As String, Optional ByVal vint_rows As Integer = 5) As TextColumn
    Return AddColumn(vstr_name, New TextColumn(vint_rows))
  End Function

  Public Function AddText(ByVal vstr_name As String, ByVal vstr_title As String, Optional ByVal vint_rows As Integer = 5) As TextColumn
    Return AddColumn(vstr_name, vstr_title, New TextColumn(vint_rows))
  End Function

  Public Function AddMasked(ByVal vstr_name As String, ByVal vstr_title As String, Optional ByVal vstr_mask As String = "") As MaskedColumn
    Return AddColumn(vstr_name, vstr_title, New MaskedColumn(vstr_mask))
  End Function

  Public Function AddNumber(ByVal vstr_name As String, ByVal vstr_title As String) As NumberColumn
    Return AddColumn(vstr_name, vstr_title, New NumberColumn)
  End Function

  Public Function AddBoolean(ByVal vstr_name As String, ByVal vstr_title As String) As BooleanColumn
    Return AddColumn(vstr_name, vstr_title, New BooleanColumn)
  End Function

  Public Function AddByte(ByVal vstr_name As String, ByVal vstr_title As String) As ByteColumn
    Return AddColumn(vstr_name, vstr_title, New ByteColumn)
  End Function

  Public Function AddPassword(ByVal vstr_name As String, ByVal vstr_title As String) As PasswordColumn
    Return AddColumn(vstr_name, vstr_title, New PasswordColumn)
  End Function

  Public Function AddFile(ByVal vstr_name As String, ByVal vstr_title As String) As FileColumn
    Return AddColumn(vstr_name, vstr_title, New FileColumn)
  End Function

  Public Function AddEditor(ByVal vstr_name As String, ByVal vstr_title As String) As EditorColumn
    Return AddColumn(vstr_name, vstr_title, New EditorColumn)
  End Function

  Public Function AddDate(ByVal vstr_name As String, ByVal vstr_title As String) As DateColumn
    Return AddColumn(vstr_name, vstr_title, New DateColumn)
  End Function

  Public Function AddDrop(ByVal vstr_name As String, ByVal vstr_title As String) As DropColumn
    Return AddColumn(vstr_name, vstr_title, New DropColumn)
  End Function

  Public Function AddSQLDrop(ByVal vstr_name As String, ByVal vstr_title As String, Optional ByVal vstr_sql As String = "") As SQLDropColumn
    Return AddColumn(vstr_name, vstr_title, New SQLDropColumn(vstr_sql))
  End Function

  Public Function AddHTML(ByVal vstr_name As String, ByVal vstr_title As String, Optional ByVal vstr_html As String = "") As HTMLColumn
    Return AddColumn(vstr_name, vstr_title, New HTMLColumn(vstr_html))
  End Function

  Public Function AddCustom(ByVal vstr_name As String, ByVal vstr_title As String, ByVal vobj_call As CustomColumn.Action) As CustomColumn
    Return AddColumn(vstr_name, vstr_title, New CustomColumn(vobj_call))
  End Function

#End Region

  Public Function Render(ByVal vint_id As Integer) As String
    Return Render("id", vint_id)
  End Function

  Public Function Render(ByVal vstr_pk As String, ByVal vint_id As Integer) As String
    Dim obj_col As IColumn
    Dim str_name As String
    Dim str_msg As String = String.Empty
    Dim str_value As String
    Dim obj_row As DataRow = Nothing
    Dim obj_cols As DataColumnCollection
    Dim obj_column As DataColumn
    ' Dim obj_table As DataTable 
    Try
      obj_cols = Data.Schema(mstr_table)
      If Request.Form.Count > 0 Then
        Try
          Dim obj_query As New QueryBuilder(mstr_table)
          If vint_id = 0 Then
            obj_query.QueryType = QueryType.Insert
          Else
            obj_query.QueryType = QueryType.Update
            obj_query.Where("id", vint_id)
          End If
          Dim obj_fields As ArrayList = GetFormList(vstr_prefix:="field")
          If mb_manual Then
            For Each str_name In mobj_names
              obj_col = mobj_columns(str_name)
              obj_col.ID = vint_id
              If obj_fields.Contains("field_" & str_name) Then
                str_value = GetForm("field_" & str_name)
                If obj_cols.Contains(str_name) Then
                  str_value = obj_col.Save(str_value)
                  'str_value = convertUnicode(str_value)
                  obj_query.Add(str_name, str_value)
                Else
                  obj_col.Save(str_value)
                End If
              Else
                obj_col.Save(String.Empty)
              End If
            Next
          Else
            For Each str_name In obj_fields
              str_name = str_name.Replace("field_", "")
              If obj_cols.Contains(str_name) Then
                obj_column = obj_cols(str_name)
                str_value = GetForm("field_" & str_name)
                If mobj_names.Contains(str_name) Then
                  obj_col = mobj_columns(str_name)
                Else
                  obj_col = GetColumn(obj_column)
                  obj_col.Name = str_name
                End If
                obj_col.ID = vint_id
                str_value = obj_col.Save(str_value)
                'str_value = convertUnicode(str_value)
                obj_query.Add(str_name, str_value)
              Else
                If mobj_names.Contains(str_name) Then
                  obj_col = mobj_columns(str_name)
                  obj_col.ID = vint_id
                  obj_col.Save(GetForm("field_" & str_name))
                End If
              End If
            Next
          End If
          If GetQuery("debug", 0) = 1 Then Return obj_query.ToString
          If Data.Execute(obj_query) Then
            If vint_id = 0 Then
              vint_id = Data.Identity
            End If
            If vint_id > 0 Then
              Try
                If Not mobj_save Is Nothing Then
                  str_msg = mobj_save(vint_id)
                End If
                If str_msg = "" Then
                  If mstr_redirect <> "" Then
                    Dim str_url As String = String.Format(mstr_redirect, vint_id)
                    Response.Redirect(str_url, True)
                    Return String.Empty
                  Else
                    str_msg = "Saved"
                  End If
                End If
              Catch ex As Exception
                str_msg = "Unable to Save"
              End Try
            Else
              str_msg = "Unable to Save"
            End If
          Else
            str_msg = "Unable to Save"
            'If Not Data.LastError Is Nothing Then
            '    str_msg &= "<br>" & Data.LastError.Message
            'End If
          End If
          str_msg &= "<!--" & obj_query.ToString() & "-->"
        Catch ex As Exception
          str_msg = "Unable to save<!--" & Server.HtmlEncode(ex.ToString) & "-->"
        End Try
      End If

      Dim obj_scripts As New ArrayList
      Dim obj_hidden As New ArrayList
      Dim obj_tp As New HTMLTemplate
      obj_tp.Template = NullSafe(mstr_template, "/cms/templates/edit.htm")
      obj_tp.SetItem("message", str_msg)
      obj_tp.SetItem("msg", str_msg)
      If Not mstr_title = "" And obj_tp.Sections.ContainsKey("title") Then
        obj_tp.Sections("title").AddNew()
        obj_tp.Sections("title").SetItem("title", mstr_title)
      End If
      obj_tp.CurrentPart = "fields"

      If vint_id > 0 Then
        obj_row = Data.Row("SELECT * FROM [" & mstr_table & "] WHERE [" & vstr_pk & "]=" & vint_id)
        'obj_table = obj_row.Table
      End If

      If mb_manual Then
        For Each str_name In mobj_names
          obj_col = mobj_columns(str_name)
          obj_col.ID = vint_id
          obj_col.Row = obj_row
          Dim int_max As Integer
          If Not obj_col.Hidden Then
            ProcessScripts(obj_col, obj_scripts)
            If obj_cols.Contains(str_name) Then
              obj_column = obj_cols(str_name)
              int_max = obj_column.MaxLength
              If vint_id > 0 Then
                str_value = NullSafe(obj_row(str_name)) ', "" & obj_col.DefaultValue) ' "" & obj_column.DefaultValue)
              Else
                If obj_row Is Nothing Then
                  str_value = NullSafe(obj_col.DefaultValue, "")
                Else
                  str_value = NullSafe(obj_row(str_name), "" & NullSafe(obj_col.DefaultValue, ""))
                End If
              End If
            Else
              str_value = ""
            End If
            RenderColumn(obj_tp, obj_col, str_value, int_max)
          Else
            obj_hidden.Add(obj_col)
          End If
        Next
      Else
        For Each obj_column In obj_cols
          str_name = obj_column.ColumnName
          If Not mobj_skip.Contains(str_name) Then
            If mobj_names.Contains(str_name) Then
              obj_col = mobj_columns(str_name)
            Else
              'Throw New HttpException(str_name & obj_column.AutoIncrement)
              If obj_column.AutoIncrement Then
                obj_col = New ViewColumn
              Else 'default
                obj_col = GetColumn(obj_column)
              End If
              obj_col.Name = str_name
              obj_col.Title = StrConv(str_name, VbStrConv.ProperCase)
            End If
            obj_col.ID = vint_id
            obj_col.Row = obj_row
            If Not obj_col.Hidden Then
              ProcessScripts(obj_col, obj_scripts)
              If vint_id > 0 Then
                str_value = NullSafe(obj_row(str_name)) ', "" & obj_col.DefaultValue) ' "" & obj_column.DefaultValue)
              Else
                str_value = "" & obj_col.DefaultValue 'obj_column.DefaultValue
              End If
              RenderColumn(obj_tp, obj_col, str_value, obj_column.MaxLength)
            Else
              obj_hidden.Add(obj_col)
            End If
          End If
        Next
        For Each str_name In mobj_names
          If Not obj_cols.Contains(str_name) Then
            obj_col = mobj_columns(str_name)
            If Not obj_col.Hidden Then
              obj_col.Row = obj_row
              ProcessScripts(obj_col, obj_scripts)
              RenderColumn(obj_tp, obj_col, "", 0)
            End If
          End If
        Next
      End If

      If obj_tp.Sections.ContainsKey("hidden") Then
        For Each obj_col In obj_hidden
          obj_col.Row = obj_row
          ProcessScripts(obj_col, obj_scripts)
          If vint_id > 0 Then
            str_value = NullSafe(obj_row(obj_col.Name))
          Else
            str_value = "" & obj_col.DefaultValue
          End If
          RenderColumn(obj_tp.Sections("hidden"), obj_col, str_value, 0)
        Next
      End If

      obj_tp.CurrentPart = ""
      For Each str_script As String In mobj_scripts
        If IO.Path.GetExtension(str_script).ToLower.TrimStart(".") = "css" Then
          obj_tp("styles").AddItem("css", str_script)
        Else
          obj_tp("scripts").AddItem("file", str_script)
        End If
      Next
      For Each str_script As String In obj_scripts
        If IO.Path.GetExtension(str_script).ToLower.TrimStart(".") = "css" Then
          obj_tp("styles").AddItem("css", str_script)
        Else
          obj_tp("scripts").AddItem("file", str_script)
        End If
      Next
      Dim str_google As String = Config("google.key")
      If Not String.IsNullOrEmpty(str_google) Then
        obj_tp.AddNew()
        obj_tp.SetItem("file", "http://maps.google.com/maps?file=api&v=2&key=" & str_google)
      End If


      obj_tp.CurrentPart = "buttons"
      For Each str_name In mobj_buttons.Keys
        obj_tp.AddNew()
        obj_tp.SetItem("name", str_name)
        obj_tp.SetItem("action", mobj_buttons(str_name))
      Next

      Return obj_tp.Render
    Catch ex As Exception
      Return ex.ToString
    End Try
  End Function

  Protected Sub RenderColumn(ByRef vobj_tp As HTMLTemplate, ByRef vobj_col As IColumn, ByVal vstr_value As String, ByVal vint_maxlen As Integer)
    On Error Resume Next
    'If vstr_value = "" Then vstr_value = "" & vobj_col.Column.DefaultValue
    Dim str_value As String = vobj_col.Render(vstr_value).Replace("\n", vbNewLine)
    str_value = str_value.Replace("{value}", Server.HtmlEncode(vstr_value))
    str_value = str_value.Replace("{name}", vobj_col.Name)
    If vint_maxlen <= 0 Then vint_maxlen = 255
    str_value = str_value.Replace("{maxlen}", vint_maxlen)
    str_value = str_value.Replace("{field}", "field_" & vobj_col.Name)
    mstr_color = IIf(mstr_color = "#FFFFFF", "#f7f7f7", "#FFFFFF")
    vobj_tp.AddNew()
    If vobj_col.Title.ToLower = "id" Then vobj_col.Title = "ID"
    vobj_tp.SetItem("title", vobj_col.Title)
    vobj_tp.SetItem("name", vobj_col.Name)
    vobj_tp.SetItem("color", mstr_color)
    vobj_tp.SetItem("value", str_value)
  End Sub

  Protected Sub ProcessScripts(ByRef vobj_col As IColumn, ByRef vobj_scripts As ArrayList)
    On Error Resume Next
    Dim str_scripts As String = NullSafe(vobj_col.Scripts)
    If str_scripts.Length > 0 Then
      For Each str_script As String In str_scripts.TrimEnd(";").Split(";")
        If Not vobj_scripts.Contains(str_script) Then
          vobj_scripts.Add(str_script)
        End If
      Next
    End If
  End Sub

  Protected Function GetColumn(ByVal vobj_column As DataColumn) As IColumn
    If vobj_column.ReadOnly Then
      Return New ViewColumn
    ElseIf vobj_column.AutoIncrement Then
      Return New ViewColumn(True)
    Else
      Select Case vobj_column.DataType.ToString.ToLower.Replace("system.", "")
        Case "byte" : Return New ByteColumn
        Case "boolean" : Return New BooleanColumn
        Case "integer", "int16", "int32", "int64", "decimal", "int", "double"
          Return New NumberColumn
        Case "datetime" : Return New DateColumn
        Case Else
          If vobj_column.MaxLength < 0 Then vobj_column.MaxLength = 255
          If vobj_column.MaxLength = 123 Then
            Return New PasswordColumn
          ElseIf vobj_column.MaxLength = 222 Then
            Return New FileColumn
          ElseIf vobj_column.MaxLength > 500 Then
            Return New EditorColumn
          ElseIf vobj_column.MaxLength > 255 Then
            Return New TextColumn
          Else
            Return New InputColumn
          End If
      End Select
    End If
  End Function


  Protected Shared Function convertUnicode(ByVal vstr_value As String) As String
    vstr_value = "" & (vstr_value)
    If vstr_value.Length > 0 Then
      Dim str_return As String = ""
      For i As Integer = 0 To vstr_value.Length - 1
        If AscW(vstr_value.Chars(i)) > 127 Then
          str_return &= "&#" & AscW(vstr_value.Chars(i)) & ";"
        Else
          str_return &= vstr_value.Chars(i)
        End If
      Next
      Return str_return
    End If
    Return String.Empty
  End Function


End Class