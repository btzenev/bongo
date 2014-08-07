Imports System.Text
Imports System.Text.RegularExpressions

Namespace Data

  Public Class QueryBuilder

#Region " Declarations "



    Public Class CustomValues
      Public Fields As String
      Public Where As String
      Public Order As String
    End Class

    Private mobj_type As QueryType = QueryType.Select
    Private mstr_table As String
    Private mobj_fields As New Dictionary(Of String, Object)
    Private mobj_fields_sort As New Dictionary(Of String, Sort)
    Private mobj_fields_types As New Dictionary(Of String, DataType)
    Private mobj_fields_sizes As New Dictionary(Of String, Integer)
    Private mobj_fields_nullable As New Dictionary(Of String, Boolean)
    Private mobj_fields_order As New List(Of String)
    Private mobj_wheres As New Dictionary(Of String, Object)
    Private mobj_wheres_compare As New Dictionary(Of String, Comparison)
    Private mobj_wheres_logic As New Dictionary(Of String, Logic)
    Private mobj_custom As New CustomValues

#End Region

#Region " Properties "

    Public Property QueryType() As QueryType
      Get
        Return mobj_type
      End Get
      Set(ByVal value As QueryType)
        mobj_type = value
      End Set
    End Property

    Public Property Table() As String
      Get
        Return mstr_table
      End Get
      Set(ByVal value As String)
        mstr_table = value
      End Set
    End Property

    Public ReadOnly Property Custom() As CustomValues
      Get
        Return mobj_custom
      End Get
    End Property

#End Region

#Region " New "

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal vstr_table As String)
      MyClass.New()
      mstr_table = vstr_table
    End Sub

    Public Sub New(ByVal vstr_table As String, ByVal vobj_type As QueryType)
      MyClass.New(vstr_table)
      mobj_type = vobj_type
    End Sub

    Public Sub New(ByVal vstr_table As String, ByVal ParamArray vstr_fields() As String)
      MyClass.New(vstr_table, QueryType.Select)
      For Each str_field As String In vstr_fields
        Add(str_field)
      Next
    End Sub

#End Region

#Region " Clear "

    Public Sub Clear()
      Custom.Fields = String.Empty
      Custom.Order = String.Empty
      mobj_fields.Clear()
      mobj_fields_order.Clear()
      mobj_fields_sort.Clear()
      Custom.Where = String.Empty
      mobj_wheres.Clear()
      mobj_wheres_compare.Clear()
      mobj_wheres_logic.Clear()
    End Sub

    Public Sub ClearAll()
      Clear()
      mstr_table = String.Empty
      mobj_type = QueryType.Select
    End Sub

#End Region

#Region " Field Methods "

#Region " Add "

    Public Sub AddList(ByVal ParamArray vstr_fields() As String)
      For Each str_field As String In vstr_fields
        Add(str_field)
      Next
    End Sub

    Public Sub AddList(ByVal vobj_fields As IDictionary)
      For Each vstr_key As String In vobj_fields.Keys
        Add(vstr_key, vobj_fields(vstr_key))
      Next
    End Sub

    Public Sub Add(ByVal vstr_field As String)
      Add(vstr_field, String.Empty)
    End Sub

    Public Sub Add(ByVal vstr_field As String, ByVal vobj_value As Object)
      vstr_field = vstr_field.ToLower.Trim("["c, "]"c)
      If mobj_fields.ContainsKey(vstr_field) Then
        mobj_fields(vstr_field) = vobj_value
      Else
        mobj_fields.Add(vstr_field, vobj_value)
      End If
    End Sub

#End Region

#Region " AddForm "

    Public Sub AddFormList(ByVal ParamArray vstr_fields() As String)
      For i As Integer = 0 To vstr_fields.Length - 1
        AddForm(vstr_fields(i))
      Next
    End Sub

    Public Sub AddFormList(ByVal vobj_fields As IDictionary)
      For Each vstr_key As String In vobj_fields.Keys
        AddForm(vstr_key, vobj_fields(vstr_key))
      Next
    End Sub

    Public Sub AddForm(ByVal vstr_field As String)
      AddForm(vstr_field, String.Empty)
    End Sub

    Public Sub AddForm(ByVal vstr_field As String, ByVal vobj_default As String)
      Add(vstr_field, GetForm(vstr_field, vobj_default))
    End Sub

    Public Sub AddForm(ByVal vstr_field As String, ByVal vobj_default As Integer)
      Add(vstr_field, GetForm(vstr_field, vobj_default))
    End Sub

    Public Sub AddForm(ByVal vstr_field As String, ByVal vobj_default As Date)
      Add(vstr_field, GetForm(vstr_field, vobj_default))
    End Sub

    Public Sub AddForm(ByVal vstr_field As String, ByVal vobj_default As Boolean)
      Add(vstr_field, GetForm(vstr_field, vobj_default))
    End Sub

    Public Sub AddForm(ByVal vstr_field As String, ByVal vobj_default As Object)
      Add(vstr_field, GetForm(vstr_field, vobj_default))
    End Sub

#End Region

#Region " AddQuery "

    Public Sub AddQueryList(ByVal ParamArray vstr_fields() As String)
      For Each str_field As String In vstr_fields
        AddQuery(str_field)
      Next
    End Sub

    Public Sub AddQueryList(ByVal vobj_fields As IDictionary)
      For Each vstr_key As String In vobj_fields.Keys
        AddQuery(vstr_key, vobj_fields(vstr_key))
      Next
    End Sub

    Public Sub AddQuery(ByVal vstr_field As String)
      AddQuery(vstr_field, String.Empty)
    End Sub

    Public Sub AddQuery(ByVal vstr_field As String, ByVal vobj_default As String)
      Add(vstr_field, GetQuery(vstr_field, vobj_default))
    End Sub

    Public Sub AddQuery(ByVal vstr_field As String, ByVal vobj_default As Integer)
      Add(vstr_field, GetQuery(vstr_field, vobj_default))
    End Sub

    Public Sub AddQuery(ByVal vstr_field As String, ByVal vobj_default As Date)
      Add(vstr_field, GetQuery(vstr_field, vobj_default))
    End Sub

    Public Sub AddQuery(ByVal vstr_field As String, ByVal vobj_default As Boolean)
      Add(vstr_field, GetQuery(vstr_field, vobj_default))
    End Sub

    Public Sub AddQuery(ByVal vstr_field As String, ByVal vobj_default As Object)
      Add(vstr_field, GetQuery(vstr_field, vobj_default))
    End Sub

#End Region

#Region " Remove "

    Public Sub Remove(ByVal vstr_field As String)
      vstr_field = vstr_field.ToLower.Trim("["c, "]"c)
      If mobj_fields.ContainsKey(vstr_field) Then
        mobj_fields.Remove(vstr_field)
      End If
    End Sub

#End Region

#End Region

#Region " Order Methods "

    Public Sub Order(ByVal vstr_field As String)
      Order(vstr_field, Sort.Asc)
    End Sub

    Public Sub Order(ByVal vstr_field As String, ByVal vobj_sort As Sort)
      vstr_field = vstr_field.ToLower.Trim("["c, "]"c)
      mobj_fields_order.Add(vstr_field)
      mobj_fields_sort.Add(vstr_field, vobj_sort)
    End Sub

#End Region

#Region " Where Methods "

#Region " Where "

    Public Sub Where(ByVal vstr_field As String, ByVal vobj_value As Object)
      Where(vstr_field, Comparison.Equals, vobj_value)
    End Sub

    Public Sub Where(ByVal vstr_field As String, ByVal vobj_compare As Comparison, ByVal vobj_value As Object)
      Where(vstr_field, Logic.And, vobj_compare, vobj_value)
    End Sub

    Public Sub Where(ByVal vstr_field As String, ByVal vobj_logic As Logic, ByVal vobj_value As Object)
      Where(vstr_field, vobj_logic, Comparison.Equals, vobj_value)
    End Sub

    Public Sub Where(ByVal vstr_field As String, ByVal vobj_logic As Logic, ByVal vobj_compare As Comparison, ByVal vobj_value As Object)
      vstr_field = vstr_field.ToLower.Trim("["c, "]"c)
      mobj_wheres.Add(vstr_field, vobj_value)
      mobj_wheres_logic.Add(vstr_field, vobj_logic)
      mobj_wheres_compare.Add(vstr_field, vobj_compare)
    End Sub

#End Region

#Region " RemoveWhere "

    Public Sub RemoveWhere(ByVal vstr_field As String)
      vstr_field = vstr_field.ToLower.Trim("["c, "]"c)
      If mobj_wheres.ContainsKey(vstr_field) Then
        mobj_wheres.Remove(vstr_field)
        mobj_wheres_logic.Remove(vstr_field)
        mobj_wheres_compare.Remove(vstr_field)
      End If
    End Sub

#End Region

#End Region

#Region " Generator Methods "

#Region " GenerateWhere "

    Private Sub GenerateWhere(ByRef vobj_sb As StringBuilder)
      If Not String.IsNullOrEmpty(Custom.Where) Then
        vobj_sb.Append(" WHERE " & Custom.Where)
      ElseIf mobj_wheres.Count > 0 Then
        vobj_sb.Append(" WHERE ")
        Dim int_count As Integer
        For Each str_where As String In mobj_wheres.Keys
          int_count += 1
          If int_count > 1 Then
            If mobj_wheres_logic(str_where) = Logic.Or Then
              vobj_sb.Append(" OR ")
            Else
              vobj_sb.Append(" AND ")
            End If
          End If
          vobj_sb.Append("[" & str_where & "]")
          Dim str_format As String
          Select Case mobj_wheres_compare(str_where)
            Case Comparison.NotEqual : str_format = "<>{0}"
            Case Comparison.GreaterThan : str_format = ">{0}"
            Case Comparison.GreaterThankEqual : str_format = ">={0}"
            Case Comparison.LessThan : str_format = "<{0}"
            Case Comparison.LessThanEqual : str_format = "<={0}"
            Case Comparison.Like : str_format = " LIKE '%{0}%'"
            Case Else : str_format = "={0}"
          End Select
          vobj_sb.AppendFormat(str_format, "@WhErE_" & str_where)
        Next
      Else
        If mobj_type = QueryType.Update Or mobj_type = QueryType.Delete Then
          Exception("QueryBuilder", mobj_type.ToString & " requires where clause declaration")
        End If
      End If
    End Sub

#End Region

#Region " GenerateOrder "

    Private Sub GenerateOrder(ByRef vobj_sb As StringBuilder)
      If Not String.IsNullOrEmpty(Custom.Order) Then
        vobj_sb.Append(" ORDER BY " & Custom.Order)
      ElseIf mobj_fields_order.Count > 0 Then
        vobj_sb.Append(" ORDER BY ")
        For Each str_order As String In mobj_fields_order
          vobj_sb.Append("[" & str_order & "]")
          If mobj_fields_sort(str_order) = Sort.Desc Then
            vobj_sb.Append(" DESC")
          Else
            vobj_sb.Append(" ASC")
          End If
          vobj_sb.Append(", ")
        Next
        vobj_sb.Remove(vobj_sb.Length - 2, 2)
      End If
    End Sub

#End Region

    Private Function GetDataType(ByVal vstr_field As String) As String
      Dim obj_type As DataType = mobj_fields_types(vstr_field)
      Select Case obj_type
        Case DataType.Identity
        Case DataType.DateTime : Return "DATETIME"
        Case DataType.Integer : Return "INT"
        Case DataType.Long : Return "BIGINT"
        Case DataType.String : Return "VARCHAR(" & mobj_fields_sizes(vstr_field) & ")"
        Case DataType.Byte : Return "BIT(8)"
      End Select

    End Function

#Region " GenerateSQL "

    Private Function GenerateSQL() As String
      Dim obj_sb As New Text.StringBuilder
      Select Case mobj_type
        Case QueryType.Drop
          obj_sb.Append("DROP TABLE [" & mstr_table & "]")
        Case QueryType.Create
          obj_sb.Append("CREATE TABLE [" & mstr_table & "] (")
          If mobj_fields.Count > 0 Then
            Dim str_type As String
            Dim str_null As String
            For Each str_field As String In mobj_fields.Keys
              str_type = GetDataType(str_field)
              str_null = IIf(mobj_fields_nullable(str_field), "NULL", "NOT NULL").ToString
              obj_sb.AppendFormat("  [{0,-31}]{1,-20}{2,8},", str_field, str_type, str_null)
            Next
          End If
          obj_sb.Append(")")
        Case QueryType.Delete
          obj_sb.Append("DELETE FROM [" & mstr_table & "]")
          GenerateWhere(obj_sb)
        Case QueryType.Insert
          obj_sb.Append("INSERT INTO [" & mstr_table & "] (")
          Dim obj_fields As New Text.StringBuilder
          Dim obj_values As New Text.StringBuilder
          If mobj_fields.Count > 0 Then
            For Each str_field As String In mobj_fields.Keys
              obj_fields.Append("[" & str_field & "], ")
              obj_values.Append("@FiElD_" & str_field & ", ")
            Next
            obj_fields.Remove(obj_fields.Length - 2, 2)
            obj_values.Remove(obj_values.Length - 2, 2)
          End If
          obj_sb.Append(obj_fields.ToString)
          obj_sb.Append(") VALUES (")
          obj_sb.Append(obj_values.ToString)
          obj_sb.Append(")")
        Case QueryType.Update
          obj_sb.Append("UPDATE [" & mstr_table & "] SET ")
          If mobj_fields.Count > 0 Then
            For Each str_field As String In mobj_fields.Keys
              obj_sb.Append("[" & str_field & "]=")
              obj_sb.Append("@FiElD_" & str_field & ", ")
            Next
            obj_sb.Remove(obj_sb.Length - 2, 2)
          End If
          GenerateWhere(obj_sb)
        Case QueryType.Count
          obj_sb.Append("SELECT COUNT(*) as [count] FROM [" & mstr_table & "]")
          GenerateWhere(obj_sb)
        Case Else 'query.Type.Select
          obj_sb.Append("SELECT ")
          If Not String.IsNullOrEmpty(Custom.Fields) Then
            obj_sb.Append(Custom.Fields)
          Else
            If mobj_fields.Count > 0 Then
              For Each str_field As String In mobj_fields.Keys
                obj_sb.Append("[" & str_field & "], ")
              Next
              obj_sb.Remove(obj_sb.Length - 2, 2)
            Else
              obj_sb.Append("*")
            End If
          End If
          obj_sb.Append(" FROM [" & mstr_table & "]")
          GenerateWhere(obj_sb)
          GenerateOrder(obj_sb)
      End Select
      Return obj_sb.ToString
    End Function

#End Region

#End Region

#Region " ToCommand "

    Public Function ToCommand() As DbCommand
      Return ToCommand(Data.Current)
    End Function

    Public Function ToCommand(ByVal vobj_engine As Engine) As DbCommand
      Dim obj_cmd As DbCommand = vobj_engine.GetCommand(GenerateSQL)
      Dim obj_param As DbParameter
      If Not mobj_type = QueryType.Delete And Not mobj_type = QueryType.Select Then
        For Each str_field As String In mobj_fields.Keys
          obj_param = vobj_engine.CreateParameter("FiElD_" & str_field, mobj_fields(str_field))
          obj_cmd.Parameters.Add(obj_param)
        Next
      End If
      If Not mobj_type = QueryType.Insert Then
        For Each str_where As String In mobj_wheres.Keys
          obj_param = vobj_engine.CreateParameter("WhErE_" & str_where, mobj_wheres(str_where))
          obj_cmd.Parameters.Add(obj_param)
        Next
      End If
      Return obj_cmd
    End Function

#End Region

#Region " ToString "

    Private Shared mreg_fields As New Regex("(\@\w*)", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    Public Shadows Function ToString() As String
      Dim str_sql As String = GenerateSQL()
      str_sql = mreg_fields.Replace(str_sql, AddressOf ProcessField)
      Return str_sql
    End Function

    Private Function ProcessField(ByVal vobj_match As Match) As String
      On Error Resume Next
      Dim str_key As String = vobj_match.Value.TrimStart("@"c)
      Dim b_field As Boolean = str_key.StartsWith("FiElD_")
      str_key = str_key.Substring(6)
      Dim obj_dic As IDictionary(Of String, Object) = IIf(b_field, mobj_fields, mobj_wheres)
      If obj_dic.ContainsKey(str_key) Then
        Dim str_value As String = ToStr(obj_dic(str_key)).Replace("'", "''")
        If Not IsEmpty(str_value) Then
          If Not IsNumeric(str_value) Then str_value = "'" & str_value & "'"
          Return str_value
        End If
      End If
      Return String.Empty
    End Function

#End Region

  End Class

End Namespace