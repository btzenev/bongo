Imports System
Imports System.Collections
Imports System.Globalization
Imports System.Text

'Build Json string.

Public Enum JsonType
    [None] = 0
    [Object] = 1
    [Property] = 2
    [Array] = 3
End Enum

Public Class JsonBuilder

    Private mobj_sb As New Text.StringBuilder
    Private mobj_stack As New Stack(Of JsonType)
    Private mobj_start As New Stack(Of JsonType)

    Public Sub New(Optional ByVal vobj_type As JsonType = JsonType.Object)
        MyBase.New()
        mobj_stack.Push(JsonType.None)
        mobj_start.Push(vobj_type)
        Select Case vobj_type
            Case JsonType.Object : StartObject()
            Case JsonType.Array : StartArray()
        End Select
    End Sub

    Public Sub New(ByVal vstr_name As String, Optional ByVal vobj_type As JsonType = JsonType.Object)
        MyClass.New(JsonType.Object)
        StartProperty(vstr_name)
        mobj_start.Push(vobj_type)
        Select Case vobj_type
            Case JsonType.Object : StartObject()
            Case JsonType.Array : StartArray()
        End Select
    End Sub

    Public Overloads Overrides Function ToString() As String
        DeleteLastComma()

        For Each obj_type As JsonType In mobj_start
            Select Case obj_type
                Case JsonType.Object : EndObject()
                Case JsonType.Array : EndArray()
                Case JsonType.Property : EndObject() : EndProperty() : DeleteLastComma() : EndObject()
            End Select
        Next

        DeleteLastComma()
        Return mobj_sb.ToString
    End Function

    Public Sub StartProperty(ByVal vstr_name As String)
        mobj_stack.Push(True)
        mobj_sb.Append(PrepString(vstr_name) & ":")
    End Sub

    Public Sub EndProperty()
        mobj_stack.Pop()
    End Sub

    Public Sub StartObject()
        mobj_stack.Push(JsonType.Object)
        mobj_sb.Append("{")
    End Sub

    Public Sub EndObject()
        DeleteLastComma()
        mobj_stack.Pop()
        mobj_sb.Append("}")
        Comma()
    End Sub

    Public Sub StartArray()
        mobj_stack.Push(JsonType.Array)
        mobj_sb.Append("[")
    End Sub

    Public Sub EndArray()
        DeleteLastComma()
        mobj_stack.Pop()
        mobj_sb.Append("]")
        Comma()
    End Sub

    Private Sub DeleteLastComma()
        On Error Resume Next
        If mobj_sb.Length > 0 Then
            If mobj_sb.Chars(mobj_sb.Length - 1) = ","c Then mobj_sb.Remove(mobj_sb.Length - 1, 1)
        End If
    End Sub

    Private Sub Comma()
        On Error Resume Next
        DeleteLastComma()
        Dim obj_type As JsonType = mobj_stack.Peek
        If obj_type = Nothing Then Return
        If obj_type = JsonType.Property Then Return
        mobj_sb.Append(",")
    End Sub

    Public Sub Write(ByVal vstr_name As String, ByVal vobj_value As Object)
        StartProperty(vstr_name)
        mobj_sb.Append(PrepValue(vobj_value))
        EndProperty()
        Comma()
    End Sub

    Public Sub Write(ByVal vobj_value As Object)
        mobj_sb.Append(PrepValue(vobj_value))
        Comma()
    End Sub

    Public Sub WriteRaw(ByVal vstr_name As String, ByVal vobj_value As String, Optional ByVal vb_comma As Boolean = True)
        StartProperty(vstr_name)
        mobj_sb.Append(vobj_value)
        EndProperty()
        If vb_comma Then Comma()
    End Sub

    Public Sub WriteRaw(ByVal vobj_value As String, Optional ByVal vb_comma As Boolean = True)
        mobj_sb.Append(vobj_value)
        If vb_comma Then Comma()
    End Sub

    Private Function PrepString(ByVal str As String) As String
        str = StripHtml(str)
        Dim sb As New StringBuilder
        sb.Append(""""c)
        For Each c As Char In str
            Select Case c
                Case """"c
                    sb.Append("\""")
                Case "\"c
                    sb.Append("\\")
                Case ControlChars.Back
                    sb.Append("\b")
                Case ControlChars.FormFeed
                    sb.Append("\f")
                Case ControlChars.Lf
                    sb.Append("\n")
                Case ControlChars.Cr
                    sb.Append("\r")
                Case ControlChars.Tab
                    sb.Append("\t")
                Case Else
                    Dim i As Integer = AscW(c)
                    If i < 32 OrElse i > 127 Then
                        sb.AppendFormat("\u{0:X04}", i)
                    Else
                        sb.Append(c)
                    End If
            End Select
        Next
        sb.Append(""""c)
        Return sb.ToString
    End Function

    Private Function PrepValue(ByVal value As Object) As String
        Dim sb As New StringBuilder
        If value Is Nothing OrElse value Is DBNull.Value Then
            sb.Append("null")
        ElseIf TypeOf value Is String OrElse TypeOf value Is Char Then
            Return PrepString(value.ToString)
        ElseIf TypeOf value Is Double OrElse TypeOf value Is Single OrElse TypeOf value Is Long OrElse TypeOf value Is Integer OrElse TypeOf value Is Short OrElse TypeOf value Is Byte OrElse TypeOf value Is Decimal Then
            sb.Append(value)
        ElseIf TypeOf value Is DateTime Then
            Dim dt As DateTime = DirectCast(value, DateTime)
            sb.Append(""""c)
            sb.Append(dt.ToString)
            sb.Append(""""c)
        ElseIf TypeOf value Is Boolean Then
            sb.Append(CBool(value).ToString.ToLower)
        ElseIf TypeOf value Is Guid Then
            sb.Append(""""c)
            sb.Append(value.ToString())
            sb.Append(""""c)
        ElseIf TypeOf value Is [Enum] Then
            Dim underlyingType As System.Type = [Enum].GetUnderlyingType(value.[GetType]())
            sb.Append(Convert.ChangeType(value, underlyingType, CultureInfo.InvariantCulture))
            'ElseIf TypeOf value Is IJson Then
            '  sb.Append(TryCast(value, IJson).ToJson())
        Else
            Throw New NotImplementedException("Unknown JSON type.")
        End If
        Return sb.ToString
    End Function

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
            vstr_html = Regex.Replace(vstr_html, "&trade;", "(™)", mobj_opts)
            vstr_html = Regex.Replace(vstr_html, "&frasl;", "/", mobj_opts)
            vstr_html = Regex.Replace(vstr_html, "&lt;", "<", mobj_opts)
            vstr_html = Regex.Replace(vstr_html, "&gt;", ">", mobj_opts)
            vstr_html = Regex.Replace(vstr_html, "&copy;", "©", mobj_opts)
            vstr_html = Regex.Replace(vstr_html, "&reg;", "®", mobj_opts)
            vstr_html = Regex.Replace(vstr_html, "&(.{2,6});", String.Empty, mobj_opts)
            vstr_html = Regex.Replace(vstr_html, "\n+", vbNewLine, mobj_opts)
            vstr_html = Regex.Replace(vstr_html, "\s+", " ", mobj_opts)
        End If
        Return vstr_html
    End Function

End Class