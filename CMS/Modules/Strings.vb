Public Module _Strings

#Region " Case "

  Public Function ToUpper(ByVal vstr_value As String) As String
    On Error Resume Next
    Return vstr_value.ToUpper
  End Function

  Public Function ToLower(ByVal vstr_value As String) As String
    On Error Resume Next
    Return vstr_value.ToLower
  End Function

  Public Function ToProper(ByVal vstr_value As String) As String
    On Error Resume Next
    Return StrConv(vstr_value, VbStrConv.ProperCase)
  End Function

#End Region

#Region " IsMatch "

  Public Class Matches
    Private Shared mobj_reg As RegexOptions = RegexOptions.IgnoreCase Or RegexOptions.Compiled
    Public Shared Guid As New Regex("^(\{){0,1}[0-9A-F]{8}\-[0-9A-F]{4}\-[0-9A-F]{4}\-[0-9A-F]{4}\-[0-9A-F]{12}(\}){0,1}$", mobj_reg)
    Public Shared IPAddress As New Regex("^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$", mobj_reg)
    Public Shared Url As New Regex("(http|ftp|https):\/\/[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&amp;:/~\+#]*[\w\-\@?^=%&amp;/~\+#])?", mobj_reg)
    Public Shared Email As New Regex("^.+@[^\.].*\.[a-z]{2,}$", mobj_reg)
    Public Shared Domain As New Regex("^([a-z0-9]([a-z0-9\-]{0,61}[a-z0-9])?\.)+[a-z]{2,6}$", mobj_reg)
    Public Shared Numeric As New Regex("^-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?$", mobj_reg)
    Public Shared [Boolean] As New Regex("^(true|1)$", mobj_reg)
    Public Shared Phone As New Regex("^(?:\([2-9]\d{2}\)\ ?|[2-9]\d{2}(?:\-?|\ ?))[2-9]\d{2}[- ]?\d{4}$", mobj_reg)
    Public Shared Postal As New Regex("^((\d{5}-\d{4})|(\d{5})|([a-z]\d[a-z]\s?\d[a-z]\d))$", mobj_reg)
  End Class

  Public Function IsMatch(ByVal vstr_value As String, ByVal vstr_exp As String) As Boolean
    On Error Resume Next
    Return IsMatch(vstr_value, New Regex(vstr_exp, RegexOptions.IgnoreCase))
  End Function

  Private Function IsMatch(ByVal vstr_value As String, ByRef vobj_reg As Regex) As Boolean
    On Error Resume Next
    If Not String.IsNullOrEmpty(vstr_value) Then
      Return vobj_reg.IsMatch(vstr_value)
    End If
  End Function

#End Region

#Region " Contains "

  Public Function Contains(ByVal vobj_value As String, ByVal ParamArray vobj_params() As String) As Boolean
    For Each obj_param As String In vobj_params
      If vobj_value.ToLower.Contains(obj_param.ToLower) Then
        Return True
      End If
    Next
  End Function

#End Region

#Region " StartsWith "

  Public Function StartsWith(ByVal vobj_value As String, ByVal ParamArray vobj_params() As String) As Boolean
    For Each obj_param As String In vobj_params
      If vobj_value.StartsWith(obj_param, True, Nothing) Then
        Return True
      End If
    Next
  End Function

#End Region

#Region " EndsWith "

  Public Function EndsWith(ByVal vobj_value As String, ByVal ParamArray vobj_params() As String) As Boolean
    For Each obj_param As String In vobj_params
      If vobj_value.EndsWith(obj_param, True, Nothing) Then
        Return True
      End If
    Next
  End Function

#End Region

#Region " Normalize "

  Public Function Normalize(ByVal vstr_value As String) As String
    On Error Resume Next
    Dim str_normal As String = vstr_value.Normalize(NormalizationForm.FormD)
    Dim obj_sb As New StringBuilder
    For Each chr As Char In str_normal
      If CharUnicodeInfo.GetUnicodeCategory(chr) <> UnicodeCategory.NonSpacingMark Then obj_sb.Append(chr)
    Next
    Return obj_sb.ToString.Normalize(NormalizationForm.FormC)
  End Function

#End Region

#Region " IsEmpty "

  Public Function IsNull(ByVal vobj_value As Object) As Boolean
    On Error Resume Next
    Return vobj_value Is DBNull.Value Or vobj_value Is Nothing
  End Function

  Public Function IsEmpty(ByVal vobj_value As Object) As Boolean
    On Error Resume Next
    Return IsEmpty(ToStr(vobj_value))
  End Function

  Public Function IsEmpty(ByVal vstr_value As String) As Boolean
    On Error Resume Next
    Return String.IsNullOrEmpty(vstr_value)
  End Function

  Public Function NotEmpty(ByVal vobj_value As Object) As Boolean
    On Error Resume Next
    Return Not IsEmpty(ToStr(vobj_value))
  End Function

  Public Function NotEmpty(ByVal vstr_value As String) As Boolean
    On Error Resume Next
    Return Not String.IsNullOrEmpty(vstr_value)
  End Function

#End Region

#Region " StrRepeat "

  Public Function StrRepeat(ByVal vstr_value As String, ByVal vint_count As Integer) As String
    On Error Resume Next
    Dim obj_sb As New StringBuilder
    For i As Integer = 1 To vint_count
      obj_sb.Append(vstr_value)
    Next
    Return obj_sb.ToString()
  End Function

#End Region

#Region " GetBetween "

  Public Function GetBetween(ByVal vstr_value As String, ByVal vstr_start As String, ByVal vstr_end As String) As String
    On Error Resume Next
    Dim int_start As Integer = vstr_value.IndexOf(vstr_start)
    If int_start >= 0 Then
      int_start += vstr_start.Length
      Dim int_end As Integer = vstr_value.IndexOf(vstr_end, int_start)
      If int_end > int_start Then
        Return vstr_value.Substring(int_start, int_end - int_start)
      End If
    End If
    Return String.Empty
  End Function

#End Region

#Region " StripWord "

  Private mreg_word_xml As New Regex("<\\?\?xml[^>]*>", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
  Private mreg_word_op As New Regex("<\/?o:p[^>]*>", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
  Private mreg_word_v As New Regex("<\/?v:[^>]*>", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
  Private mreg_word_o As New Regex("<\/?o:[^>]*>", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
  Private mreg_word_st1 As New Regex("<\/?st1:[^>]*>", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
  Public Function StripWord(ByVal vstr_html As String) As String
    vstr_html = mreg_word_xml.Replace(vstr_html, "")
    vstr_html = mreg_word_op.Replace(vstr_html, "")
    vstr_html = mreg_word_v.Replace(vstr_html, "")
    vstr_html = mreg_word_o.Replace(vstr_html, "")
    vstr_html = mreg_word_st1.Replace(vstr_html, "")
    Return vstr_html
  End Function

#End Region

#Region " Hash "

  Public Function GetCRC(ByVal vstr_data As String) As String
    Return GetHash(vstr_data, Encryption.Hash.Provider.CRC32)
  End Function

  Public Function GetHash(ByVal vstr_data As String, ByVal vobj_provider As Encryption.Hash.Provider) As String
    Return GetHash(vstr_data, String.Empty, vobj_provider)
  End Function

  Public Function GetHash(ByVal vstr_data As String, ByVal vstr_salt As String, ByVal vobj_provider As Encryption.Hash.Provider) As String
    Dim obj_enc As New Encryption.Hash(vobj_provider)
    Dim obj_data As New Encryption.Data(vstr_data)
    Dim obj_salt As New Encryption.Data(vstr_salt)
    obj_enc.Calculate(obj_data, obj_salt)
    Return obj_enc.Value.ToHex
  End Function

  Public Function GetFileCRC(ByVal vstr_file As String) As String
    Return GetFileHash(vstr_file, Encryption.Hash.Provider.CRC32)
  End Function

  Public Function GetFileHash(ByVal vstr_file As String, Optional ByVal vobj_provider As Encryption.Hash.Provider = Encryption.Hash.Provider.CRC32) As String
    vstr_file = MapPath(vstr_file)
    Dim obj_enc As New Encryption.Hash(vobj_provider)
    Using obj_stream As New IO.StreamReader(vstr_file)
      obj_enc.Calculate(obj_stream.BaseStream)
    End Using
    Return obj_enc.Value.ToHex
  End Function

#End Region

#Region " Encrypt/Decrypt "

  Public Function Encrypt(ByVal vstr_source As String) As String
    Return Encrypt(vstr_source, ToStr(Config("site.encryption"), AppID))
  End Function

  Public Function Encrypt(ByVal vstr_source As String, ByVal vstr_key As String) As String
    If Not String.IsNullOrEmpty(vstr_source) Then
      Dim obj_sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
      Dim obj_key As New Encryption.Data(vstr_key)
      Dim obj_dec As New Encryption.Data(vstr_source)
      Dim obj_enc As Encryption.Data = obj_sym.Encrypt(obj_dec, obj_key)
      Return obj_enc.ToHex.ToUpper
    End If
  End Function

  Public Function Decrypt(ByVal vstr_source As String) As String
    Return Decrypt(vstr_source, ToStr(Config("site.encryption"), AppID))
  End Function

  Public Function Decrypt(ByVal vstr_source As String, ByVal vstr_key As String) As String
    Dim obj_sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
    Dim obj_key As New Encryption.Data(vstr_key)
    Dim obj_enc As New Encryption.Data()
    obj_enc.Hex = vstr_source
    Return obj_sym.Decrypt(obj_enc, obj_key).ToString
  End Function

#End Region

End Module
