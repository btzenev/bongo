Public Module _Converters

#Region " ToString "

  Public Function ToStr(ByVal vobj_value As Object) As String
    Return ToString(vobj_value, String.Empty)
  End Function

  Public Function ToStr(ByVal vobj_value As Object, ByVal vstr_default As String) As String
    Return ToString(vobj_value, vstr_default)
  End Function

  Public Function ToString(ByVal vobj_value As Object) As String
    Return ToString(vobj_value, String.Empty)
  End Function

  Public Function ToString(ByVal vobj_value As Object, ByVal vstr_default As String) As String
    On Error Resume Next
    If vobj_value IsNot Nothing And vobj_value IsNot DBNull.Value Then
      Dim str_value As String
      If vobj_value.GetType Is GetType(String) Then
        str_value = DirectCast(vobj_value, String)
      Else
        str_value = Convert.ToString(vobj_value)
      End If
      If Not String.IsNullOrEmpty(str_value) Then Return str_value
    End If
    Return vstr_default
  End Function

#End Region

#Region " ToInteger "

  Public Function ToInt(ByVal vobj_value As Object) As Integer
    Return ToInteger(vobj_value, 0)
  End Function

  Public Function ToInt(ByVal vobj_value As Object, ByVal vint_default As Integer) As Integer
    Return ToInteger(vobj_value, vint_default)
  End Function

  Public Function ToInteger(ByVal vobj_value As Object) As Integer
    Return ToInteger(vobj_value, 0)
  End Function

  Public Function ToInteger(ByVal vobj_value As Object, ByVal vint_default As Integer) As Integer
    On Error Resume Next
    If vobj_value IsNot Nothing And vobj_value IsNot DBNull.Value Then
      If vobj_value.GetType Is GetType(Integer) Then
        Return DirectCast(vobj_value, Integer)
      Else
        Dim int_value As Integer
        If Integer.TryParse(vobj_value, int_value) Then Return int_value
      End If
    End If
    Return vint_default
  End Function

#End Region

#Region " ToDouble "

  Public Function ToDbl(ByVal vobj_value As Object) As Double
    Return ToDouble(vobj_value, 0.0)
  End Function

  Public Function ToDbl(ByVal vobj_value As Object, ByVal vdbl_default As Double) As Double
    Return ToDouble(vobj_value, vdbl_default)
  End Function

  Public Function ToDouble(ByVal vobj_value As Object) As Double
    Return ToDouble(vobj_value, 0.0)
  End Function

  Public Function ToDouble(ByVal vobj_value As Object, ByVal vdbl_default As Double) As Double
    On Error Resume Next
    If vobj_value IsNot Nothing And vobj_value IsNot DBNull.Value Then
      If vobj_value.GetType Is GetType(Double) Then
        Return DirectCast(vobj_value, Double)
      Else
        Dim dbl_value As Double
        If Double.TryParse(vobj_value, dbl_value) Then Return dbl_value
      End If
    End If
    Return vdbl_default
  End Function

#End Region

#Region " ToBoolean "

  Public Function ToBool(ByVal vobj_value As Object) As Boolean
    Return ToBoolean(vobj_value, False)
  End Function

  Public Function ToBool(ByVal vobj_value As Object, ByVal vb_default As Boolean) As Boolean
    Return ToBoolean(vobj_value, vb_default)
  End Function

  Public Function ToBoolean(ByVal vobj_value As Object) As Boolean
    Return ToBoolean(vobj_value, False)
  End Function

  Public Function ToBoolean(ByVal vobj_value As Object, ByVal vb_default As Boolean) As Boolean
        On Error Resume Next
    If vobj_value IsNot Nothing And vobj_value IsNot DBNull.Value Then
      If vobj_value.GetType Is GetType(Boolean) Then
                Return DirectCast(vobj_value, Boolean)
            ElseIf vobj_value.GetType Is GetType(Integer) Then
                Return ToInt(vobj_value, 0) = 1
            Else
                Dim b_value As Boolean
                If Boolean.TryParse(vobj_value, b_value) Then Return b_value
            End If
    End If
    Return vb_default
  End Function

#End Region

#Region " ToDate "

  Public Function ToDate(ByVal vobj_value As Object) As Date
    Return ToDate(vobj_value, Nothing)
  End Function

  Public Function ToDate(ByVal vobj_value As Object, ByVal vdat_default As Date) As Date
    On Error Resume Next
    If vobj_value IsNot Nothing And vobj_value IsNot DBNull.Value Then
      If vobj_value.GetType Is GetType(Date) Then
        Return DirectCast(vobj_value, Date)
      Else
        Dim dat_value As Date
        If Date.TryParse(vobj_value, dat_value) Then Return dat_value
      End If
    End If
    Return vdat_default
  End Function

#End Region

#Region " ToObject "

  Public Function ToObj(ByVal vobj_value As Object) As Object
    Return ToObject(vobj_value, Nothing)
  End Function

  Public Function ToObj(Of T)(ByVal vobj_value As Object) As T
    Return ToObject(Of T)(vobj_value, Nothing)
  End Function

  Public Function ToObj(ByVal vobj_value As Object, ByVal vobj_default As Object) As Object
    Return ToObject(vobj_value, vobj_default)
  End Function

  Public Function ToObj(Of T)(ByVal vobj_value As Object, ByVal vobj_default As T) As T
    Return ToObject(Of T)(vobj_value, vobj_default)
  End Function

  Public Function ToObject(ByVal vobj_value As Object) As Object
    Return ToObject(vobj_value, Nothing)
  End Function

  Public Function ToObject(Of T)(ByVal vobj_value As Object) As T
    Return ToObject(Of T)(vobj_value, Nothing)
  End Function

  Public Function ToObject(ByVal vobj_value As Object, ByVal vobj_default As Object) As Object
    Try
      If vobj_value IsNot Nothing And vobj_value IsNot DBNull.Value Then
        Return vobj_value
      End If
    Catch : End Try
    Return vobj_default
  End Function

  Public Function ToObject(Of T)(ByVal vobj_value As Object, ByVal vobj_default As T) As T
    Try
      If vobj_value IsNot Nothing And vobj_value IsNot DBNull.Value Then
        Return DirectCast(vobj_value, T)
      End If
    Catch : End Try
    Return vobj_default
  End Function

#End Region

End Module

