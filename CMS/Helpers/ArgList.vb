
Public Class ArgList
    Inherits List(Of String)

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal collection As System.Collections.Generic.IEnumerable(Of String))
        MyBase.New(collection)
    End Sub

    Public Sub New(ByVal vstr_url As String)
        MyBase.New()
        MyBase.AddRange(GetUrlArgs(vstr_url))
    End Sub

    Default Public Overloads Property Item(ByVal vint_index As Integer) As String
        Get
            On Error Resume Next
            If vint_index < MyBase.Count And vint_index >= 0 Then
                Return MyBase.Item(vint_index)
            Else
                Return String.Empty
            End If
        End Get
        Set(ByVal value As String)
            On Error Resume Next
            If vint_index < MyBase.Count And vint_index >= 0 Then
                MyBase.Item(vint_index) = value
            Else
                MyBase.Insert(vint_index, value)
            End If
        End Set
    End Property

    Public Function First(ByVal vint_num As Integer) As ArgList
        On Error Resume Next
        Dim obj_args As New ArgList()
        If vint_num > 0 And MyBase.Count > 0 Then
            If vint_num > MyBase.Count Then
                vint_num = MyBase.Count
            End If
            For i As Integer = 0 To (vint_num - 1)
                obj_args.Add(MyBase.Item(i))
            Next
        End If
        Return obj_args
    End Function

    Public Function Last() As String
        On Error Resume Next
        Return ToStr(MyBase.Item(MyBase.Count - 1))
    End Function

    Public Function Last(ByVal vint_num As Integer) As ArgList
        On Error Resume Next
        Dim obj_args As New ArgList()
        If vint_num > 0 And MyBase.Count > 0 Then
            If vint_num > MyBase.Count Then
                vint_num = MyBase.Count
            End If
            For i As Integer = (MyBase.Count - 1) To (MyBase.Count - vint_num - 1)
                obj_args.Add(MyBase.Item(i))
            Next
        End If
        Return obj_args
    End Function

End Class
