Namespace Plugins.Hidden

    Public Class Root
        Inherits Plugins.Hidden.Base

    Public Overrides Function Render() As String
      On Error Resume Next
            Dim int_root As Integer = GetRoot(Page.ID, True)
            If int_root = Page.ID Then
                Return "<span class=""unlink"">" & Pages.Get(int_root).Title & "</span>"
            Else
                Return "<a href=""" & Pages.Get(int_root).Url & """>" & Pages.Get(int_root).Title & "</a>"
            End If
    End Function


    End Class

End Namespace