Namespace Plugins.Hidden

    Public Class [Error]
        Inherits Plugins.Hidden.Base

        Private mstr_name As String = String.Empty
        Private mobj_ex As System.Exception
        Public Sub New(ByVal vstr_name As String, ByRef vobj_ex As System.Exception)
            mstr_name = vstr_name
            mobj_ex = vobj_ex
        End Sub

        Public Overrides Function Render() As String
            Dim obj_sb As New StringBuilder
            obj_sb.AppendLine("[" & mstr_name & " is not available to this time]")
            If mobj_ex IsNot Nothing Then
                obj_sb.AppendLine("<!--")
                obj_sb.AppendLine(HTMLEncode(mobj_ex.ToString))
                obj_sb.AppendLine("-->")
            End If
            Return obj_sb.ToString
        End Function

    End Class

End Namespace