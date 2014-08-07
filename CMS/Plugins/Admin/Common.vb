Namespace Plugins.Admin

    Public Module Common

        Public Interface [Interface]
            ReadOnly Property Name() As String
            ReadOnly Property Full() As Boolean
            Sub Menu(ByRef vobj_tp As HTMLTemplate)
            Function Render() As String
        End Interface

        Private mobj_plugins As New List(Of String)

        Public ReadOnly Property List() As List(Of String)
            Get
                Return mobj_plugins
            End Get
        End Property

        Public ReadOnly Property [Get](ByVal vstr_name As String) As [Interface]
            Get
                Try
                    Dim obj_ex As System.Exception = Nothing
                    Dim obj_type As Type = Plugins.Get(vstr_name, obj_ex)
                    If obj_ex IsNot Nothing Then Throw obj_ex
                    Dim obj_const As Reflection.ConstructorInfo = obj_type.GetConstructor(New Type() {})
                    Return obj_const.Invoke(Nothing)
                Catch ex As Exception
                    Return New Admin.Error(vstr_name, ex)
                End Try
            End Get
        End Property

        Public Sub Clear()
            List.Clear()
        End Sub

    End Module

End Namespace