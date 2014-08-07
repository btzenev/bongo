
Namespace Plugins.Admin

    Public MustInherit Class Base
        Implements [Interface]

        Public MustOverride ReadOnly Property Name() As String Implements [Interface].Name

        Public Overridable ReadOnly Property Full() As Boolean Implements [Interface].Full
            Get
                Return False
            End Get
        End Property

        Public MustOverride Function Render() As String Implements [Interface].Render

        Public Overridable Sub Menu(ByRef vobj_tp As HTMLTemplate) Implements [Interface].Menu
            Return
        End Sub

    End Class

End Namespace