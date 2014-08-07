Namespace Plugins.Hidden

    Public Class Crumbs
        Inherits Plugins.Hidden.Base

    Public Overrides Function Render() As String
      On Error Resume Next
            Return BreadCrumbs(Page.ID)
    End Function

    End Class

End Namespace