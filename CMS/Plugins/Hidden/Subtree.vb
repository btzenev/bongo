Namespace Plugins.Hidden

  Public Class Subtree
    Inherits Plugins.Hidden.Base

    Public Overrides Function Render() As String
      Dim obj_sub As New Client.Subtree()
      obj_sub.Page = MyBase.Page
      obj_sub.ModuleID = -1
      Return obj_sub.Render()
    End Function

  End Class

End Namespace