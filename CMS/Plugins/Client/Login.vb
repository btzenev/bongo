Namespace Plugins.Client

    Public Class Login
        Inherits Plugins.Client.Base

        Public Overrides ReadOnly Property Name() As String
            Get
                Return "Login Plugin"
            End Get
        End Property

        Public Overrides Function Render() As String
      On Error Resume Next
      If Request.IsAuthenticated And Not Security.IsAdmin Then Security.RedirectToMemberPage()
      Dim obj_tp As New HTMLTemplate
      If MultiLingual Then
        If PathExists("/templates/" & GetLang() & "/login.htm") Then
          obj_tp.Template = "/templates/" & GetLang() & "/login.htm"
        ElseIf PathExists("/" & GetLang() & "/login.htm") Then
          obj_tp.Template = "/templates/login.htm"
        Else
          obj_tp.Template = "/cms/templates/loginform.htm"
        End If
      Else
        If PathExists("/templates/login.htm") Then
          obj_tp.Template = "/templates/login.htm"
        Else
          obj_tp.Template = "/cms/templates/loginform.htm"
        End If
      End If

      If Request.Form.Count > 0 Then
        Dim str_msg As String = Security.Login(False)
        If String.IsNullOrEmpty(str_msg) Then
          Security.RedirectToMemberPage()
          Return String.Empty
        Else
          obj_tp.SetItem("message", str_msg)
        End If
      End If
      Return obj_tp.Render
    End Function

    End Class

End Namespace