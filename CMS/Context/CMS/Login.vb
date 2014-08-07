Namespace Context.CMS

    Public Class Login
        Inherits System.Web.UI.Page

        Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            Dim str_msg As String = String.Empty
            If Request.Form.Count > 0 Then
                str_msg = Security.Login()
            End If
            Dim obj_tp As New HTMLTemplate("/cms/templates/login.htm")
            obj_tp.SetItem("message", str_msg)
            obj_tp.SetItem("cms.version", Config("cms.version"))
      obj_tp.SetItem("cms.title", NullSafe(Config("cms.name"), "CMS"))
            obj_tp.Render(Response.OutputStream)
        End Sub

    End Class

End Namespace
