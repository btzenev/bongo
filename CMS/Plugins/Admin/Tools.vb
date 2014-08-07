Namespace Plugins.Admin

    Public Class Tools
        Inherits Admin.Base

        Public Overrides ReadOnly Property Name() As String
            Get
                Return "Management Tools"
            End Get
        End Property

        Public Overrides Function Render() As String
            Return vbNewLine
        End Function

        Public Overrides Sub Menu(ByRef vobj_tp As HTMLTemplate)
            On Error Resume Next
            Dim str_path As String = MapPath("/cms/tools/")
            vobj_tp("sections").AddItem("name", MyClass.Name)
            For Each str_file As String In GetFiles(str_path)
                WriteFile(vobj_tp, "", str_file)
            Next
            For Each str_folder As String In IO.Directory.GetDirectories(str_path)
                vobj_tp("sections").AddItem("name", StrConv(IO.Path.GetFileName(str_folder).Replace("-", " "), VbStrConv.ProperCase))
                For Each str_file As String In GetFiles(str_folder)
                    WriteFile(vobj_tp, IO.Path.GetFileName(str_folder), str_file)
                Next
            Next
        End Sub

        Private Function GetFiles(ByVal vstr_path As String) As String()
            Dim obj_list As New List(Of String)
            obj_list.AddRange(IO.Directory.GetFiles(vstr_path, "*.aspx"))
            obj_list.AddRange(IO.Directory.GetFiles(vstr_path, "*.htm"))
            obj_list.Sort()
            Return obj_list.ToArray
        End Function

        Private Sub WriteFile(ByRef vobj_tp As HTMLTemplate, ByVal vstr_folder As String, ByVal vstr_file As String)
            vstr_file = IO.Path.GetFileName(vstr_file)
            If vstr_folder.Length > 0 Then vstr_folder &= "/"
            vobj_tp("sections.links").AddItem("name", StrConv(IO.Path.GetFileNameWithoutExtension(vstr_file).Replace("-", " "), VbStrConv.ProperCase))
            vobj_tp("sections.links").SetItem("link", "/cms/tools/" & vstr_folder & vstr_file)
        End Sub

    End Class

End Namespace