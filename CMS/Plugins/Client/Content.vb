Namespace Plugins.Client

    Public Class Content
        Inherits Plugins.Client.Base

        Public Overrides ReadOnly Property Name() As String
            Get
                Return "Content Plugin"
            End Get
        End Property

        Public Overrides Function Render() As String
            Try
        Dim obj_sb As New System.Text.StringBuilder(String.Empty)
        Dim obj_row As System.Data.DataRow = Data.Current.Row("select * from client_html where modid=" & ModuleID)
                If Not obj_row Is Nothing Then
                    Dim str_html As String = "" & obj_row("html")
                    str_html = StripWord(str_html)
                    str_html = ConvertUnicode(str_html)
                    obj_sb.Append(str_html & vbNewLine)
                    obj_sb.Append("" & obj_row("script"))
                End If
                Return obj_sb.ToString
            Catch ex As System.Exception
                Return FormatError(ex)
            End Try
        End Function

        Public Overrides Function Edit() As String
      Try
        Dim str_html As String = String.Empty
        Dim str_script As String = String.Empty
        Dim obj_tp As New HTMLTemplate("/cms/plugins/client_html/edit.htm")
        If Request.Form.Count > 0 Then
          Dim obj_query As New QueryBuilder("client_html", QueryType.Update)
          obj_query.Where("modid", ModuleID)
          str_html = ConvertUnicode(GetForm("html"))
          str_html = str_html.Replace("[[", "{{").Replace("]]", "}}")
          str_html = str_html.Replace("<br>", "<br />")
          obj_query.Add("html", str_html)
          str_script = ConvertUnicode(GetForm("script"))
          str_script = str_script.Replace("[[", "{{").Replace("]]", "}}")
          obj_query.Add("script", str_script)
          Data.Current.Execute(obj_query)
        Else
          Dim obj_row As DataRow = Data.Current.Row("select * from [client_html] where [modid]=" & ModuleID)
          If obj_row IsNot Nothing Then
            str_html = "" & obj_row("html")
            str_script = "" & obj_row("script")
          End If
        End If
        obj_tp.SetItem("html", str_html.Replace("{{", "[[").Replace("}}", "]]"))
        obj_tp.SetItem("script", str_script.Replace("{{", "[[").Replace("}}", "]]"))
        Return obj_tp.Render() & "<input type=""hidden"" name=""action"" value=""close"">"
      Catch ex As Exception
        Return FormatError(ex)
      End Try
        End Function

        Public Overrides Function Add() As Boolean
            'If Not Data.TableExists("client_html") Then
            '    Data.ExecuteNonQuery("CREATE TABLE [client_html]([id] INT IDENTITY (1, 1) PRIMARY KEY, [modid] INT NOT NULL, [html] varchar(10000), [script] varchar(10000))")
            'End If
            Dim obj_query As New QueryBuilder("client_html", QueryType.Insert)
            obj_query.Add("modid", ModuleID)
            obj_query.Add("html", String.Empty)
            Return Data.Current.Execute(obj_query)
        End Function

        Public Overrides Function Delete() As Boolean
            Return Data.Current.Execute("delete from [client_html] where modid=" & ModuleID)
        End Function

    End Class

End Namespace