Namespace Plugins.Client

    Public Class Include
        Inherits Plugins.Client.Base

        Public Overrides ReadOnly Property Name() As String
            Get
                Return "Include Plugin"
            End Get
        End Property

        Public Overrides Function Render() As String
            Try
                Dim obj_sb As New System.Text.StringBuilder(String.Empty)
                Dim obj_row As System.Data.DataRow = Data.Current.Row("select * from [client_include] where modid=" & ModuleID)
                If Not obj_row Is Nothing Then
                    Dim str_file As String = MapPath("" & obj_row("file"))
                    If IO.File.Exists(str_file) Then
                        Return ReadFile(str_file)
                    End If
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
                Dim obj_tp As New HTMLTemplate("/cms/plugins/client_include/edit.htm")
                If Request.Form.Count > 0 Then
                    Dim obj_query As New QueryBuilder("client_include", QueryType.Update)
                    obj_query.Where("modid", ModuleID)
                    obj_query.AddForm("file")
                    Data.Current.Execute(obj_query)
                Else
                    Dim obj_row As DataRow = data.Current.Row("select * from [client_include] where [modid]=" & ModuleID)
                    If obj_row IsNot Nothing Then
                        obj_tp.SetItem("file", "" & obj_row("file"))
                    End If
                End If
                Return obj_tp.Render() & "<input type=""hidden"" name=""action"" value=""close"">"
            Catch ex As Exception
                Return FormatError(ex)
            End Try
        End Function

        Public Overrides Function Add() As Boolean
            'If Not Data.TableExists("client_include") Then
            '    Data.ExecuteNonQuery("CREATE TABLE [client_include]([id] INT IDENTITY (1, 1) PRIMARY KEY, [modid] INT NOT NULL, [file] varchar(255))")
            'End If
            Dim obj_query As New QueryBuilder("client_include", QueryType.Insert)
            obj_query.Add("modid", ModuleID)
            obj_query.Add("file", String.Empty)
            Return Data.Current.Execute(obj_query)
        End Function

        Public Overrides Function Delete() As Boolean
            Return Data.Current.Execute("delete from [client_include] where [modid]=" & ModuleID)
        End Function

    End Class

End Namespace