Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Web.Caching

Namespace Context.CMS

    Public Class Image
        Inherits System.Web.UI.Page


        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            On Error Resume Next
            Response.Cache.SetCacheability(HttpCacheability.Public)
            Response.Cache.SetExpires(DateTime.Now.AddDays(1))

            Dim str_file As String = GetQuery("file").ToLower
            Dim str_ext As String = IO.Path.GetExtension(str_file).TrimStart(".")
            If String.IsNullOrEmpty(str_file) Then Throw New Exception("IMAGE: No Image specified.")
            If GetQuery("reload", 0) = 1 Then Cache.Remove(RawUrl)

            'load image
            Dim obj_image As Drawing.Image = Cache.Get(RawUrl)
            If obj_image Is Nothing Then
                Dim str_path As String = MapPath(str_file)
                obj_image = Drawing.Image.FromFile(str_path)
                Dim obj_org As Size = obj_image.Size
                Dim b_fit As Boolean = GetQuery("fit", 0) = 1
                Dim obj_max As New Size(GetQuery("mw", 0), GetQuery("mh", 0))
                Dim obj_size As New Size(GetQuery("w", 0), GetQuery("h", 0))
                Dim dbl_scale As Double = GetQuery("s", GetQuery("scale", 100))
                Dim obj_rect As New Rectangle(0, 0, obj_image.Width, obj_image.Height)
                Dim b_oversize As Boolean = GetQuery("oversize", 1) = 1

                If Not dbl_scale = 100 Then 'scale
                    Dim dlb_per As Double = dbl_scale / 100
                    obj_rect.Width = CInt(obj_org.Width * dlb_per)
                    obj_rect.Height = CInt(obj_org.Height * dlb_per)
                    Resize(obj_image, obj_rect)
                ElseIf obj_max.Width > 0 Or obj_max.Height > 0 Then 'max size
                    obj_size = New Size(obj_org.Width, obj_org.Height)
                    If obj_max.Width > 0 And obj_max.Width < obj_size.Width Then
                        obj_rect.Width = obj_max.Width
            obj_rect.Height = Math.Round(obj_size.Height * (obj_max.Width / obj_size.Width))
            If obj_rect.Height <= 0 Then obj_rect.Height = 1
                        obj_size = New Size(obj_rect.Width, obj_rect.Height)
                    End If
                    If obj_max.Height > 0 And obj_max.Height < obj_size.Height Then
                        obj_rect.Height = obj_max.Height
            obj_rect.Width = Math.Round(obj_size.Width * (obj_max.Height / obj_size.Height))
            If obj_rect.Width <= 0 Then obj_rect.Width = 1
                    End If
                    Resize(obj_image, obj_rect)
                ElseIf obj_size.Width > 0 Or obj_size.Height > 0 Then 'resize

                    If b_fit Then

                        If obj_size.Width > 0 And obj_size.Height > 0 Then

                            Dim s_org As Single = obj_org.Width / obj_org.Height
                            Dim s_dest As Single = obj_size.Width / obj_size.Height

                            If s_dest > s_org Then
                                obj_rect.Width = obj_size.Width
                                If Not b_oversize And obj_rect.Width > obj_org.Width Then obj_rect.Width *= obj_org.Width / obj_rect.Width
                                obj_rect.Height = Math.Round((obj_org.Height * obj_rect.Width) / obj_org.Width)
                                Resize(obj_image, obj_rect)
                                obj_rect.Height = obj_size.Height
                                Crop(obj_image, obj_rect)
                            ElseIf s_dest < s_org Then
                                obj_rect.Height = obj_size.Height
                                If Not b_oversize And obj_rect.Height > obj_org.Height Then obj_rect.Height *= obj_org.Height / obj_rect.Height
                                obj_rect.Width = Math.Round((obj_org.Width * obj_rect.Height) / obj_org.Height)
                                Resize(obj_image, obj_rect)
                                obj_rect.Width = obj_size.Width
                                Crop(obj_image, obj_rect)
                            Else
                                obj_rect.Width = obj_size.Width
                                obj_rect.Height = obj_rect.Height
                                Resize(obj_image, obj_rect)
                            End If

                        Else
                            If obj_size.Width > 0 Then
                                obj_rect.Width = obj_size.Width
                                obj_rect.Height = obj_rect.Height
                            Else
                                obj_rect.Height = obj_size.Height
                                obj_rect.Width = obj_rect.Width
                            End If
                        End If

                    Else
                        If obj_size.Width > 0 And obj_size.Height > 0 Then
                            obj_rect.Width = obj_size.Width
                            obj_rect.Height = obj_size.Height
                        Else
                            If obj_size.Width > 0 Then
                                obj_rect.Width = obj_size.Width
                                obj_rect.Height = Math.Round(obj_image.Height * (obj_size.Width / obj_image.Width))
                            Else
                                obj_rect.Height = obj_size.Height
                                obj_rect.Width = Math.Round(obj_image.Width * (obj_size.Height / obj_image.Height))
                            End If
                        End If
                        Resize(obj_image, obj_rect)
                    End If

                Else
                    'nothing
                End If

                If Not String.IsNullOrEmpty(GetQuery("flip")) Then Flip(obj_image, GetQuery("flip"))
                If GetQuery("gray", GetQuery("grey", 0)) = 1 Then GrayScale(obj_image)
                If Not String.IsNullOrEmpty(GetQuery("adjust")) Then ColorAdjust(obj_image, GetQuery("adjust"))

                Using obj_dep As New Caching.CacheDependency(str_path)
                    Cache.Insert(RawUrl, obj_image, obj_dep, Cache.NoAbsoluteExpiration, TimeSpan.FromDays(1), Caching.CacheItemPriority.Default, New CacheItemRemovedCallback(AddressOf RemovedCache))
                End Using

            End If

            'show image
            Dim str_format As String = GetQuery("format", GetQuery("frm", str_ext))
            Dim obj_format As ImageFormat
            Select Case str_format.ToLower
                Case "bmp"
                    Response.ContentType = "image/x-bmp"
                    obj_format = ImageFormat.Bmp
                Case "png"
                    Response.ContentType = "image/png"
                    obj_format = ImageFormat.Png
                Case "gif"
                    Response.ContentType = "image/gif"
                    obj_format = ImageFormat.Gif
                Case Else
                    Response.ContentType = "image/jpeg"
                    obj_format = ImageFormat.Jpeg
            End Select

            If str_file.Length = 0 Then str_file = DateAndTime.Now.ToShortDateString
            If GetQuery("force", False) Then
                Me.Response.AddHeader("Content-Disposition", ("attachment;filename=" & str_file & "." & obj_format.ToString.ToLower))
            Else
                Me.Response.AddHeader("Content-Disposition", ("inline;filename=" & str_file & "." & obj_format.ToString.ToLower))
            End If

            obj_image.Save(Response.OutputStream, obj_format)
            'obj_image.Dispose()
            Response.Flush()
            Response.AddHeader("Content-Length", Response.OutputStream.Length)
        End Sub

        Private Sub RemovedCache(ByVal vstr_key As String, ByVal vobj_value As Object, ByVal vobj_reason As CacheItemRemovedReason)
            On Error Resume Next
            DirectCast(vobj_value, Drawing.Image).Dispose()
        End Sub

        Private Function GetSize(ByRef vobj_image As Drawing.Image, ByVal vstr_dim As String, Optional ByVal vb_auto As Boolean = True) As Size
            On Error Resume Next
            Dim obj_size As New Size()
            Dim arr_dims() As String = vstr_dim.Split("x")
            obj_size.Width = NullSafe(arr_dims(0), 0)
            If vb_auto And obj_size.Width <= 0 Then obj_size.Width = vobj_image.Width
            obj_size.Height = NullSafe(arr_dims(1), 0)
            If vb_auto And obj_size.Height <= 0 Then obj_size.Height = vobj_image.Height
            Return obj_size
        End Function

        Private Sub Crop(ByRef vobj_image As Drawing.Image, ByVal vobj_rect As Rectangle)
            On Error Resume Next
            Using obj_bmp As Drawing.Bitmap = vobj_image.Clone
                vobj_image.Dispose()
                vobj_image = obj_bmp.Clone(vobj_rect, obj_bmp.PixelFormat)
            End Using
        End Sub

        Private Sub Resize(ByRef vobj_image As Drawing.Image, ByVal vobj_rect As Rectangle)
            On Error Resume Next
            Using obj_bmp As Drawing.Bitmap = New Bitmap(vobj_rect.Width, vobj_rect.Height)
                Using obj_graph As Graphics = Graphics.FromImage(obj_bmp)
                    obj_graph.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                    obj_graph.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                    obj_graph.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    obj_graph.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
                    obj_graph.DrawImage(vobj_image, vobj_rect)
                    vobj_image.Dispose()
                    vobj_image = obj_bmp.Clone
                End Using
            End Using
        End Sub

        Private Sub ColorAdjust(ByRef vobj_image As Drawing.Image, ByVal vstr_hex As String)
            On Error Resume Next
            Dim obj_color As Color = ColorTranslator.FromHtml("#" & vstr_hex.Trim("#"))
            ' noramlize the color components to 1            
            Dim s_red As Single = obj_color.R / 255
            Dim s_green As Single = obj_color.G / 255
            Dim s_blue As Single = obj_color.B / 255
            ' create the color matrix
            Dim obj_matrix As New ColorMatrix(New Single()() { _
              New Single() {1, 0, 0, 0, 0}, _
              New Single() {0, 1, 0, 0, 0}, _
              New Single() {0, 0, 1, 0, 0}, _
              New Single() {0, 0, 0, 1, 0}, _
              New Single() {s_red, s_green, s_blue, 0, 1} _
            })
            ApplyColorMatrix(vobj_image, obj_matrix)
        End Sub

        Private Sub GrayScale(ByRef vobj_image As Drawing.Image)
            On Error Resume Next
            Dim obj_matrix As New ColorMatrix(New Single()() { _
              New Single() {0.3086F, 0.3086F, 0.3086F, 0.0F, 0.0F}, _
              New Single() {0.6094F, 0.6094F, 0.6094F, 0.0F, 0.0F}, _
              New Single() {0.082F, 0.082F, 0.082F, 0.0F, 0.0F}, _
              New Single() {0.0F, 0.0F, 0.0F, 1.0F, 0.0F}, _
              New Single() {0.0F, 0.0F, 0.0F, 0.0F, 1.0F} _
            })
            ApplyColorMatrix(vobj_image, obj_matrix)
        End Sub

        Private Sub Flip(ByRef vobj_image As Drawing.Image, ByVal vstr_dir As String)
            On Error Resume Next
            Using obj_bmp As Drawing.Bitmap = New Bitmap(vobj_image)
                obj_bmp.RotateFlip(IIf(vstr_dir.ToLower = "h", RotateFlipType.RotateNoneFlipX, RotateFlipType.RotateNoneFlipY))
                vobj_image.Dispose()
                vobj_image = obj_bmp.Clone
            End Using
        End Sub

        Private Sub ApplyColorMatrix(ByRef vobj_image As Drawing.Image, ByVal vobj_matrix As ColorMatrix)
            Using obj_graph As Graphics = Graphics.FromImage(vobj_image)
                Using obj_attr As New ImageAttributes
                    obj_attr.SetColorMatrix(vobj_matrix)
                    obj_graph.DrawImage(vobj_image, New Rectangle(0, 0, vobj_image.Width, vobj_image.Height), 0, 0, vobj_image.Width, vobj_image.Height, GraphicsUnit.Pixel, obj_attr)
                    obj_graph.Save()
                End Using
            End Using
        End Sub

    End Class

End Namespace
