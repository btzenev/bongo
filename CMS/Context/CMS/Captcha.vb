Imports System
Imports System.text
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging

Namespace Context.CMS

    Public Class Captcha
        Inherits System.Web.UI.Page

        Private Shared mobj_images As New Dictionary(Of String, String)

        Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
            SetCache(HttpCacheability.NoCache)
            Dim str_guid As String = GetQuery("guid")
            If String.IsNullOrEmpty(str_guid) Then Return
            Dim str_val As String = GetQuery("val")
            If Not String.IsNullOrEmpty(str_val) Then
                Dim b_success As Boolean = False
                If mobj_images.ContainsKey(str_guid) Then
                    b_success = (mobj_images(str_guid) = str_val)
                    mobj_images.Remove(str_guid)
                End If
                Response.Write("{""success"":" & b_success.ToString.ToLower & "}")
            Else
                If mobj_images.ContainsKey(str_guid) Then mobj_images.Remove(str_guid)
                Dim str_code As String = GenerateRandomCode()
                mobj_images.Add(str_guid, str_code)
                Context.Response.Clear()
                Context.Response.ContentType = "image/gif"
                Using obj_bmp As Bitmap = GenerateImage(str_code, GetQuery("w", 230), GetQuery("h", 40))
                    obj_bmp.Save(Response.OutputStream, ImageFormat.Gif)
                End Using
            End If
        End Sub

        Private Function GenerateRandomCode() As String
            Dim obj_sb As New System.Text.StringBuilder
            Dim str_chars As String = "123456789abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ"
            Dim obj_rnd As New Random()
            For i As Integer = 0 To 5
                obj_sb.Append(str_chars.Substring(obj_rnd.Next(str_chars.Length), 1))
            Next
            Dim str_temp As String = obj_sb.ToString
            Return str_temp
        End Function

        Private Function GenerateImage(ByVal vstr_text As String, ByVal vint_width As Integer, ByVal vint_height As Integer) As Bitmap
            ' Create a new 32-bit bitmap image.
            Using obj_bmp As Bitmap = New Bitmap(vint_width, vint_height, PixelFormat.Format32bppArgb)
                ' Create a graphics object for drawing.
                Using obj_graph As Graphics = Graphics.FromImage(obj_bmp)
                    obj_graph.SmoothingMode = SmoothingMode.AntiAlias
                    Dim rect As RectangleF = New RectangleF(0, 0, vint_width, vint_height)
                    ' Fill in the background.
                    Dim hatchBrush As HatchBrush = New HatchBrush(HatchStyle.SmallConfetti, Color.LightGray, Color.White)
                    obj_graph.FillRectangle(hatchBrush, rect)
                    ' Set up the text font.
                    Dim size As SizeF
                    Dim fontSize As Single = rect.Height + 1
                    Dim font As Font
                    ' Adjust the font size until the text fits within the image.
                    While (size.Width > rect.Width)
                        fontSize = fontSize - 1
                        font = New Font("Arial", fontSize, FontStyle.Bold)
                        size = obj_graph.MeasureString(vstr_text, font)
                    End While
                    ' Set up the text format.
                    Dim format As StringFormat = New StringFormat()
                    format.Alignment = StringAlignment.Center
                    format.LineAlignment = StringAlignment.Center
                    font = New Font("Arial", fontSize, FontStyle.Bold)
                    ' Create a path using the text and warp it randomly.
                    Dim path As GraphicsPath = New GraphicsPath()
                    path.AddString(vstr_text, font.FontFamily, CType(font.Style, Integer), font.Size, rect, format)
                    Dim v As Single = 5.0F ' 4.0F
                    Dim points(4) As PointF
                    Dim obj_rnd As New Random
                    points(0) = New PointF(obj_rnd.Next(rect.Width) / v, obj_rnd.Next(rect.Height) / v)
                    points(1) = New PointF(rect.Width - obj_rnd.Next(rect.Width) / v, obj_rnd.Next(rect.Height) / v)
                    points(2) = New PointF(obj_rnd.Next(rect.Width) / v, rect.Height - obj_rnd.Next(rect.Height) / v)
                    points(3) = New PointF(rect.Width - obj_rnd.Next(rect.Width) / v, rect.Height - obj_rnd.Next(rect.Height) / v)
                    Dim obj_matrix As New Matrix()
                    obj_matrix.Translate(0.0F, 0.0F)
                    path.Warp(points, rect, obj_matrix, WarpMode.Perspective, 0.0F)
                    ' Draw the text.
                    hatchBrush = New HatchBrush(HatchStyle.LargeConfetti, Color.DimGray, Color.SlateGray)
                    obj_graph.FillPath(hatchBrush, path)
                    hatchBrush = New HatchBrush(HatchStyle.LargeConfetti, Color.DimGray, Color.LightGray)
                    ' Add some random noise.
                    Dim m As Integer = Math.Max(rect.Width, rect.Height)
                    For i As Integer = 0 To CInt((rect.Width * rect.Height / 30.0F))
                        Dim x As Integer = obj_rnd.Next(rect.Width)
                        Dim y As Integer = obj_rnd.Next(rect.Height)
                        Dim w As Integer = obj_rnd.Next(CInt(m / 50))
                        Dim h As Integer = obj_rnd.Next(CInt(m / 50))
                        obj_graph.FillEllipse(hatchBrush, x, y, w, h)
                    Next
                End Using
                Return obj_bmp
            End Using
        End Function

    End Class

End Namespace
