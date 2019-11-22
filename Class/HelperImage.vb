Imports System.Drawing
Imports System.Windows
Imports System.Windows.Forms

Public Class HelperImage


    Public Shared Function GetImageBytes(ByVal img As Image) As Byte()
        If img Is Nothing Then
            Return Nothing
        End If
        Dim converter As ImageConverter = New ImageConverter()
        Return CType(converter.ConvertTo(img, GetType(Byte())), Byte())
        'Using ms As IO.MemoryStream = New IO.MemoryStream()
        '    img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
        '    Return ms.ToArray
        'End Using
    End Function

    Public Shared Function GetImageBytes(ByVal PictureBox As PictureBox) As Byte()
        Dim Img As Image = PictureBox.Image

        Return GetImageBytes(Img)
    End Function


    Public Shared Function GetImageBytes(ByVal FileName As String) As Byte()
        Return GetImageBytes(GetImage(FileName))
    End Function

    Public Shared Function GetImage(ByVal imageData As Byte()) As Image


        Dim NewImage As Image

        If imageData Is Nothing Then
            NewImage = Nothing
        Else

            Using ms As New IO.MemoryStream(imageData, 0, imageData.Length)
                ms.Write(imageData, 0, imageData.Length)
                NewImage = Image.FromStream(ms, True, True)
            End Using





        End If

        Return NewImage
    End Function

    Public Shared Function GetImage(ByVal filename As String) As Image

        Using originalImage As Image = Image.FromFile(filename)
            Dim newWidth As Integer = originalImage.Width
            Dim newHeight As Integer = originalImage.Height
            Dim aspectRatio As Double = CDbl(originalImage.Width) / CDbl(originalImage.Height)

            Dim newImage As Bitmap = New Bitmap(newWidth, newHeight)

            Using g As Graphics = Graphics.FromImage(newImage)
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality
                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                g.DrawImage(originalImage, 0, 0, newImage.Width, newImage.Height)
                Return newImage
            End Using
        End Using


    End Function

    Public Shared Function ResizeImage(ByVal filename As String, ByVal maxWidth As Integer, ByVal maxHeight As Integer) As Bitmap
        Using originalImage As Image = Image.FromFile(filename)
            Dim newWidth As Integer = originalImage.Width
            Dim newHeight As Integer = originalImage.Height
            Dim aspectRatio As Double = CDbl(originalImage.Width) / CDbl(originalImage.Height)

            If aspectRatio <= 1 AndAlso originalImage.Width > maxWidth Then
                newWidth = maxWidth
                newHeight = CInt(Math.Round(newWidth / aspectRatio))
            ElseIf aspectRatio > 1 AndAlso originalImage.Height > maxHeight Then
                newHeight = maxHeight
                newWidth = CInt(Math.Round(newHeight * aspectRatio))
            End If

            Dim newImage As Bitmap = New Bitmap(newWidth, newHeight)

            Using g As Graphics = Graphics.FromImage(newImage)
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality
                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                g.DrawImage(originalImage, 0, 0, newImage.Width, newImage.Height)
                Return newImage
            End Using
        End Using
    End Function

    Public Shared Function ResizeImageTo(ByVal filename As String, ByVal Width As Integer, ByVal Height As Integer) As Bitmap
        Using originalImage As Image = Image.FromFile(filename)
            Dim newWidth As Integer = originalImage.Width
            Dim newHeight As Integer = originalImage.Height
            Dim aspectRatio As Double = CDbl(originalImage.Width) / CDbl(originalImage.Height)

            Dim newImage As Bitmap = New Bitmap(Width, Height)

            Using g As Graphics = Graphics.FromImage(newImage)
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality
                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality
                g.DrawImage(originalImage, 0, 0, newImage.Width, newImage.Height)
                Return newImage
            End Using
        End Using
    End Function
End Class
