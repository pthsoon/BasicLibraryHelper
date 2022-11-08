Imports System.Drawing
Imports System.IO

Public Class ImageHelper


    Public Shared Function GetImage(ByVal Path As String) As Image
        Dim result As Image
        Using stream As FileStream = New FileStream(Path, FileMode.Open, FileAccess.Read)
            result = Image.FromStream(stream)
        End Using

        Return result
    End Function

    Public Shared Function GetImage(ByVal Data As Byte()) As Image
        Try
            Dim ms As New MemoryStream(Data)
            Return Image.FromStream(ms)
        Catch ex As Exception
            Throw New Exception(ex.Message, ex.InnerException)
            Return Nothing
        End Try


    End Function


    Public Shared Function GetImageBytes(ByVal Image As Image) As Byte()
        Dim data As Byte()
        Using ms As New MemoryStream()
            Dim bmpImage As New Bitmap(Image)
            bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            data = ms.GetBuffer()
        End Using
        Return data
    End Function

   


    Public Shared Sub ExportImage(ByVal Data As Byte(), Destination As String)
        Using ms As New MemoryStream(Data)
            Using fs As New FileStream(Destination, FileMode.Create)
                ms.WriteTo(fs)
            End Using
        End Using
    End Sub

    Public Shared Sub ExportImage(ByVal Image As Image, Destination As String)

        Dim data As Byte() = GetImageBytes(Image)
        Using ms As New MemoryStream(Data)
            Using fs As New FileStream(Destination, FileMode.Create)
                ms.WriteTo(fs)
            End Using
        End Using
    End Sub

End Class
