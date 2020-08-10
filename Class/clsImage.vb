Imports System.Drawing.Imaging

Public Class clsImage
    Private Image As Image

    Public Function LoadImage(ByVal SourceFile As String) As Image

        Using bitmap = New Bitmap(SourceFile)
            image = New Bitmap(bitmap)

        End Using
        Return Image
    End Function


    Public Sub SaveFileJPG(ByVal SourceFile As String, ByVal DestinationFile As String)
        Dim Image As Image = LoadImage(SourceFile)
        Image.Save(DestinationFile, ImageFormat.Jpeg)

    End Sub

    Private _disposed As Boolean
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If _disposed Then Return

        If disposing Then


            If image IsNot Nothing Then
                image.Dispose()
            End If
        End If

        _disposed = True
    End Sub
    Public Sub Dispose()
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
End Class
