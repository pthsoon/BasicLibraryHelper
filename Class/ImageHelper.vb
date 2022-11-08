Imports System.Drawing
Imports System.IO

Public Class ImageHelper


    Public Shared Function GetImage(ByVal Path As String) As Image
        Dim result As Image = Nothing
        If IO.File.Exists(Path) Then
            Using stream As FileStream = New FileStream(Path, FileMode.Open, FileAccess.Read)
                result = Image.FromStream(stream)
            End Using
        End If
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
        Dim data As Byte() = Nothing
        If Image IsNot Nothing Then
            Using ms As New MemoryStream()
                Dim bmpImage As New Bitmap(Image)
                bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
                data = ms.GetBuffer()
            End Using
        End If
      
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


    '#Save Image
    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    Try
    '        Dim localRef As New BLLCustomer
    '        Dim ms As New MemoryStream()
    '        Dim image As Drawing.Image = PictureBox1.Image
    '        Dim data As Byte() = ImageHelper.GetImageBytes(image)
    '        localRef.AddCustomer(Me.TextBox1.Text, data)

    '        LoadCustomer()
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub

    'Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    '    Try
    '        With OpenFileDialog1
    '            .Filter = ("Images |*.png; *.bmp; *.jpg;*.jpeg; *.gif;")
    '            .FilterIndex = 4
    '        End With
    '        'Clear the file name
    '        OpenFileDialog1.FileName = ""
    '        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
    '            PictureBox1.Image = ImageHelper.GetImage(OpenFileDialog1.FileName)
    '            'PictureBox1.Image = Image.FromFile(OpenFileDialog1.FileName.Clone)
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.ToString())
    '    End Try
    'End Sub

    'Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

    '    If Me.DataGridView1.Rows.Count > 0 Then
    '        Dim dr As DataGridViewRow = DataGridView1.CurrentRow
    '        TextBox1.Text = dr.Cells(1).Value.ToString()
    '        If IsDBNull(dr.Cells(2).Value) = False Then
    '            Me.PictureBox1.Image = ImageHelper.GetImage(dr.Cells(2).Value)
    '        Else
    '            Me.PictureBox1.Image = Nothing
    '        End If


    '    End If
    'End Sub

    'Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
    '    Try
    '        With SaveFileDialog1
    '            .Filter = ("Images |*.png; *.bmp; *.jpg;*.jpeg; *.gif;")
    '            .FilterIndex = 4
    '        End With
    '        'Clear the file name

    '        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
    '            ImageHelper.ExportImage(PictureBox1.Image, SaveFileDialog1.FileName)

    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.ToString())
    '    End Try
    'End Sub

End Class
