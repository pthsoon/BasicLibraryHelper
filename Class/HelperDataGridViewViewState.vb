Imports System.Windows.Forms

Public Class HelperDataGridViewViewState


    Private Property ScrollIndex As Integer
    Private Property RowIndex As Integer
    Private Property ColumnIndex As Integer

   

    Public Function GetScrollingRowIndex(ByVal dgv As DataGridView) As Integer
        ScrollIndex = dgv.FirstDisplayedScrollingRowIndex
        Return ScrollIndex
    End Function

    Public Sub SetScrollingRowIndex(ByRef dgv As DataGridView)
        If ScrollIndex > 0 Then
            Try
                dgv.FirstDisplayedScrollingRowIndex = ScrollIndex
            Catch ex As Exception

            End Try

        End If

    End Sub



End Class
