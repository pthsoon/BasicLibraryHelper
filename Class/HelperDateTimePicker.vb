Imports System.Windows.Forms

Public Class HelperDateTimePicker

    Public Shared Sub SetValue(ByRef dtp As DateTimePicker, ByVal DateValue As Date)

        If DateValue = Nothing Then
            dtp.Checked = False
        Else
            dtp.Checked = True
            dtp.Value = DateValue
        End If

    End Sub

    Public Shared Sub SetValue(ByRef dtp As DateTimePicker, ByVal DateValue As Object)

        If IsDBNull(DateValue) Then
            dtp.Checked = False
            Exit Sub
        End If

        If DateValue = Nothing Then
            dtp.Checked = False
            Exit Sub
        End If

        dtp.Checked = True
        dtp.Value = DateValue


    End Sub


End Class
