Imports System.Windows.Forms

Public Class HelperDate

    'Date Formate dd/MM/yyyy Only
    Private Const strDateFormat As String = "dd/MM/yyyy"
    Private Const strDateHourMinuteFormat As String = "dd/MM/yyyy   hh:mm tt"

    Public Shared Function GetDate(ByVal textValue As String) As Date
        Dim intDay As Integer
        Dim IntMonth As Integer
        Dim IntYear As Integer
        Dim ErrorMsg As String = ""
        Dim DateOutput As Date = Nothing

        textValue = textValue.Trim
        If textValue.Length = 10 Then

            Try
                intDay = CInt(textValue.Substring(0, 2))
                IntMonth = CInt(textValue.Substring(3, 2))
                IntYear = CInt(textValue.Substring(6, 4))

                If IntMonth < 0 Or IntMonth > 12 Then
                    ErrorMsg = "Month Not in Range 1 to 12"
                Else
                    DateOutput = DateSerial(IntYear, IntMonth, intDay)
                End If


            Catch ex As Exception
                ErrorMsg = "Date Format not in dd/MM/yyyy"

            End Try





        Else
            ErrorMsg = "Date Format not in dd/MM/yyyy"

        End If


        If DateOutput = Nothing Then
            If String.IsNullOrEmpty(ErrorToString) = False Then
                Throw New Exception(ErrorMsg)
            End If
        End If

        Return DateOutput
    End Function

    Public Shared Function GetDate(ByVal dtp As DateTimePicker) As Date
        If dtp.Checked = False Then
            Return Nothing
        Else
            Return dtp.Value
        End If
    End Function

    Public Shared Function GetDaysInMonth(ByVal YearMonth As String) As Date
        If YearMonth <> "" And IsNumeric(YearMonth) And Len(YearMonth) = 6 Then
            Return DateSerial(Left(YearMonth, 4), Right(YearMonth, 2), Date.DaysInMonth(Left(YearMonth, 4), Right(YearMonth, 2)))
        Else
            Return Nothing
        End If

    End Function

    Public Shared Function GetHourMinute(ByVal Value As Date) As Date
        If Value = Nothing Then
            Return Nothing
        Else
            Return Value.Date.Add(New TimeSpan(Value.Hour, Value.Minute, 0))
        End If

    End Function

    Public Shared Function GetHourMinute(ByVal dtp As DateTimePicker) As Date
        Return GetHourMinute(GetDate(dtp))
    End Function

    Public Shared Function GetString_Date(ByVal Value As Object) As String
        Try
            If IsDBNull(Value) Then
                Return ""
            End If

            If IsDate(Value) = False Then
                Return ""
            Else
                Dim DateValue As Date = CDate(Value)
                If DateValue = Nothing Then
                    Return ""
                Else
                    Return DateValue.ToString(strDateFormat)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return ""
        End Try
    End Function

    Public Shared Function GetString_DateHourMinute(ByVal value As Object) As String
        Try

            If IsDBNull(value) Then
                Return ""
            ElseIf IsNothing(value) Then
                Return ""
            ElseIf value = Nothing Then
                Return ""
            ElseIf IsDate(value) = False Then
                Return ""
            Else
                Dim DateValue As Date = CDate(value)

                Return DateValue.ToString(strDateHourMinuteFormat)

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Return ""
        End Try


    End Function

    Public Shared Function GetTimeSpan_HourMinute(ByVal value As Date) As TimeSpan
        Return New TimeSpan(value.Hour, value.Minute, 0)
    End Function

    Public Shared Function GetTimeSpan_HourMinute(ByVal value As Integer) As TimeSpan
        Dim Hour As Integer = 0
        Dim Minute As Integer = 0

        Hour = value \ 60
        If Hour > 24 Then
            Hour = 0
        End If
        Minute = value Mod 60

        Return New TimeSpan(Hour, Minute, 0)
    End Function


End Class
