Public Class dmBase


    Public Shared Function GetStringIsEmpty(ByVal Value As String, ByVal ReturnValue As Object) As Object
        If String.IsNullOrEmpty(Value) Then
            Return ReturnValue
        Else
            Return Value
        End If
    End Function

    Public Shared Function GetNothing(ByVal Value As Object, ByVal ReturnValue As Object) As Object

        If TypeOf Value Is Date Then
            If Value = Nothing Then
                Return ReturnValue
            Else
                Return Value
            End If
        ElseIf Value Is Nothing Then
            Return ReturnValue
        Else
            Return Value
        End If




    End Function

    Public Shared Function GetDBNull(ByVal Value As Object, ByVal ReturnValue As Object) As Object
        If IsDBNull(Value) Then
            Return ReturnValue
        Else
            Return Value
        End If
    End Function



    Public Shared Function GetString(ByVal value As Object) As String
        If IsDBNull(value) Then
            Return ""
        ElseIf IsNothing(value) Then
            Return ""
        Else
            Return value.ToString
        End If
    End Function

    Public Shared Function GetInt(ByVal value As Object) As Integer
        If IsDBNull(value) Then
            Return Nothing
        ElseIf IsNothing(value) Then
            Return Nothing
        ElseIf value = Nothing Then
            Return Nothing
        ElseIf IsNumeric(value) Then
            Return CInt(value)
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function GetDecimal(ByVal value As Object) As Decimal
        If IsDBNull(value) Then
            Return Nothing
        ElseIf IsNothing(value) Then
            Return Nothing
        ElseIf value = Nothing Then
            Return Nothing
        ElseIf IsNumeric(value) Then
            Return CDec(value)
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function GetDate(ByVal value As Object) As Date
        If IsDBNull(value) Then
            Return Nothing
        ElseIf IsNothing(value) Then
            Return Nothing
        ElseIf value = Nothing Then
            Return Nothing
        ElseIf IsDate(value) Then
            Return CDate(value)
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function GetTimeSpan(ByVal value As Object) As TimeSpan
        If IsDBNull(value) Then
            Return Nothing
        ElseIf IsNothing(value) Then
            Return Nothing
        ElseIf value = Nothing Then
            Return Nothing
        ElseIf TypeOf value Is TimeSpan Then
            Dim Time As TimeSpan = value
            Return New TimeSpan(Time.Hours, Time.Minutes, 0)
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function GetBoolean(ByVal value As Object) As Boolean
        If IsDBNull(value) Then
            Return False
        ElseIf IsNothing(value) Then
            Return False
        Else
            Return CBool(value)
        End If
    End Function



End Class
