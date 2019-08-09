Public Class misc
    Public Shared Function getRound(ByRef myvalue As Decimal, ByVal factor As Integer) As Decimal
        Dim myResult As Decimal = 0
        Dim i As Integer = 0
        Dim myPoint() As Decimal = {1, 0.1, 0.01, 0.001, 0.0001, 0.00001, 0.000001, 0.0000001, 0.00000001}

        While True
            myResult = Math.Round(myvalue, i)
            If myResult > 0 Then
                myResult = Math.Round(myvalue, i + 3)
                Select Case factor
                    Case 1
                        If myResult < myvalue Then
                            myResult = myResult + (myPoint(i + 3) * factor)
                        End If
                    Case -1
                        If myResult > myvalue Then
                            myResult = myResult + (myPoint(i + 3) * factor)
                        End If
                End Select
                Exit While
            End If

            i = i + 1
        End While

        Return myResult
    End Function
End Class
