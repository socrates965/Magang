Module MainContainer
    Public mainContainer As Boolean

    Public Sub mainContainerChecker()
        If mainContainer = False Then
            Form2.Close()
        End If
    End Sub

End Module
