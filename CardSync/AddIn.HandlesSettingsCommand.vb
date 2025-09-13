Imports EPDM.Interop.epdm

Partial Public Class AddIn



    Private Sub HandlesSettingsCommand(ppoData() As EdmCmdData, vault As IEdmVault5, loggedInUser As IEdmUser5)


        Dim variableNamesStoredRaw As String = String.Empty

        Dim storage As IEdmDictionary5 = vault.GetDictionary(ADDIN_NAME, True)

        storage.StringGetAt(AddIn.STORAGEKEY, variableNamesStoredRaw)


        If String.IsNullOrWhiteSpace(variableNamesStoredRaw) Then

            variableNamesStoredRaw = String.Empty

        End If

        ' handle case when user clicks cancel

        Dim Answer As String = InputBox("Enter the variable names to synchronize, separated by semicolon (;)", "Variable Names", variableNamesStoredRaw)

        If Answer <> "" Then
            variableNamesStoredRaw = Answer
        End If


        storage.StringSetAt(STORAGEKEY, variableNamesStoredRaw)

    End Sub
End Class
