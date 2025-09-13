Imports System.Text
Imports EPDM.Interop.epdm

Partial Public Class AddIn
    Private Sub HandlesCardButtonClick(ppoData() As EdmCmdData, vault As IEdmVault5, loggedInUser As IEdmUser5)
        Dim affectedItem = ppoData.First()

        If affectedItem.mlObjectID1 = 0 Then
            Exit Sub
        End If

        ' this is the file that was affected by the card button click
        Dim affectedDocument As IEdmFile5 = vault.GetObject(EdmObjectType.EdmObject_File, affectedItem.mlObjectID1)

        Dim affectedDocumentFolder As IEdmFolder5 = vault.GetObject(EdmObjectType.EdmObject_Folder, affectedItem.mlObjectID2)


        If affectedDocument.Name.ToLower().EndsWith(".slddrw") = False Then
            Throw New Exception("This action is only drawings.")
        End If

        If affectedDocument.IsLocked = False Then
            Throw New Exception("The file must be checked out before synchronizing.")
        End If

        If affectedDocument.LockedByUserID <> loggedInUser.ID Then
            Throw New Exception("The file is locked by another user.")
        End If

        If affectedDocument.LockedOnComputer <> Environment.MachineName Then
            Throw New Exception($"The file is locked on another computer. [{affectedDocument.LockedOnComputer}]")
        End If

        Dim associatedModel As IEdmFile5 = Nothing

        Dim associatedModelFolder As IEdmFolder5 = Nothing

        Dim affectedDocumentRootReference As IEdmReference5 = affectedDocument.GetReferenceTree(affectedDocumentFolder.ID)

        Dim position As IEdmPos5 = affectedDocumentRootReference.GetFirstChildPosition(ADDIN_NAME, True, False)

        Dim firstReference As IEdmReference5 = affectedDocumentRootReference.GetNextChild(position)

        If firstReference Is Nothing Then
            Throw New Exception("The model has no associated model.")
        End If

        associatedModel = vault.GetObject(EdmObjectType.EdmObject_File, firstReference.FileID)

        associatedModelFolder = vault.GetObject(EdmObjectType.EdmObject_Folder, firstReference.FolderID)


        Dim variableNames As String() = GetVariableNames(vault)

        If variableNames.Length = 0 Then
            Throw New Exception("No variable names configured. Please configure the variable names in the add-in settings.")
        End If

        Dim variables As New Dictionary(Of String, String)

        'add all variable names with empty value
        For Each variableName As String In variableNames
            If Not variables.ContainsKey(variableName) Then
                variables.Add(variableName, String.Empty)
            End If
        Next




        GetVariables(associatedModel, associatedModelFolder, variables, errorLogs)


        SetVariables(affectedDocument, affectedDocumentFolder, variables, errorLogs)


    End Sub

End Class
