Imports System.Runtime.InteropServices
Imports System.Text
Imports EPDM.Interop.epdm




<Guid("25421DEE-47BC-4E5B-ADC7-AB3E4548D18B")>
<ComVisible(True)>
Partial Public Class AddIn
    Implements IEdmAddIn5




    Private Const ADDIN_NAME As String = "CardSync"
    Private Const STORAGEKEY As String = "VariableNames"
    Dim errorLogs As New StringBuilder()

    ' This method is called by the vault to get information about the add-in.
    Public Sub GetAddInInfo(ByRef poInfo As EdmAddInInfo, poVault As IEdmVault5, poCmdMgr As IEdmCmdMgr5) Implements IEdmAddIn5.GetAddInInfo

        ' Set the add-in information.

        poInfo.mbsAddInName = ADDIN_NAME
        poInfo.mbsCompany = "BLUE BYTE SYSTEMS INC."
        poInfo.mbsDescription = "Synchronizes data card values between drawing and model"
        'suggest format yyyy-mm-dd for versioning 
        poInfo.mlAddInVersion = 1

        ' 18 corresponding to SolidWorks 2018
        ' more information here: https://github.com/BlueByteSystemsInc/SOLIDWORKS-PDM-API-SDK/blob/master/src/BlueByte.SOLIDWORKS.PDMProfessional.SDK/Enums/PDMProfessionalVersion_e.cs#L12
        poInfo.mlRequiredVersionMajor = 18
        ' this is the service pack number, 0 to 5
        poInfo.mlRequiredVersionMinor = 0

        ' Register a command for the add-in.

        ' listen to card button 
        poCmdMgr.AddHook(EdmCmdType.EdmCmd_CardButton)

        'listen to menu
        poCmdMgr.AddHook(EdmCmdType.EdmCmd_Menu)


        ' Add a menu item to the tools menu.
        poCmdMgr.AddCmd(Commands.Sync, "Sync\\Datacard", EdmMenuFlags.EdmMenu_OnlyFiles, "Synchronizes data card values between drawing and model", "Sync datacards", -1, 0)
        poCmdMgr.AddCmd(Commands.Settings, "Settings", EdmMenuFlags.EdmMenu_Administration, "Settings", "Settings", -1, 0)




    End Sub

    ' This method is called by the vault when a command associated with the add-in is executed.
    Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData() As EdmCmdData) Implements IEdmAddIn5.OnCmd

        Dim vault As IEdmVault5 = poCmd.mpoVault
        Dim userMgr As IEdmUserMgr5 = vault
        Dim handle As Integer = poCmd.mlParentWnd
        Dim loggedInUser As IEdmUser5 = userMgr.GetLoggedInUser()
        Dim er As String = String.Empty

        Try

            If poCmd.meCmdType = EdmCmdType.EdmCmd_CardButton And poCmd.mbsComment = ADDIN_NAME Then


                'handle card button click
                HandlesCardButtonClick(ppoData, vault, loggedInUser)


                If String.IsNullOrWhiteSpace(errorLogs.ToString()) = False Then
                    er = vbCrLf & "Errors:" & vbCrLf & errorLogs.ToString()
                    errorLogs.Clear()
                End If

                vault.MsgBox(handle, $"Sync complete{er}", EdmMBoxType.EdmMbt_OKOnly, ADDIN_NAME)

                'unfortunately there is no way to refresh the datcard 
                'please vote here for DS to add it: https://r1132100503382-eu1-3dswym.3dexperience.3ds.com/community/swym:prd:R1132100503382:community:I23-EB1ZRQyk2sLG5y9Kig?content=swym:prd:R1132100503382:idea:RMWyoOyfRRmz4qHmAsJjDA

            ElseIf (poCmd.meCmdType = EdmCmdType.EdmCmd_Menu And poCmd.mlCmdID = Commands.Sync) Then
                'handle menu settings
                HandlesSyncCommand(ppoData, vault, loggedInUser)

                If String.IsNullOrWhiteSpace(errorLogs.ToString()) = False Then
                    er = vbCrLf & "Errors:" & vbCrLf & errorLogs.ToString()
                    errorLogs.Clear()
                End If

                vault.MsgBox(handle, $"Sync complete{er}", EdmMBoxType.EdmMbt_OKOnly, ADDIN_NAME)

            ElseIf (poCmd.meCmdType = EdmCmdType.EdmCmd_Menu And poCmd.mlCmdID = Commands.Settings) Then
                'handle settings command
                HandlesSettingsCommand(ppoData, vault, loggedInUser)
            End If





        Catch ex As Exception

            vault.MsgBox(handle, ex.Message, EdmMBoxType.EdmMbt_OKOnly, "Error")

        End Try
    End Sub

End Class



Public Enum Commands
    Sync = 156156
    Settings = 156155
End Enum
