Namespace RegistryManagement

    ''' <summary>
    ''' Represents a registry hive and provides access to the Windows registry.
    ''' This class implements the <see cref="IRegistryHive"/> interface to manage registry keys.
    ''' </summary>
    Friend Class MicrosoftRegistryHive
        Implements IRegistryHive

        ''' <summary>
        ''' The base registry key for the hive, specifically targeting the current user.
        ''' </summary>
        Private Shared ReadOnly Hive As RegistryKey = Registry.CurrentUser

        ''' <summary>
        ''' Creates a new subkey or opens an existing subkey in the registry hive.
        ''' </summary>
        ''' <param name="keyName">The name or path of the subkey to create or open.</param>
        ''' <returns>
        ''' An instance of <see cref="IRegistryKey"/> representing the created or opened subkey.
        ''' If the subkey is newly created, it will be empty; if opened, it will contain existing values.
        ''' </returns>
        Friend Function CreateSubKey(keyName As String) As IRegistryKey Implements IRegistryHive.CreateSubKey
            Dim registryKey As RegistryKey = Hive.CreateSubKey(keyName)
            Return New MicrosoftRegistryKey(registryKey)
        End Function

        ''' <summary>
        ''' Opens a subkey in the registry hive with the specified access mode.
        ''' </summary>
        ''' <param name="keyName">The name or path of the subkey to open.</param>
        ''' <param name="accessMode">The access mode to open the subkey with, either read or read/write.</param>
        ''' <returns>
        ''' An instance of <see cref="IRegistryKey"/> representing the opened subkey.
        ''' Returns <c>null</c> if the subkey does not exist.
        ''' </returns>
        Friend Function OpenSubKey(keyName As String, accessMode As RegistryMode) As IRegistryKey Implements IRegistryHive.OpenSubKey
            Dim writable As Boolean = (accessMode = RegistryMode.ReadWrite)
            Dim key As RegistryKey = Hive.OpenSubKey(keyName, writable)
            Dim abstractKey As MicrosoftRegistryKey = Nothing
            
            If key IsNot Nothing Then
                abstractKey = New MicrosoftRegistryKey(key)
            End If

            Return abstractKey
        End Function
    End Class
End Namespace
