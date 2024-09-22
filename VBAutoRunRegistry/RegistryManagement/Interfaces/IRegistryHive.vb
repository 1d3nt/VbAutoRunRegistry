Imports VBAutoRunRegistry.RegistryManagement.RegistryAccess

Namespace RegistryManagement.Interfaces

    ''' <summary>
    ''' Provides an interface for managing registry keys within a registry hive.
    ''' This interface allows for the creation and opening of subkeys, enabling
    ''' interaction with the Windows Registry.
    ''' </summary>
    Public Interface IRegistryHive

        ''' <summary>
        ''' Creates a new subkey or opens an existing subkey within the registry hive.
        ''' </summary>
        ''' <param name="keyName">The name or path of the subkey to create or open.
        ''' It can include a full path or a relative path based on the current context.</param>
        ''' <returns>
        ''' An instance of <see cref="IRegistryKey"/> representing the created or opened subkey.
        ''' Returns <c>Nothing</c> if the subkey could not be created or opened.
        ''' </returns>
        Function CreateSubKey(keyName As String) As IRegistryKey

        ''' <summary>
        ''' Opens a specified subkey in the registry hive with the given access mode.
        ''' </summary>
        ''' <param name="keyName">The name or path of the subkey to open.
        ''' This can be a full path to the subkey in the registry.</param>
        ''' <param name="accessMode">The access mode to apply when opening the subkey.
        ''' This determines the permissions available for the opened key.</param>
        ''' <returns>
        ''' An instance of <see cref="IRegistryKey"/> representing the opened subkey.
        ''' Returns <c>Nothing</c> if the subkey could not be found or opened.
        ''' </returns>
        Function OpenSubKey(keyName As String, accessMode As RegistryMode) As IRegistryKey
    End Interface
End Namespace
