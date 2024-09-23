Namespace RegistryManagement.Interfaces

    ''' <summary>
    ''' Provides methods for interacting with the Windows registry.
    ''' </summary>
    Public Interface IRegistryHelper

        ''' <summary>
        ''' Retrieves the registry key for the specified path with the specified access mode.
        ''' If the key does not exist, it is created.
        ''' </summary>
        ''' <param name="path">The path of the registry key.</param>
        ''' <param name="accessMode">The mode in which to open the registry key.</param>
        ''' <returns>An <see cref="IRegistryKey"/> instance representing the registry key.</returns>
        ''' <exception cref="Exception">Thrown when the registry key cannot be opened or created.</exception>
        ''' <remarks>
        ''' This method attempts to open an existing registry key at the specified path with the specified access mode.
        ''' If the key does not exist, it creates a new subkey at that path.
        ''' </remarks>
        Function GetOrCreateKey(path As String, accessMode As RegistryMode) As IRegistryKey

        ''' <summary>
        ''' Sets the value of the specified registry key.
        ''' </summary>
        ''' <param name="key">The registry key.</param>
        ''' <param name="valueName">The name of the value to set.</param>
        ''' <param name="value">The value to set.</param>
        ''' <exception cref="Exception">Thrown when the value cannot be set.</exception>
        ''' <remarks>
        ''' This method assigns the specified value to the named entry in the provided registry key. 
        ''' It throws an exception if the operation fails.
        ''' </remarks>
        Sub SetValue(key As IRegistryKey, valueName As String, value As String)

        ''' <summary>
        ''' Deletes the value of the specified registry key.
        ''' </summary>
        ''' <param name="key">The registry key.</param>
        ''' <param name="valueName">The name of the value to delete.</param>
        ''' <exception cref="Exception">Thrown when the value cannot be deleted.</exception>
        ''' <remarks>
        ''' This method removes the specified value from the given registry key. 
        ''' If the operation fails, it throws an exception.
        ''' </remarks>
        Sub DeleteValue(key As IRegistryKey, valueName As String)
    End Interface
End Namespace
