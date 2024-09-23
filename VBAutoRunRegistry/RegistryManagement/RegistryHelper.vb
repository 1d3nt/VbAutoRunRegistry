Namespace RegistryManagement

    ''' <summary>
    ''' Provides helper methods for interacting with the Windows registry.
    ''' </summary>
    Friend Class RegistryHelper
        Implements IRegistryHelper

        ''' <summary>
        ''' The registry hive where the settings will be managed.
        ''' </summary>
        Private ReadOnly _hive As IRegistryHive

        ''' <summary>
        ''' Initializes a new instance of the <see cref="RegistryHelper"/> class.
        ''' </summary>
        ''' <param name="hive">The registry hive to use for managing registry keys.</param>
        Public Sub New(hive As IRegistryHive)
            _hive = hive
        End Sub

        ''' <summary>
        ''' Retrieves the registry key for the specified path. If the key does not exist, it is created.
        ''' </summary>
        ''' <param name="path">The path of the registry key.</param>
        ''' <param name="accessMode">The access mode for the registry key.</param>
        ''' <returns>An <see cref="IRegistryKey"/> instance representing the registry key.</returns>
        Friend Function GetOrCreateKey(path As String, accessMode As RegistryMode) As IRegistryKey Implements IRegistryHelper.GetOrCreateKey
            Dim key As IRegistryKey = If(_hive.OpenSubKey(path, accessMode), _hive.CreateSubKey(path))
            Return key
        End Function

        ''' <summary>
        ''' Sets the value of the specified registry key.
        ''' </summary>
        ''' <param name="key">The registry key.</param>
        ''' <param name="valueName">The name of the value to set.</param>
        ''' <param name="value">The value to set.</param>
        Friend Sub SetValue(key As IRegistryKey, valueName As String, value As String) Implements IRegistryHelper.SetValue
            key.SetValue(valueName, value)
        End Sub

        ''' <summary>
        ''' Deletes the value of the specified registry key.
        ''' </summary>
        ''' <param name="key">The registry key.</param>
        ''' <param name="valueName">The name of the value to delete.</param>
        Friend Sub DeleteValue(key As IRegistryKey, valueName As String) Implements IRegistryHelper.DeleteValue
            key.DeleteValue(valueName)
        End Sub
    End Class
End Namespace
