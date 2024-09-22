Namespace RegistryManagement.Interfaces

    ''' <summary>
    ''' Provides methods for uninstalling registry keys.
    ''' </summary>
    Public Interface IRegistryUninstaller

        ''' <summary>
        ''' Uninstalls the registry key.
        ''' </summary>
        ''' <returns>A task representing the asynchronous operation.</returns>
        Function UninstallRegistryKeyAsync() As Task
    End Interface
End Namespace
