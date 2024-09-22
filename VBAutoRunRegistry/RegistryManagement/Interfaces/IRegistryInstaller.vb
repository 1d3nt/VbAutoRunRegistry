Namespace RegistryManagement.Interfaces

    ''' <summary>
    ''' Provides methods for installing registry keys.
    ''' </summary>
    Public Interface IRegistryInstaller

        ''' <summary>
        ''' Installs the registry key.
        ''' </summary>
        ''' <returns>A task representing the asynchronous operation.</returns>
        Function InstallRegistryKeyAsync() As Task(of Boolean)
    End Interface
End Namespace
