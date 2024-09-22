Namespace RegistryManagement.Interfaces
    ''' <summary>
    ''' Defines methods for managing the Windows auto-run registry key.
    ''' </summary>
    Public Interface IStartupRegistryManager
        ''' <summary>
        ''' Configures the application to automatically start with Windows by adding or updating 
        ''' the auto-run registry entry with the current executable's path.
        ''' </summary>
        Sub InstallStartupKey()
        ''' <summary>
        ''' Removes the application's entry from the Windows startup registry, 
        ''' preventing it from starting automatically with Windows.
        ''' </summary>
        Sub UninstallStartupKey()
    End Interface
End Namespace
