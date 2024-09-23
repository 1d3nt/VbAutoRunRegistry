Namespace RegistryManagement.Interfaces

    ''' <summary>
    ''' Defines methods for managing the Windows auto-run registry key.
    ''' </summary>
    Public Interface IStartupRegistryManager

        ''' <summary>
        ''' Configures the application to automatically start with Windows by adding or updating 
        ''' the auto-run registry entry with the current executable's path.
        ''' </summary>
        ''' <returns>A Boolean indicating whether the registry key was installed successfully.</returns>
        ''' <exception cref="Exception">Thrown when the registry entry could not be created or updated.</exception>
        ''' <remarks>
        ''' This method checks if the application is already configured to start with Windows.
        ''' If not, it adds or updates the corresponding entry in the registry.
        ''' </remarks>
        Function InstallStartupKey() As Boolean

        ''' <summary>
        ''' Removes the application's entry from the Windows startup registry, 
        ''' preventing it from starting automatically with Windows.
        ''' </summary>
        ''' <returns>A Boolean indicating whether the registry key was uninstalled successfully.</returns>
        ''' <exception cref="Exception">Thrown when the registry entry could not be deleted.</exception>
        ''' <remarks>
        ''' This method attempts to remove the application's entry from the startup registry.
        ''' If the operation fails, it throws an exception.
        ''' </remarks>
        Function UninstallStartupKey() As Boolean
    End Interface
End Namespace
