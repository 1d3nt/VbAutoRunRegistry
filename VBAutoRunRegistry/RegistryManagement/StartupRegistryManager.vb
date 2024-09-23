Namespace RegistryManagement

    ''' <summary>
    ''' Manages the Windows auto-run registry key for configuring applications to start with Windows.
    ''' </summary>
    Friend Class StartupRegistryManager
        Implements IStartupRegistryManager

        ''' <summary>
        ''' The registry key path for auto-run programs in Windows.
        ''' </summary>
        Private Const AutoRunKey As String = "Software\Microsoft\Windows\CurrentVersion\Run"
        ''' <summary>
        ''' The name of the registry value used to identify the application's auto-run setting.
        ''' 
        ''' </summary>
        Private Const ValueName As String = "VbAutoRunRegistry"

        ''' <summary>
        ''' The registry helper used to interact with the registry.
        ''' </summary>
        Private ReadOnly _registryHelper As IRegistryHelper

        ''' <summary>
        ''' The application registry manager used to manage application settings in the registry.
        ''' </summary>
        Private ReadOnly _applicationRegistryManager As IApplicationRegistryManager

        ''' <summary>
        ''' Initializes a new instance of the <see cref="StartupRegistryManager"/> class.
        ''' </summary>
        ''' <param name="registryHelper">The registry helper to use for interacting with the registry.</param>
        ''' <param name="applicationRegistryManager">The application registry manager to use for managing application settings in the registry.</param>
        Public Sub New(registryHelper As IRegistryHelper, applicationRegistryManager As IApplicationRegistryManager)
            _registryHelper = registryHelper
            _applicationRegistryManager = applicationRegistryManager
        End Sub

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
        Friend Function InstallStartupKey() As Boolean Implements IStartupRegistryManager.InstallStartupKey
            Dim currentPath As String
            Try
                currentPath = _applicationRegistryManager.GetSetting(ValueName)
            Catch ex As Exception
                Throw
            End Try
            Dim applicationPath As String = Environment.ProcessPath
            If currentPath Is Nothing OrElse Not currentPath.Equals(applicationPath, StringComparison.OrdinalIgnoreCase) Then
                Try
                    Using key As IRegistryKey = _registryHelper.GetOrCreateKey(AutoRunKey, RegistryMode.ReadWrite)
                        _registryHelper.SetValue(key, ValueName, applicationPath)
                    End Using
                Catch ex As Exception
                    Throw
                End Try
            End If
            Return True
        End Function

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
        Friend Function UninstallStartupKey() As Boolean Implements IStartupRegistryManager.UninstallStartupKey
            Try
                Using key As IRegistryKey = _registryHelper.GetOrCreateKey(AutoRunKey, RegistryMode.ReadWrite)
                    _registryHelper.DeleteValue(key, ValueName)
                End Using
                Return True
            Catch ex As Exception
                Throw
            End Try
        End Function
    End Class
End Namespace
