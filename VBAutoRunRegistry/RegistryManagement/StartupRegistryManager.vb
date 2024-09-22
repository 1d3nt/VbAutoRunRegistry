Namespace RegistryManagement

    ''' <summary>
    ''' Manages the Windows auto-run registry key for configuring applications to start with Windows.
    ''' </summary>
    Friend Class StartupRegistryManager
        Implements IStartupRegistryManager

        ''' <summary>
        ''' The registry key path for auto-run programs in Windows.
        ''' </summary>
        Friend Const AutoRunKey As String = "Software\Microsoft\Windows\CurrentVersion\Run"

        ''' <summary>
        ''' The name of the registry value used to identify the application's auto-run setting.
        ''' </summary>
        Friend Const ValueName As String = "VbAutoRunRegistry"

        ''' <summary>
        ''' Gets or sets the registry hive where the startup settings will be managed.
        ''' </summary>
        Friend Property Hive As IRegistryHive

        ''' <summary>
        ''' The application registry manager used to manage application settings in the registry.
        ''' </summary>
        Private ReadOnly _applicationRegistryManager As IApplicationRegistryManager

        ''' <summary>
        ''' Initializes a new instance of the <see cref="StartupRegistryManager"/> class,
        ''' setting the default <see cref="IRegistryHive"/> implementation to <see cref="MicrosoftRegistryHive"/>.
        ''' </summary>
        ''' <param name="applicationRegistryManager">The application registry manager to use for managing application settings in the registry.</param>
        Public Sub New(applicationRegistryManager As IApplicationRegistryManager)
            Hive = New MicrosoftRegistryHive()
            _applicationRegistryManager = applicationRegistryManager
        End Sub

        ''' <summary>
        ''' Configures the application to automatically start with Windows by adding or updating 
        ''' the auto-run registry entry with the current executable's path.
        ''' </summary>
        ''' <returns>A Boolean indicating whether the registry key was installed successfully.</returns>
        Friend Function InstallStartupKey() As Boolean Implements IStartupRegistryManager.InstallStartupKey
            Dim currentPath As String = _applicationRegistryManager.GetSetting(ValueName)
            Dim applicationPath As String = Environment.ProcessPath
    
            If currentPath Is Nothing OrElse Not currentPath.Equals(applicationPath, StringComparison.OrdinalIgnoreCase) Then
                Try
                    Using key As IRegistryKey = GetStartupKey()
                        key.SetValue(ValueName, applicationPath)
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
        Friend Function UninstallStartupKey() As Boolean Implements IStartupRegistryManager.UninstallStartupKey
            Try
                Using key As IRegistryKey = GetStartupKey()
                    key.DeleteValue(ValueName)
                End Using
                Return True
            Catch ex As Exception
                Throw
            End Try
        End Function

        ''' <summary>
        ''' Retrieves the auto-run registry key for the current user. If the key does not exist, it is created.
        ''' </summary>
        ''' <returns>An <see cref="IRegistryKey"/> instance representing the auto-run registry key.</returns>
        Private Function GetStartupKey() As IRegistryKey
            Dim key As IRegistryKey = If(Hive.OpenSubKey(AutoRunKey, RegistryMode.ReadWrite), Hive.CreateSubKey(AutoRunKey))
            Return key
        End Function
    End Class
End Namespace
