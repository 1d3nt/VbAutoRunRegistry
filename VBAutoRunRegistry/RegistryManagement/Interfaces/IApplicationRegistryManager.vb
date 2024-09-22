Namespace RegistryManagement.Interfaces
    ''' <summary>
    ''' Defines methods for managing application settings in the registry.
    ''' </summary>
    Public Interface IApplicationRegistryManager
        ''' <summary>
        ''' Gets the application's settings registry key.
        ''' </summary>
        ''' <value>The registry key path for the application's settings.</value>
        ReadOnly Property SettingsKey As String
        ''' <summary>
        ''' Returns the value of the setting with the specified name. If the setting does not exist, returns Nothing.
        ''' </summary>
        ''' <param name="settingName">The name of the setting to retrieve.</param>
        ''' <returns>The value of the setting, or Nothing if the setting does not exist.</returns>
        ''' <exception cref="ArgumentNullException">Thrown when <paramref name="settingName"/> is null.</exception>
        ''' <exception cref="ArgumentException">Thrown when <paramref name="settingName"/> is invalid.</exception>
        Function GetSetting(settingName As String) As String
    End Interface
End Namespace
