Namespace RegistryManagement
    ''' <summary>
    ''' Provides methods for managing application settings in the registry.
    ''' </summary>
    Friend Class ApplicationRegistryManager
        Implements IApplicationRegistryManager
        ''' <summary>
        ''' The default registry key for the application.
        ''' </summary>
        Private Const DefaultApplicationRegistryKey As String = "Software\VbAutoRunRegistry"
        ''' <summary>
        ''' Represents the registry hive that will be used.
        ''' </summary>
        Private ReadOnly _hive As IRegistryHive
        ''' <summary>
        ''' Initializes a new instance of the <see cref="ApplicationRegistryManager"/> class.
        ''' </summary>
        ''' <param name="hive">The registry hive to use for accessing the registry.</param>
        Public Sub New(hive As IRegistryHive)
            _hive = hive
        End Sub
        ''' <summary>
        ''' Gets the application's settings registry key.
        ''' </summary>
        ''' <value>The registry key path for the application's settings.</value>
        Public ReadOnly Property SettingsKey As String Implements IApplicationRegistryManager.SettingsKey
            Get
                Return DefaultApplicationRegistryKey
            End Get
        End Property
        ''' <summary>
        ''' Returns the value of the setting with the specified name. If the setting does not exist, returns Nothing.
        ''' </summary>
        ''' <param name="settingName">The name of the setting to retrieve.</param>
        ''' <returns>The value of the setting, or Nothing if the setting does not exist.</returns>
        ''' <exception cref="ArgumentNullException">Thrown when <paramref name="settingName"/> is null.</exception>
        ''' <exception cref="ArgumentException">Thrown when <paramref name="settingName"/> is invalid.</exception>
        Public Function GetSetting(settingName As String) As String Implements IApplicationRegistryManager.GetSetting
            ArgumentNullException.ThrowIfNull(settingName)
            If IsInvalidName(settingName) Then
                Throw New ArgumentException("The setting name is invalid.", NameOf(settingName))
            End If
            Dim value As String = Nothing
            Using key As IRegistryKey = GetRegistryKey(RegistryMode.Read)
                Dim rawValue As Object = key.GetValue(settingName, Nothing)
                If rawValue IsNot Nothing Then
                    value = rawValue.ToString()
                End If
            End Using
            Return value
        End Function
        ''' <summary>
        ''' Creates the application settings registry key and returns it.
        ''' </summary>
        ''' <returns>An instance of <see cref="IRegistryKey"/> representing the created registry key.</returns>
        ''' <exception cref="InvalidOperationException">Thrown when the registry key could not be created.</exception>
        Private Function CreateSettingsKey() As IRegistryKey
            Dim key As IRegistryKey = _hive.CreateSubKey(SettingsKey)
            If key Is Nothing Then
                Throw New InvalidOperationException("Could not create application settings registry key.")
            End If
            Return key
        End Function
        ''' <summary>
        ''' Gets a reference to the application settings registry key. If the key does not exist, it is created.
        ''' </summary>
        ''' <param name="accessMode">The access mode for the registry key.</param>
        ''' <returns>An instance of <see cref="IRegistryKey"/> representing the registry key.</returns>
        ''' <exception cref="InvalidOperationException">Thrown when the registry key could not be created.</exception>
        Private Function GetRegistryKey(accessMode As RegistryMode) As IRegistryKey
            Dim key As IRegistryKey = If(_hive.OpenSubKey(SettingsKey, accessMode), CreateSettingsKey())
            Return key
        End Function
        ''' <summary>
        ''' Determines whether the specified setting name is invalid.
        ''' </summary>
        ''' <param name="settingName">The setting name to check.</param>
        ''' <returns><c>True</c> if the setting name is invalid; otherwise, <c>False</c>.</returns>
        Private shared Function IsInvalidName(settingName As String) As Boolean
            Return settingName.Contains("/"c) OrElse settingName.Contains("\"c)
        End Function
    End Class
End Namespace
