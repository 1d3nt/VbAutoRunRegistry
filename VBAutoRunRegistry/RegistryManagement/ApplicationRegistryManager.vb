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
        ''' The registry helper used to interact with the registry.
        ''' </summary>
        Private ReadOnly _registryHelper As IRegistryHelper

        ''' <summary>
        ''' Initializes a new instance of the <see cref="ApplicationRegistryManager"/> class.
        ''' </summary>
        ''' <param name="registryHelper">The registry helper to use for interacting with the registry.</param>
        Public Sub New(registryHelper As IRegistryHelper)
            _registryHelper = registryHelper
        End Sub

        ''' <summary>
        ''' Gets the application's settings registry key.
        ''' </summary>
        ''' <value>The registry key path for the application's settings.</value>
        Friend ReadOnly Property SettingsKey As String Implements IApplicationRegistryManager.SettingsKey
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
        Friend Function GetSetting(settingName As String) As String Implements IApplicationRegistryManager.GetSetting
            Try
                ValidateSettingName(settingName)
            Catch ex As Exception
                Throw
            End Try
            Dim value As String = Nothing
            Using key As IRegistryKey = _registryHelper.GetOrCreateKey(SettingsKey, RegistryMode.Read)
                Dim rawValue As Object = key.GetValue(settingName, Nothing)
                If rawValue IsNot Nothing Then
                    value = rawValue.ToString()
                End If
            End Using
            Return value
        End Function

        ''' <summary>
        ''' Validates the setting name.
        ''' </summary>
        ''' <param name="settingName">The setting name to validate.</param>
        ''' <exception cref="ArgumentNullException">Thrown when <paramref name="settingName"/> is null.</exception>
        ''' <exception cref="ArgumentException">Thrown when <paramref name="settingName"/> is invalid.</exception>
        Private Shared Sub ValidateSettingName(settingName As String)
            ArgumentNullException.ThrowIfNull(settingName)
            If IsInvalidName(settingName) Then
                Throw New ArgumentException("The setting name is invalid.", NameOf(settingName))
            End If
        End Sub

        ''' <summary>
        ''' Determines whether the specified setting name is invalid.
        ''' </summary>
        ''' <param name="settingName">The setting name to check.</param>
        ''' <returns><c>True</c> if the setting name is invalid; otherwise, <c>False</c>.</returns>
        Private Shared Function IsInvalidName(settingName As String) As Boolean
            Return settingName.Contains("/"c) OrElse settingName.Contains("\"c)
        End Function
    End Class
End Namespace
