Namespace RegistryManagement

    ''' <summary>
    ''' Provides methods for installing registry keys.
    ''' </summary>
    Friend Class RegistryInstaller
        Implements IRegistryInstaller

        ''' <summary>
        ''' The startup registry manager used to manage startup settings in the registry.
        ''' </summary>
        Private ReadOnly _startupRegistryManager As IStartupRegistryManager

        ''' <summary>
        ''' Initializes a new instance of the <see cref="RegistryInstaller"/> class.
        ''' </summary>
        ''' <param name="startupRegistryManager">The startup registry manager to use for managing startup settings in the registry.</param>
        Public Sub New(startupRegistryManager As IStartupRegistryManager)
            _startupRegistryManager = startupRegistryManager
        End Sub

        ''' <summary>
        ''' Installs the registry key.
        ''' </summary>
        ''' <returns>A task representing the asynchronous operation, with a Boolean indicating the success of the installation.</returns>
        Friend Async Function InstallRegistryKeyAsync() As Task(Of Boolean) Implements IRegistryInstaller.InstallRegistryKeyAsync
            Try
                Await Task.Run(Sub()
                    _startupRegistryManager.InstallStartupKey()
                End Sub)
                Return True
            Catch ex As Exception
                Throw
            End Try
        End Function
    End Class
End Namespace
