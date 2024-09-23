Namespace RegistryManagement

    ''' <summary>
    ''' Provides methods for uninstalling registry keys.
    ''' </summary>
    Friend Class RegistryUninstaller
        Implements IRegistryUninstaller

        ''' <summary>
        ''' The startup registry manager used to manage startup settings in the registry.
        ''' </summary>
        Private ReadOnly _startupRegistryManager As IStartupRegistryManager

        ''' <summary>
        ''' Initializes a new instance of the <see cref="RegistryUninstaller"/> class.
        ''' </summary>
        ''' <param name="startupRegistryManager">The startup registry manager to use for managing startup settings in the registry.</param>
        Public Sub New(startupRegistryManager As IStartupRegistryManager)
            _startupRegistryManager = startupRegistryManager
        End Sub

        ''' <summary>
        ''' Uninstalls the registry key.
        ''' </summary>
        ''' <returns>A task representing the asynchronous operation.</returns>
        Friend Async Function UninstallRegistryKeyAsync() As Task Implements IRegistryUninstaller.UninstallRegistryKeyAsync
            Try
                Await Task.Run(Sub()
                    _startupRegistryManager.UninstallStartupKey()
                End Sub)
            Catch ex As Exception
                Throw
            End Try
        End Function
    End Class
End Namespace
