Namespace Application

    ''' <summary>
    ''' Represents an application that uses dependency injection to obtain services and perform operations.
    ''' </summary>
    ''' <remarks>
    ''' The <see cref="AppRunner"/> class relies on dependency injection to obtain an <see cref="IServiceProvider"/> 
    ''' which is used to retrieve services throughout the application. The main functionality of the class is to
    ''' run the application's core logic, including installing and uninstalling registry keys.
    ''' </remarks>
    Friend Class AppRunner
        Implements IAppRunner

        ''' <summary>
        ''' The service provider used for retrieving services.
        ''' </summary>
        Private ReadOnly _serviceProvider As IServiceProvider
        ''' <summary>
        ''' The user input reader used to read user input from the console.
        ''' </summary>
        Private ReadOnly _userInputReader As IUserInputReader
        ''' <summary>
        ''' The user prompter used to display messages and prompts to the user.
        ''' </summary>
        Private ReadOnly _userPrompter As IUserPrompter

        ''' <summary>
        ''' Initializes a new instance of the <see cref="AppRunner"/> class.
        ''' </summary>
        ''' <param name="serviceProvider">An instance of <see cref="IServiceProvider"/> used to resolve dependencies and obtain services.</param>
        ''' <param name="userInputReader">An instance of <see cref="IUserInputReader"/> for reading user input.</param>
        ''' <param name="userPrompter">An instance of <see cref="IUserPrompter"/> for displaying messages.</param>
        Public Sub New(serviceProvider As IServiceProvider, userInputReader As IUserInputReader, userPrompter As IUserPrompter)
            _serviceProvider = serviceProvider
            _userInputReader = userInputReader
            _userPrompter = userPrompter
        End Sub

        ''' <summary>
        ''' Runs the application.
        ''' </summary>
        ''' <remarks>
        ''' The <see cref="RunAsync"/> method retrieves the <see cref="IUserInputChecker"/> from the service provider,
        ''' uses it to determine if the user wants to proceed with installing the registry key, and prints the result to the console.
        ''' The method then waits for the user to press Enter before terminating.
        ''' </remarks>
        Friend Async Function RunAsync() As Task Implements IAppRunner.RunAsync
            Dim userInputChecker = _serviceProvider.GetService(Of IUserInputChecker)()
            Dim shouldProceed = userInputChecker.ShouldProceed()
            If shouldProceed Then
                Try
                    Dim installationSuccess = Await InstallRegistryKeyAsync()
                    _userPrompter.Prompt($"Registry key installation success: {installationSuccess}")
                Catch ex As Exception
                    _userPrompter.Prompt($"Registry key installation failed: {ex.Message}")
                End Try
            Else
                _userPrompter.Prompt("Registry key installation was not performed.")
            End If
            _userInputReader.ReadInput()
        End Function

        ''' <summary>
        ''' Stops the application.
        ''' </summary>
        ''' <remarks>
        ''' The <see cref="StopAsync"/> method uninstalls the registry key asynchronously.
        ''' </remarks>
        Friend Async Function StopAsync() As Task Implements IAppRunner.StopAsync
            Try
                Dim uninstallationSuccess = Await UninstallRegistryKeyAsync()
                _userPrompter.Prompt($"Registry key uninstallation success: {uninstallationSuccess}")
            Catch ex As Exception
                _userPrompter.Prompt($"Registry key uninstallation failed: {ex.Message}")
            End Try
        End Function

        ''' <summary>
        ''' Installs the registry key.
        ''' </summary>
        ''' <remarks>
        ''' This method handles the logic for installing the registry key.
        ''' </remarks>
        Private Async Function InstallRegistryKeyAsync() As Task(Of Boolean)
            Dim registryInstaller = _serviceProvider.GetService(Of IRegistryInstaller)()
            Try
                Await registryInstaller.InstallRegistryKeyAsync()
                Return True
            Catch ex As Exception
                Throw
            End Try
        End Function

        ''' <summary>
        ''' Uninstalls the registry key.
        ''' </summary>
        ''' <remarks>
        ''' This method handles the logic for uninstalling the registry key.
        ''' </remarks>
        Private Async Function UninstallRegistryKeyAsync() As Task(Of Boolean)
            Dim registryUninstaller = _serviceProvider.GetService(Of IRegistryUninstaller)()
            Try
                Await registryUninstaller.UninstallRegistryKeyAsync()
                Return True
            Catch ex As Exception
                Throw
            End Try
        End Function
    End Class
End Namespace
