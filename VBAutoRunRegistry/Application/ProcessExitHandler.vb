Namespace Application

    ''' <summary>
    ''' Handles the process exit event to perform cleanup tasks.
    ''' </summary>
    Public Class ProcessExitHandler
        ''' <summary>
        ''' The IAppRunner instance used to stop the application.
        ''' </summary>
        Private ReadOnly _appRunner As IAppRunner
        ''' <summary>
        ''' Initializes a new instance of the <see cref="ProcessExitHandler"/> class.
        ''' </summary>
        ''' <param name="appRunner">The IAppRunner instance to stop the application.</param>
        Public Sub New(appRunner As IAppRunner)
            _appRunner = appRunner
            AddHandler AppDomain.CurrentDomain.ProcessExit, AddressOf OnProcessExit
        End Sub
        ''' <summary>
        ''' Handles the ProcessExit event to uninstall the registry key.
        ''' </summary>
        ''' <param name="sender">The source of the event.</param>
        ''' <param name="e">An <see cref="EventArgs"/> that contains the event data.</param>
        Private Sub OnProcessExit(sender As Object, e As EventArgs)
            _appRunner.StopAsync().GetAwaiter().GetResult()
        End Sub
    End Class
End Namespace
