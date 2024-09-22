Namespace Application.Interfaces
    ''' <summary>
    ''' Defines the contract for running and stopping the application.
    ''' </summary>
    Public Interface IAppRunner
        ''' <summary>
        ''' Runs the application.
        ''' </summary>
        ''' <returns>A task representing the asynchronous operation.</returns>
        Function RunAsync() As Task
        ''' <summary>
        ''' Stops the application.
        ''' </summary>
        ''' <returns>A task representing the asynchronous operation.</returns>
        Function StopAsync() As Task
    End Interface
End Namespace
