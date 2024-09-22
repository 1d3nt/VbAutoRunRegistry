Imports System.Diagnostics.CodeAnalysis

''' <author>
''' Sam (ident)
''' Twitter: <see href="https://twitter.com/1d3nt">https://twitter.com/1d3nt</see>
''' GitHub: <see href="https://github.com/1d3nt">https://github.com/1d3nt</see>
''' Email: <see href="mailto:ident@simplecoders.com">ident@simplecoders.com</see>
''' VBForums: <see href="https://www.vbforums.com/member.php?113656-ident">https://www.vbforums.com/member.php?113656-ident</see>
''' ORCID: <see href="https://orcid.org/0009-0007-1363-3308">https://orcid.org/0009-0007-1363-3308</see>
''' </author>
''' <date>22/09/2024</date>
''' <version>1.0.0</version>
''' <license>Creative Commons Attribution 4.0 International (CC BY 4.0) - See LICENSE.md for details</license>
''' <summary>
''' The entry point for the VBAutoRunRegistry application.
''' </summary>
''' <remarks>
''' This module contains the <see cref="Main"/> method, which sets up and starts the application host.
''' 
''' This project focuses on managing the auto-run registry settings for Windows applications.
''' Contributions and enhancements are welcomed to further extend the functionality and improve the integration.
''' 
''' Just a hobby programmer that enjoys working with registry interactions and system-level APIs.
'''
''' As a tribute to AtmaWeapon(Xtreme VB Talk), whose guidance helped me understand registry interactions and the true essence of programming, 
''' moving beyond merely copying and pasting code. This project reflects that learning journey.
''' </remarks>
Module Program

    ''' <summary>
    ''' Entry point for the console application. This application manages user interactions to configure auto-run settings in the registry. 
    ''' It prompts users for input to determine whether to proceed with various setup tasks related to auto-run configurations.
    ''' </summary>
    ''' <param name="args">Command-line arguments (not used in this implementation).</param>
    ''' <remarks>
    ''' The <paramref name="args"/> parameter is included to adhere to the standard signature for a console application's
    ''' <see cref="Main"/> method. While this parameter is not utilized in the current implementation, including it follows 
    ''' best practices and ensures consistency with the conventional entry point signature.
    ''' 
    ''' This method initializes the <see cref="IServiceProvider"/> by calling <see cref="ServiceConfigurator.ConfigureServices"/> 
    ''' to configure and provide the required services for the application. The services are then passed into the <see cref="AppRunner"/> 
    ''' class, which handles the execution of the core logic for the auto-run registry settings.
    ''' 
    ''' The method calls <see cref="AppRunner.RunAsync"/> and waits for its completion using 
    ''' <see cref="M:System.Threading.Tasks.Task.GetAwaiter().GetResult"/>. This approach ensures that the asynchronous operations 
    ''' in <see cref="AppRunner.RunAsync"/> complete before the application exits. Unlike <see cref="M:System.Threading.Tasks.Task.Wait()"/>, 
    ''' which can cause deadlocks in some cases, the use of <see cref="M:System.Threading.Tasks.Task.GetAwaiter().GetResult"/> 
    ''' is a safer alternative in a console application's synchronous entry point.
    ''' 
    ''' The variable <c>serviceProvider</c> manages service dependencies for the application. 
    ''' The <c>appRunner</c> variable is responsible for orchestrating the auto-run setup process by executing the application's core logic.
    ''' </remarks>
    <SuppressMessage("Style", "IDE0060:Remove unused parameter", Justification:="Standard Main method parameter signature.")>
    Sub Main(args As String())
        Try
            Dim serviceProvider As IServiceProvider = ServiceConfigurator.ConfigureServices()
            Dim appRunner = serviceProvider.GetService(Of IAppRunner)()
            Dim processExitHandler As New ProcessExitHandler(appRunner)
            appRunner.RunAsync().GetAwaiter().GetResult()
            GC.KeepAlive(processExitHandler)
        Catch ex As Exception
            Console.WriteLine($"An error occurred: {ex.Message}")
        End Try
    End Sub
End Module