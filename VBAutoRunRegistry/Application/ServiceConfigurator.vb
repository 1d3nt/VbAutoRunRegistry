Imports VBAutoRunRegistry.Utilities

Namespace Application
    ''' <summary>
    ''' Configures the services for dependency injection.
    ''' </summary>
    ''' <remarks>
    ''' The <see cref="ServiceConfigurator"/> class provides methods to configure and register services
    ''' for dependency injection. It creates a new instance of the <see cref="ServiceCollection"/> class,
    ''' registers various services and their corresponding implementations, and builds an <see cref="IServiceProvider"/>
    ''' which can be used to resolve services at runtime.
    ''' </remarks>
    Friend Class ServiceConfigurator
        ''' <summary>
        ''' Registers all services required by the application.
        ''' </summary>
        ''' <param name="services">
        ''' The service collection to add the services to.
        ''' </param>
        ''' <remarks>
        ''' This method calls individual methods to register different categories of services:
        '''   <item>
        '''     <description>
        '''       <see cref="RegisterUserInputServices"/> to register user input services.
        '''     </description>
        '''   </item>
        '''   <item>
        '''     <description>
        '''       <see cref="RegisterRegistryManagementServices"/> to register registry management services.
        '''     </description>
        '''   </item>
        '''   <item>
        '''     <description>
        '''       <see cref="RegisterAppRunnerService"/> to register the application runner service.
        '''     </description>
        '''   </item>
        ''' </remarks>
        Private Shared Sub RegisterServices(services As IServiceCollection)
            RegisterUserInputServices(services)
            RegisterRegistryManagementServices(services)
            RegisterAppRunnerService(services) 
        End Sub

        ''' <summary>
        ''' Registers user input services.
        ''' </summary>
        ''' <param name="services">
        ''' The service collection to which the user input services are added.
        ''' </param>
        ''' <remarks>
        ''' This method registers the following user input services:
        ''' <list type="bullet">
        '''   <item>
        '''     <description>
        '''       <see cref="IUserInputReader"/> is implemented by <see cref="UserInputReader"/>.
        '''     </description>
        '''   </item>
        '''   <item>
        '''     <description>
        '''       <see cref="IUserPrompter"/> is implemented by <see cref="UserPrompter"/>.
        '''     </description>
        '''   </item>
        '''   <item>
        '''     <description>
        '''       <see cref="IUserInputChecker"/> is implemented by <see cref="UserInputChecker"/>.
        '''     </description>
        '''   </item>
        ''' </list>
        ''' </remarks>
        Private Shared Sub RegisterUserInputServices(services As IServiceCollection)
            Dim userInputServices As New Dictionary(Of Type, Type) From {
                {GetType(IUserInputReader), GetType(UserInputReader)},
                {GetType(IUserPrompter), GetType(UserPrompter)},
                {GetType(IUserInputChecker), GetType(UserInputChecker)}
            }
            AddServices(services, userInputServices)
        End Sub

        ''' <summary>
        ''' Registers registry management services.
        ''' </summary>
        ''' <param name="services">
        ''' The service collection to which the registry management services are added.
        ''' </param>
        ''' <remarks>
        ''' This method registers the following registry management services:
        ''' <list type="bullet">
        '''   <item>
        '''     <description>
        '''       <see cref="IRegistryInstaller"/> is implemented by <see cref="RegistryInstaller"/>.
        '''     </description>
        '''   </item>
        '''   <item>
        '''     <description>
        '''       <see cref="IRegistryUninstaller"/> is implemented by <see cref="RegistryUninstaller"/>.
        '''     </description>
        '''   </item>
        '''   <item>
        '''     <description>
        '''       <see cref="IApplicationRegistryManager"/> is implemented by <see cref="ApplicationRegistryManager"/>.
        '''     </description>
        '''   </item>
        '''   <item>
        '''     <description>
        '''       <see cref="IStartupRegistryManager"/> is implemented by <see cref="StartupRegistryManager"/>.
        '''     </description>
        '''   </item>
        '''   <item>
        '''     <description>
        '''       <see cref="IRegistryHive"/> is implemented by <see cref="MicrosoftRegistryHive"/>.
        '''     </description>
        '''   </item>
        ''' </list>
        ''' </remarks>
        Private Shared Sub RegisterRegistryManagementServices(services As IServiceCollection)
            Dim registryManagementServices As New Dictionary(Of Type, Type) From {
                {GetType(IRegistryInstaller), GetType(RegistryInstaller)},
                {GetType(IRegistryUninstaller), GetType(RegistryUninstaller)},
                {GetType(IApplicationRegistryManager), GetType(ApplicationRegistryManager)},
                {GetType(IStartupRegistryManager), GetType(StartupRegistryManager)},
                {GetType(IRegistryHive), GetType(MicrosoftRegistryHive)}
            }
            AddServices(services, registryManagementServices)
        End Sub

        ''' <summary>
        ''' Registers the application runner service.
        ''' </summary>
        ''' <param name="services">
        ''' The service collection to which the application runner service is added.
        ''' </param>
        ''' <remarks>
        ''' This method registers the <see cref="IAppRunner"/> service, which is implemented by <see cref="AppRunner"/>.
        ''' </remarks>
        Private Shared Sub RegisterAppRunnerService(services As IServiceCollection)
            services.AddTransient(Of IAppRunner, AppRunner)()
        End Sub

        ''' <summary>
        ''' Adds the specified services to the service collection.
        ''' </summary>
        ''' <param name="services">
        ''' The service collection to which the services are added.
        ''' </param>
        ''' <param name="serviceRegistrations">
        ''' The dictionary containing service registrations.
        ''' </param>
        ''' <remarks>
        ''' This method iterates over the provided dictionary of service registrations and adds each service
        ''' to the <paramref name="services"/> collection with a transient lifetime.
        ''' </remarks>
        Private Shared Sub AddServices(services As IServiceCollection, serviceRegistrations As Dictionary(Of Type, Type))
            For Each kvp As KeyValuePair(Of Type, Type) In serviceRegistrations
                services.AddTransient(kvp.Key, kvp.Value)
            Next
        End Sub

        ''' <summary>
        ''' Configures the services for dependency injection.
        ''' </summary>
        ''' <returns>
        ''' An <see cref="IServiceProvider"/> that provides the configured services.
        ''' </returns>
        ''' <remarks>
        ''' The <see cref="ConfigureServices"/> method creates a new instance of the <see cref="ServiceCollection"/>
        ''' and registers various services and their corresponding implementations by calling the <see cref="RegisterServices"/> method.
        ''' The method then builds and returns an <see cref="IServiceProvider"/>.
        ''' </remarks>
        Friend Shared Function ConfigureServices() As IServiceProvider
            Dim services As New ServiceCollection()
            RegisterServices(services)
            Return services.BuildServiceProvider()
        End Function
    End Class
End Namespace
