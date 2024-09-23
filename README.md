# VBAutoRunRegistry

## Overview

**VBAutoRunRegistry** is a VB.NET application designed to manage Windows startup registry entries for automatically launching applications. The project provides functionality for adding, removing, and managing registry keys to control the auto-run behavior of applications.

This solution leverages various interfaces and services to ensure that registry operations are performed in a modular, testable, and maintainable manner. It includes key components for handling user input, registry access, and service configuration using dependency injection principles.

## Key Features

- **Registry Management**: Create, read, update, and delete registry keys.
- **Auto-Run Configuration**: Automatically configure applications to run at Windows startup.
- **Dependency Injection**: Built with dependency injection principles to ensure modularity and maintainability.
- **Error Handling**: Exception handling for robust registry operations.
- **Unit Testable**: Designed with interfaces to support unit testing.

## Technologies Used

- **VB.NET**
- **.NET Core/Framework**
- **Windows Registry API** (via Microsoft.Win32 namespace)
- **Dependency Injection**
- **Unit Testing Frameworks** (e.g., MSTest, NUnit)

## Project Structure

```bash
|-- Application
|   |-- Interfaces
|   |   +-- IAppRunner.vb
|   |-- AppRunner.vb
|   |-- ProcessExitHandler.vb
|   +-- ServiceConfigurator.vb
|
|-- RegistryManagement
|   |-- Interfaces
|   |   |-- IApplicationRegistryManager.vb
|   |   |-- IRegistryHelper.vb
|   |   |-- IRegistryHive.vb
|   |   |-- IRegistryInstaller.vb
|   |   |-- IRegistryKey.vb
|   |   |-- IRegistryUninstaller.vb
|   |   +-- IStartupRegistryManager.vb
|   |-- RegistryAccess
|   |   +-- RegistryMode.vb
|   |-- ApplicationRegistryManager.vb
|   |-- MicrosoftRegistryHive.vb
|   |-- MicrosoftRegistryKey.vb
|   |-- RegistryHelper.vb
|   |-- RegistryInstaller.vb
|   |-- RegistryUninstaller.vb
|   +-- StartupRegistryManager.vb
|
|-- Utilities
|   |-- Interfaces
|   |   |-- IUserInputChecker.vb
|   |   |-- IUserInputReader.vb
|   |   +-- IUserPrompter.vb
|   |-- UserInputChecker.vb
|   |-- UserInputReader.vb
|   +-- UserPrompter.vb
|
|-- GlobalAttribute.vb
+-- Program.vb
```

## Getting Started

### Prerequisites

- **.NET SDK**: Ensure that the .NET SDK (Core or Framework) is installed on your machine. You can download it [here](https://dotnet.microsoft.com/download).
- **Visual Studio**: While any IDE can be used, the project is optimized for Visual Studio 2019/2022.

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/YourUsername/VBAutoRunRegistry.git
   cd VBAutoRunRegistry
   ```
2. Open the solution in Visual Studio and restore the NuGet packages.
