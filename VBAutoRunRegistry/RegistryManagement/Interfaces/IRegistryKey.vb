Namespace RegistryManagement.Interfaces

    ''' <summary>
    ''' Provides an interface for manipulating registry key values.
    ''' This interface inherits from <see cref="IDisposable"/> to ensure proper resource management.
    ''' </summary>
    Public Interface IRegistryKey
        Inherits IDisposable

        ''' <summary>
        ''' Deletes a specified value from the registry key.
        ''' This method permanently removes the value associated with the specified name.
        ''' </summary>
        ''' <param name="valueName">The name of the value to delete.</param>
        Sub DeleteValue(valueName As String)

        ''' <summary>
        ''' Retrieves the value associated with the specified name from the registry key.
        ''' If the specified value does not exist, the provided default value is returned.
        ''' </summary>
        ''' <param name="valueName">The name of the value to retrieve.</param>
        ''' <param name="defaultValue">The value to return if the specified value does not exist.</param>
        ''' <returns>
        ''' The value associated with the specified name, or the default value if the name does not exist.
        ''' </returns>
        Function GetValue(valueName As String, defaultValue As Object) As Object

        ''' <summary>
        ''' Sets the specified value in the registry key.
        ''' If a value with the specified name already exists, it is overwritten.
        ''' </summary>
        ''' <param name="valueName">The name of the value to set.</param>
        ''' <param name="value">The value to set.</param>
        Sub SetValue(valueName As String, value As Object)
    End Interface
End Namespace
