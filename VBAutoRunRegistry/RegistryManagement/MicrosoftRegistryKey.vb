Namespace RegistryManagement

    ''' <summary>
    ''' Provides access to a registry key and implements the <see cref="IRegistryKey"/> interface.
    ''' This class also implements <see cref="IDisposable"/> to manage resource cleanup.
    ''' </summary>
    Friend Class MicrosoftRegistryKey
        Implements IRegistryKey, IDisposable

        ''' <summary>
        ''' The root registry key.
        ''' </summary>
        Private ReadOnly _root As RegistryKey

        ''' <summary>
        ''' Indicates whether the object has been disposed.
        ''' </summary>
        Private _isDisposed As Boolean

        ''' <summary>
        ''' Initializes a new instance of the <see cref="MicrosoftRegistryKey"/> class.
        ''' </summary>
        ''' <param name="root">The root registry key.</param>
        ''' <exception cref="ArgumentNullException">Thrown when <paramref name="root"/> is null.</exception>
        Friend Sub New(root As RegistryKey)
            _root = root
        End Sub

        ''' <summary>
        ''' Deletes the specified value from the registry key.
        ''' </summary>
        ''' <param name="valueName">The name of the value to delete.</param>
        ''' <exception cref="ObjectDisposedException">Thrown when the object has been disposed.</exception>
        Friend Sub DeleteValue(valueName As String) Implements IRegistryKey.DeleteValue
            ObjectDisposedException.ThrowIf(_isDisposed, NameOf(MicrosoftRegistryKey))
            _root.DeleteValue(valueName, False)
        End Sub

        ''' <summary>
        ''' Sets the specified value in the registry key.
        ''' </summary>
        ''' <param name="valueName">The name of the value to set.</param>
        ''' <param name="value">The value to set.</param>
        ''' <exception cref="ObjectDisposedException">Thrown when the object has been disposed.</exception>
        Friend Sub SetValue(valueName As String, value As Object) Implements IRegistryKey.SetValue
            ObjectDisposedException.ThrowIf(_isDisposed, NameOf(MicrosoftRegistryKey))
            _root.SetValue(valueName, value, RegistryValueKind.String)
        End Sub

        ''' <summary>
        ''' Gets the value associated with the specified name in the registry key.
        ''' </summary>
        ''' <param name="valueName">The name of the value to retrieve.</param>
        ''' <param name="defaultValue">The value to return if the specified name does not exist.</param>
        ''' <returns>
        ''' The value associated with the specified name, or <paramref name="defaultValue"/> if the name does not exist.
        ''' </returns>
        ''' <exception cref="ObjectDisposedException">Thrown when the object has been disposed.</exception>
        Friend Function GetValue(valueName As String, defaultValue As Object) As Object Implements IRegistryKey.GetValue
            ObjectDisposedException.ThrowIf(_isDisposed, NameOf(MicrosoftRegistryKey))
            Return _root.GetValue(valueName, defaultValue)
        End Function

        ''' <summary>
        ''' Releases all resources used by the <see cref="MicrosoftRegistryKey"/> instance.
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        ''' <summary>
        ''' Disposes the managed resources of the <see cref="MicrosoftRegistryKey"/> instance.
        ''' </summary>
        ''' <param name="disposing">A boolean value indicating whether to dispose managed resources.</param>
        Private Sub Dispose(disposing As Boolean)
            If Not _isDisposed Then
                If disposing Then
                    _root.Close()
                End If
                _isDisposed = True
            End If
        End Sub
    End Class
End Namespace
