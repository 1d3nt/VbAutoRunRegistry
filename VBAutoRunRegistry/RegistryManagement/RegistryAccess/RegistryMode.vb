Namespace RegistryManagement.RegistryAccess

    ''' <summary>
    ''' Enumerates the modes of access for interacting with registry keys.
    ''' This enumeration specifies the level of access granted when opening a registry key.
    ''' </summary>
    Public Enum RegistryMode

        ''' <summary>
        ''' Indicates read-only access to a registry key.
        ''' Users can retrieve values from the key but cannot modify or delete any data.
        ''' </summary>
        Read

        ''' <summary>
        ''' Indicates read and write access to a registry key.
        ''' Users can retrieve values from the key and also modify, add, or delete data.
        ''' </summary>
        ReadWrite
    End Enum
End Namespace
