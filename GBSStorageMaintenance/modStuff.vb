Module modStuff

    Public Const MaximumPathLength As Integer = 260
    Public Const MaximumAgeDefault As Integer = 90
    Public Const StartTag As String = "<volume>"
    Public Const EndTag As String = "</volume>"

    Public ParameterFileName As String = "GBSStorageMaintenanceParameterFile.txt"
    Public LogFileName As String = "GBSStorageMaintenanceLog"
    Public CommentPrefix As String = "!--"
    Public ParameterDelimiter As Char = "="c

    Public Volumes As List(Of Volume)

End Module
