Imports System.Runtime.InteropServices

Public Class Volume

    Public Property Path As String
    Public Property MaximumAge As Integer
    Public Property Exclusions As String
    Public Property VerboseLogging As Boolean
    Public Property ActuallyDelete As Boolean

    Public Sub New(ByVal _Path As String, ByVal _MaximumAge As Integer, ByVal _Exclusions As String, ByVal _VerboseLogging As Boolean, ByVal _ActuallyDelete As Boolean)

        Path = _Path
        MaximumAge = _MaximumAge
        Exclusions = _Exclusions
        VerboseLogging = _VerboseLogging
        ActuallyDelete = _ActuallyDelete

    End Sub

End Class

Public Class Block

    Public Property Lines As List(Of String)

    Public Sub New()

        Lines = New List(Of String)

    End Sub

End Class
