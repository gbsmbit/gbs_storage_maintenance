Imports System.IO

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

    Public Sub WriteToLog(ByVal msg As String)

        Using lg As New StreamWriter(LogFileName, append:=True)
            lg.WriteLine(Now.ToString("yyyy/MM/dd HH:mm:ss") & " - " & msg)
        End Using

    End Sub

    Public Sub CreateInitialParameterFileName()

        Using fl As New StreamWriter(ParameterFileName)

            fl.WriteLine(CommentPrefix & " GBSStorageMaintenance parameters" & vbCrLf)

            fl.WriteLine(CommentPrefix & " Sample storage instance block - add as many of these as you like - try to get the syntax right, otherwise nothing is guaranteed")
            fl.WriteLine(CommentPrefix & " Everything is cASE iNSENSITive and spacing between   words  or  lines          is        irrelevant")
            fl.WriteLine(CommentPrefix & " The equals sign (=) is the delimiter, and is thus somewhat important")
            fl.WriteLine(CommentPrefix & " Comment lines, prefixed with !-- are ignored" & vbCrLf)
            fl.WriteLine(CommentPrefix & " " & StartTag & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "(Compulsory - let's not omit)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "- Denotes the beginning of a volume definition block")
            fl.WriteLine(CommentPrefix & "     Path = D:" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "(Compulsory)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "- UNC paths accepted")
            fl.WriteLine(CommentPrefix & "     IncludeSubFolders = [Yes|No]" & vbTab & "(Optional - defaults to No if omitted)" & vbTab & vbTab & vbTab & vbTab & "- To include or not include files contained in subfolders of Path")
            fl.WriteLine(CommentPrefix & "     MaxAge = 90" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "(Optional - defaults to 90 days if omitted)" & vbTab & vbTab & vbTab & "- Files older than this many days will be removed")
            fl.WriteLine(CommentPrefix & "     Exclusions = pdf|xlsx|..." & vbTab & vbTab & "(Optional - defaults to None if omitted)" & vbTab & vbTab & vbTab & "- Files with these extensions will be ignored")
            fl.WriteLine(CommentPrefix & "     VerboseLog = [Yes|No]" & vbTab & vbTab & vbTab & "(Optional - defaults to Yes if omitted)" & vbTab & vbTab & vbTab & vbTab & "- Log either all actions, or file removals only")
            fl.WriteLine(CommentPrefix & "     ActuallyRemove = [Yes|No]" & vbTab & vbTab & "(Optional - defaults to Yes if omitted)" & vbTab & vbTab & vbTab & vbTab & "- allows dry run")
            fl.WriteLine(CommentPrefix & " " & EndTag & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "(Compulsory - reads only last volume if omitted)" & vbTab & "- Denotes the end of a volume definition block" & vbCrLf)

        End Using

        WriteToLog("No parameter file found - new sample framework file created")

    End Sub

    Public Function FileSizeReadable(ByVal flSize As Long) As String

        Dim FileSizeUnits() As String = {"bytes", "KB", "MB", "GB", "TB", "PB"}
        Dim i As Integer = 0
        Dim dblSize As Double = CDbl(flSize)

        While dblSize >= 1024 AndAlso i < FileSizeUnits.Length - 1
            dblSize /= 1024
            i += 1
        End While

        Return Math.Round(dblSize, 3).ToString() & " " & FileSizeUnits(i)

    End Function

End Module
