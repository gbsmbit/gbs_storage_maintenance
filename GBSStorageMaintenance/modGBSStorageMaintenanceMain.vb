Imports System.IO
Imports System.Linq.Expressions
Imports System.Reflection
Imports System.Runtime.CompilerServices

Module modGBSStorageMaintenanceMain

    Sub Main()

        ParameterFileName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ParameterFileName)
        LogFileName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), LogFileName & "_" & Now.ToString("yyyymmdd_HHmmss") & ".log")

        WriteToLog("Beginning execution")

        ReadParametersFromFile()

        If Not Volumes Is Nothing AndAlso Volumes.Count > 0 Then

            Dim volNo As Integer = 0

            For Each vol As Volume In Volumes
                volNo += 1
                PerformVolumeMaintenance(vol, volNo)
            Next

        End If

        WriteToLog("Done")

    End Sub

    Private Sub ReadParametersFromFile()

        If File.Exists(ParameterFileName) Then

            Dim blk As Block = Nothing
            Dim Blocks As New List(Of Block)

            Volumes = New List(Of Volume)

            Dim flEntries As String() = File.ReadAllLines(ParameterFileName)

            For Each ln As String In flEntries
                ln = ln.Replace(vbTab, " ").Replace("""", " ")
                If ln.Trim.Length >= CommentPrefix.Length AndAlso ln.Trim.Substring(0, CommentPrefix.Length) <> CommentPrefix Then
                    If blk Is Nothing AndAlso ln.Trim.ToUpper.Contains(StartTag.ToUpper) Then
                        blk = New Block
                    ElseIf Not blk Is Nothing AndAlso Not ln.Trim.ToUpper.Contains(EndTag.ToUpper) Then
                        blk.Lines.Add(ln)
                    ElseIf Not blk Is Nothing AndAlso ln.Trim.ToUpper.Contains(EndTag.ToUpper) Then
                        Blocks.Add(blk)
                        blk = Nothing
                    End If
                End If
            Next

            If Not blk Is Nothing Then
                Blocks.Add(blk)
            End If

            For Each blck As Block In Blocks

                Dim pth As String = ""
                Dim incl As Boolean = False
                Dim age As Integer = MaximumAgeDefault
                Dim exc As String = ""
                Dim vbs As Boolean = True
                Dim rmv As Boolean = True

                For Each ln As String In blck.Lines

                    Dim parm As String() = ln.Split({ParameterDelimiter}, StringSplitOptions.RemoveEmptyEntries)

                    If parm.Count > 1 Then

                        Try
                            If parm(0).Trim.ToUpper.Contains("PATH") Then
                                pth = parm(1).Trim
                            ElseIf parm(0).Trim.ToUpper.Contains("INCLUDE") Then
                                If parm(1).Trim.ToUpper = "YES" Then
                                    incl = True
                                End If
                            ElseIf parm(0).Trim.ToUpper.Contains("MAXAGE") Then
                                age = CInt(parm(1).Trim)
                            ElseIf parm(0).Trim.ToUpper.Contains("EXCLUSION") Then
                                exc = parm(1).Trim
                            ElseIf parm(0).Trim.ToUpper.Contains("VERBOSE") Then
                                If parm(1).Trim.ToUpper = "NO" Then
                                    vbs = False
                                End If
                            ElseIf parm(0).Trim.ToUpper.Contains("REMOVE") Then
                                If parm(1).Trim.ToUpper = "NO" Then
                                    rmv = False
                                End If
                            End If
                        Catch ex As Exception
                        End Try

                    End If

                Next

                If pth.Length > 0 Then
                    Volumes.Add(New Volume(pth, incl, age, exc, vbs, rmv))
                End If

            Next

            Dim msg As String = ""
            Dim cnt As Integer = 0
            If Volumes.Count > 0 Then
                msg = "Read parameters for " & Volumes.Count.ToString & " volumes targeted for maintenance:" & vbCrLf
                For Each vol As Volume In Volumes
                    cnt += 1
                    msg += "        Volume " & cnt.ToString & ":" & vbCrLf
                    msg += "                Path = " & vol.Path & ";" & vbCrLf
                    msg += "                IncludeSubFolders = " & vol.IncludeSubFolders.ToString & ";" & vbCrLf
                    msg += "                MaxAge = " & vol.MaximumAge & " days;" & vbCrLf
                    msg += "                Exclusions = " & vol.Exclusions & ";" & vbCrLf
                    msg += "                VerboseLog = " & vol.VerboseLogging.ToString & ";" & vbCrLf
                    msg += "                ActuallyRemove = " & vol.ActuallyDelete.ToString & ";" & vbCrLf
                Next
            Else
                msg = "Apparently there are no volumes targeted for maintenance:"
            End If

            WriteToLog(msg)

        Else
            CreateInitialParameterFileName()
        End If

    End Sub

    Private Sub PerformVolumeMaintenance(ByRef vol As Volume, ByVal volNo As Integer)

        WriteToLog("Beginning maintenance on Volume No. " & volNo.ToString & ": " & vol.Path)

        If Directory.Exists(vol.Path) Then

            Dim StorageReleased As Long = 0
            Dim strStorageReleased As String = ""
            Dim dir As New DirectoryInfo(vol.Path)
            Dim files As FileInfo() = Nothing

            If vol.IncludeSubFolders Then
                files = dir.GetFiles("*.*", SearchOption.AllDirectories)
            Else
                files = dir.GetFiles("*.*", SearchOption.TopDirectoryOnly)
            End If

            For Each fl As FileInfo In files

                Dim LastModified As DateTime = fl.LastWriteTime
                Dim DaysSinceLastModified As Integer = DateDiff(DateInterval.Day, LastModified, Now)
                Dim flSize As Long = fl.Length
                Dim flDetail As String = fl.FullName & vbTab & " (Last Modified: " & LastModified.ToString("yyyy/MM/dd HH:mm:ss") & "; " &
                                                             "Days since Last Modified: " & DaysSinceLastModified.ToString & "; " &
                                                             "Size: " & flSize.ToString & " bytes)"

                If Not vol.Exclusions.ToUpper.Contains(fl.Extension.Replace(".", "").ToUpper) Then
                    If DateDiff(DateInterval.Day, LastModified, Now) > vol.MaximumAge Then
                        If Year(LastModified) > 1980 Then
                            If vol.ActuallyDelete Then
                                Try
                                    fl.Delete()
                                    WriteToLog("Successfully removed file: " & flDetail)
                                Catch ex As Exception
                                    WriteToLog("Failed to removed file: " & flDetail & vbCrLf & "        Error: " & ex.Message)
                                End Try
                            Else
                                WriteToLog("File would be removed if I was allowed to: " & flDetail)
                            End If
                            StorageReleased += flSize
                        Else
                            If fl.Length >= MaximumPathLength Then
                                WriteToLog("File reports weird last modification date (path may be too long), so not risking a removal: " & flDetail)
                            Else
                                WriteToLog("File reports weird last modification date, so not risking a removal: " & flDetail)
                            End If
                        End If
                    Else
                        If vol.VerboseLogging Then
                            WriteToLog("File younger than age threshold: " & flDetail)
                        End If
                    End If
                Else
                    If vol.VerboseLogging Then
                        WriteToLog("File ignored due to exclusions: " & flDetail)
                    End If
                End If

            Next

            WriteToLog("Maintenance complete on Volume No. " & volNo.ToString & ": " & vol.Path)

            If StorageReleased < 1024 Then
                strStorageReleased = StorageReleased.ToString & " bytes"
            ElseIf StorageReleased < (1024 * 1024) Then
                strStorageReleased = Math.Round((StorageReleased / 1024), 3).ToString & " Kb"
            ElseIf StorageReleased < (1024 * 1024 * 1024) Then
                strStorageReleased = Math.Round((StorageReleased / (1024 * 1024)), 3).ToString & " Mb"
            Else
                strStorageReleased = Math.Round((StorageReleased / (1024 * 1024 * 1024)), 3).ToString & " Gb"
            End If

            If vol.ActuallyDelete Then
                WriteToLog("Total storage released on Volume No. " & volNo.ToString & ": " & strStorageReleased & vbCrLf)
            Else
                WriteToLog("Total storage I could have released on Volume No. " & volNo.ToString & ": " & strStorageReleased & vbCrLf)
            End If

        Else
            WriteToLog("Path specified does not exist, or is not accessible: " & vol.Path & vbCrLf)
        End If

    End Sub

    Private Sub WriteToLog(ByVal msg As String)

        Using lg As New StreamWriter(LogFileName, append:=True)
            lg.WriteLine(Now.ToString("yyyy/MM/dd HH:mm:ss") & " - " & msg)
        End Using

    End Sub

    Private Sub CreateInitialParameterFileName()

        Using fl As New StreamWriter(ParameterFileName)

            fl.WriteLine(CommentPrefix & " GBSStorageMaintenance parameters" & vbCrLf)

            fl.WriteLine(CommentPrefix & " Sample storage instance block - add as many of these as you like - try to get the syntax right, otherwise nothing is guaranteed")
            fl.WriteLine(CommentPrefix & " Everything is cASE iNSENSITive and spacing between   words  or  lines          is        irrelevant")
            fl.WriteLine(CommentPrefix & " The equals sign (=) is the delimiter, and is thus somewhat important")
            fl.WriteLine(CommentPrefix & " Comment lines, prefixed with !-- are ignored" & vbCrLf)
            fl.WriteLine(CommentPrefix & " " & StartTag)
            fl.WriteLine(CommentPrefix & "     Path = D:" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "(Compulsory)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "- UNC paths accepted")
            fl.WriteLine(CommentPrefix & "     IncludeSubFolders = [Yes|No]" & vbTab & "(Optional - defaults to No if omitted)" & vbTab & vbTab & vbTab & "- To include or not include files contained in subfolders of Path")
            fl.WriteLine(CommentPrefix & "     MaxAge = 90" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "(Optional - defaults to 90 days if omitted)" & vbTab & vbTab & "- Files older than this many days will be removed")
            fl.WriteLine(CommentPrefix & "     Exclusions = pdf|xlsx|..." & vbTab & vbTab & "(Optional - defaults to None if omitted)" & vbTab & vbTab & "- Files with these extensions will be ignored")
            fl.WriteLine(CommentPrefix & "     VerboseLog = [Yes|No]" & vbTab & vbTab & vbTab & "(Optional - defaults to Yes if omitted)" & vbTab & vbTab & vbTab & "- Log either all actions, or file removals only")
            fl.WriteLine(CommentPrefix & "     ActuallyRemove = [Yes|No]" & vbTab & vbTab & "(Optional - defaults to Yes if omitted)" & vbTab & vbTab & vbTab & "- allows dry run")
            fl.WriteLine(CommentPrefix & " " & EndTag & vbCrLf)

        End Using

        WriteToLog("No parameter file found - new sample framework file created")

    End Sub

End Module
