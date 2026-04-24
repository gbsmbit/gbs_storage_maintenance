Imports System.IO
Imports System.Linq.Expressions
Imports System.Reflection
Imports System.Runtime.CompilerServices

Module modGBSStorageMaintenanceMain

    Sub Main()

        ParameterFileName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ParameterFileName)

        If Not Directory.Exists(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Logs")) Then
            Try
                Directory.CreateDirectory(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Logs"))
                LogFileName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Logs", LogFileName & "_" & Now.ToString("yyyymmdd_HHmmss") & ".log")
            Catch ex As Exception
                LogFileName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), LogFileName & "_" & Now.ToString("yyyymmdd_HHmmss") & ".log")
            End Try
        Else
            LogFileName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Logs", LogFileName & "_" & Now.ToString("yyyymmdd_HHmmss") & ".log")
        End If

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
                    If vol.Exclusions.Length = 0 Then
                        msg += "                Exclusions = None;" & vbCrLf
                    Else
                        msg += "                Exclusions = " & vol.Exclusions & ";" & vbCrLf
                    End If
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
                                                                "Size: " & FileSizeReadable(flSize) & ")"

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

            If vol.ActuallyDelete Then
                WriteToLog("Total storage released on Volume No. " & volNo.ToString & ": " & FileSizeReadable(StorageReleased) & vbCrLf)
            Else
                WriteToLog("Total storage I could have released on Volume No. " & volNo.ToString & ": " & FileSizeReadable(StorageReleased) & vbCrLf)
            End If

        Else
            WriteToLog("Path specified does not exist, or is not accessible: " & vol.Path & vbCrLf)
        End If

    End Sub

End Module
