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
                    Volumes.Add(New Volume(pth, age, exc, vbs, rmv))
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

        WriteToLog("Beginning maintenance on Volume No. " & volNo.ToString)

        If Directory.Exists(vol.Path) Then

            Dim files = Directory.GetFiles(vol.Path, "*.*", SearchOption.AllDirectories)

            For Each fl In files

                Dim ext As String = Path.GetExtension(fl).Replace(".", "")
                Dim dt As DateTime = File.GetLastWriteTime(fl)
                Dim age As Integer = DateDiff(DateInterval.Day, dt, Now)

                If Not vol.Exclusions.ToUpper.Contains(Path.GetExtension(fl).Replace(".", "").ToUpper) Then
                    If DateDiff(DateInterval.Day, File.GetLastWriteTime(fl), Now) > vol.MaximumAge Then
                        If vol.ActuallyDelete Then
                            Try
                                File.Delete(fl)
                                WriteToLog("Successfully removed file: " & fl)
                            Catch ex As Exception
                                WriteToLog("Failed to removed file: " & fl & vbCrLf & "        Error: " & ex.Message)
                            End Try
                        Else
                            WriteToLog("File would be removed if I was allowed to: " & fl)
                        End If
                    Else
                        If vol.VerboseLogging Then
                            WriteToLog("File younger than age threshold: " & fl)
                        End If
                    End If
                Else
                    If vol.VerboseLogging Then
                        WriteToLog("File ignored due to exclusions: " & fl)
                    End If
                End If

            Next

            WriteToLog("Maintenance complete on Volume No. " & volNo.ToString & vbCrLf)

        Else
            WriteToLog("Path specified does not exist, or is not accessible: " & vol.Path & vbCrLf)
        End If

    End Sub

    Private Sub WriteToLog(ByVal msg As String)

        Using lg As New StreamWriter(LogFileName, append:=True)
            lg.WriteLine(Now.ToString("yyyy/mm/dd HH:mm:ss") & " - " & msg)
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
            fl.WriteLine(CommentPrefix & "     Path = D: (UNC paths accepted)")
            fl.WriteLine(CommentPrefix & "     MaxAge = 90 (Files older than this many days will be removed)")
            fl.WriteLine(CommentPrefix & "     Exclusions = pdf|xlsx (Files with these extensions will be ignored)")
            fl.WriteLine(CommentPrefix & "     VerboseLog = [Yes|No] (Log either all actions, or file removals only)")
            fl.WriteLine(CommentPrefix & "     ActuallyRemove = [Yes|No] (Optional (defaults to Yes) - allows dry run)")
            fl.WriteLine(CommentPrefix & " " & EndTag & vbCrLf)

        End Using

        WriteToLog("No parameter file found - new sample framework file created")

    End Sub

End Module
