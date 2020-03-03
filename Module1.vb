Imports System.IO
Imports System.Text.RegularExpressions

Module Module1

    Sub Main()
        'call this console App to retrieve the most current file matching a specific ending

        'setup vars
        Dim directorystr
        Dim suspectedfilenameending

        'Link should be follow
        'consult local "DynLinks.xml" File for alternate names and RegExpatterns
        'basically a run CMD with additional Logic
        'Shortly searches XML to fguess a filename
        'if guessed it open the file

        'arg4 = Path to file/Executable
        'arg5 Regex pattern (X for numbers 0-9)
        'check if argument is passed on startup
        ' If Environment.GetCommandLineArgs.Count <> 1 Then
        'testcases

        If Environment.GetCommandLineArgs.Count = 3 Then

            directorystr = Environment.GetCommandLineArgs(1)
            suspectedfilenameending = Environment.GetCommandLineArgs(2)

        Else
        Return
        End If


        Dim Listofallfilesindirectory As New ArrayList
        'get a list of all files in directory and sub dirs
        'then check if any of them matches the given regex, if so, start the file an store new path in Array
        If directorystr.ToString.Contains("&&") Then
            Dim temparray = StringtoArray(directorystr, "&&")
            For Each dirstring In temparray
                Dim tempresultarray As New ArrayList
                tempresultarray = GetFileList(Environment.ExpandEnvironmentVariables(dirstring), True)
                For Each CandidateFile In tempresultarray
                    'add each found file to fiellistarray 
                    Listofallfilesindirectory.Add(CandidateFile)
                Next

            Next
                Else
            Listofallfilesindirectory = GetFileList(Environment.ExpandEnvironmentVariables(directorystr), True)
        End If
        'compare regex


        Dim ii As Integer = 0
        Dim ResultRegexString As String = suspectedfilenameending.ToString & "$"
        'now build the reg ex char


        'when done convert to actual regex =>

        Dim filefound As Boolean = False
        Dim filefoundlist As New ArrayList
        For Each file As String In Listofallfilesindirectory
            Dim regsuccess = Regex.IsMatch(file, ResultRegexString)
            If regsuccess = True Then
                'the new filename is found, store in array and execute

                'store info that file was found
                filefound = True
                filefoundlist.Add(file)
            End If
        Next

        If filefound = False Then
            'no file found, ask user to readd - maybe file was moved or deleted
            Console.WriteLine("Couldn't find a File")
            Return
        Else
            'get file with most current edit date and execute it
            Dim infoReader As System.IO.FileInfo
            Dim filetoexecute
            Dim mostcurrenteditdate
            If filefoundlist.Count > 1 Then
                For Each file In filefoundlist
                    infoReader = My.Computer.FileSystem.GetFileInfo(file)
                    If mostcurrenteditdate = Nothing Then
                        mostcurrenteditdate = infoReader.LastWriteTime
                        filetoexecute = file
                    Else
                        If mostcurrenteditdate < infoReader.LastWriteTime Then
                            mostcurrenteditdate = infoReader.LastWriteTime
                            filetoexecute = file
                        End If
                    End If
                Next
            Else
                'execute found file ,only 1 
                filetoexecute = filefoundlist(0)
            End If
            'execute
            'Process.Start(filetoexecute)
            Console.WriteLine(filetoexecute)
            ' Return

        End If

    End Sub

    Function StringtoArray(seperatedmultistring As String, seperator As String) As ArrayList
        'reset returnarray if it s not null (redim it with 0 members)
        Dim returnarray As New ArrayList
        'set counter
        Dim Charsalreadyscanned = 0
        'make temp var to not change original input var
        Dim tempstring = seperatedmultistring
        Dim foundwords = 0
        Dim lastchar = ""
        For i = 0 To seperatedmultistring.ToString.Length Step 1
            If (seperatedmultistring.ToString.Length = 0) Then
                'empty string was given, reutrn null
                Throw New ArgumentException("Calling this Function without Input (enmpty string) is NOT valid!")
            End If
            If (i = seperatedmultistring.ToString.Length) Then
                Exit For
            End If
            'check if current char is the seperator sign
            If (seperatedmultistring.Substring(i, 1) & lastchar = seperator) Then
                'seperator sign found
                Dim scannedtext = tempstring.Substring(Charsalreadyscanned, i - Charsalreadyscanned)
                Dim temparraymember = scannedtext.Substring(0, scannedtext.Length)
                Charsalreadyscanned = Charsalreadyscanned + scannedtext.Length + 1
                If temparraymember.Length = 0 Then
                    returnarray.Add(temparraymember.Substring(0, temparraymember.Length).ToString)

                ElseIf temparraymember.Length > 0 Then
                    returnarray.Add(temparraymember.Substring(0, temparraymember.Length - 1).ToString)
                End If
                foundwords = foundwords + 1
            End If
            lastchar = seperatedmultistring.Substring(i, 1)
        Next
        Return returnarray
    End Function


    Function GetFileList(ByVal DirPath As String,
 Optional IncludeSubFolders As Boolean = True) As ArrayList

        Dim Filelist As New ArrayList
        Dim objDir As DirectoryInfo = New DirectoryInfo(DirPath)
        Dim objSubFolder As DirectoryInfo
        Try
            'add number of files in current dir
            For Each fileinfoobj1 As FileInfo In objDir.GetFiles
                Filelist.Add(fileinfoobj1.FullName)
            Next
            'Filelist.Add(objDir.GetFiles)
            'call recursively to get sub folders
            'if you don't want this set optional
            'parameter to false 
            If IncludeSubFolders Then
                For Each objSubFolder In objDir.GetDirectories()
                    For Each fileinfoobj As FileInfo In GetFileList(objSubFolder.FullName)
                        Filelist.Add(fileinfoobj.FullName)
                    Next
                Next
            End If

        Catch Ex As Exception

        End Try

        Return Filelist
    End Function
End Module
