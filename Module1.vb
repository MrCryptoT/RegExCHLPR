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
        Listofallfilesindirectory = GetFileList(Environment.ExpandEnvironmentVariables(directorystr), True)

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
