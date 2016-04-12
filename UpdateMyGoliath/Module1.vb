Option Explicit On
Imports System.Windows

Module Module1
    '---Author:  Sherry Measom
    '---Dept:    DR
    '---Created: 02/08/2016
    '---Updated: 03/04/2016


    Dim cVersion As String = "2.0.1.2U"



    Dim ObjRegistry
    Dim Key
    Dim WShell As New Object
    Dim cFullUserName
    Dim cUserName
    Dim nButton
    Dim ObjFileSystem
    Dim cFileName
    Dim ObjTxtFile
    Dim pb As New ProgressBar
    Dim PercentComplete
    Dim lContinue
    Dim nCount
    Dim nStep
    Dim cSourceFileName
    Dim cDestinationPath
    Dim cSourcePath
    Dim lUpdatedEnvironment = False
    Dim lGoliathUpdateComplete = False
    Dim nMaxSteps = 5
    Dim lCheckGoliath = True
    Dim lCheckIRap = True
    Dim lCheckOpenRequest = True
    Dim cCopyType
    Dim MsgReply


    Function UpdateMe()

        UpdateMe = False
        '---Prompt User

        MsgReply = MsgBox("This VBScript will update Goliath minor versions." & Chr(13) &
            "Rebooting of your machine Is Not REQUIRED." & Chr(13) & Chr(13) &
            "Do you want to continue?  (defaults to No after 30 seconds)", vbYesNo, "Data Resources Application Update")

        If MsgReply = 7 Then
            MsgBox("Canceled by User!", vbOKOnly, "Data Resources Application Update")
            Exit Function
        End If

        '---Create a Progress Bar Window
        pb = New ProgressBar

        '---Prep the Progress Bar
        PercentComplete = 0
        pb.SetTitle("Step 1 of " & nMaxSteps)
        pb.SetText("Obtain User Information from LDAP")
        pb.Show()

        '---Pause briefly  
        Threading.Thread.Sleep(5000)

        '---Get the Stored Username and Full Username
        cUserName = WShell.ExpandEnvironmentStrings("%UserName%")
        cFullUserName = WShell.ExpandEnvironmentStrings("%FULLUSERNAME%")
        If cFullUserName = "%FULLUSERNAME%" Then
            '---Get the Full User Name from LDAP
            cFullUserName = GetFullNameFromLDAP()
        End If

        '---Updated Progress Bar Percentage
        pb.Update(10)
        pb.BarRefresh()


        '---Updating Progress Bar
        If lCheckGoliath Then
            pb.SetText("Checking if 'Goliath' is running...")
            pb.BarRefresh()

            If IsRunning("Goliath.exe") Then
                '---Updating Progress Bar
                pb.SetText("Terminating 'Goliath'...")
                pb.BarRefresh()

                '---Terminate Goliath Application
                KillProcess("Goliath.exe")
            End If
        End If

        '---Updated Progress Bar Percentage
        pb.Update(16)

        '---Updating Progress Bar
        If lCheckIRap Then
            pb.SetText("Checking if 'I-Rap' is running...")
            pb.BarRefresh()

            If IsRunning("i-Rap.exe") Then
                '---Updating Progress Bar
                pb.SetText("Terminating 'i-Rap'...")
                pb.BarRefresh()

                '---Terminate i-Rap Application
                KillProcess("i-Rap.exe")
            End If
        End If

        '---Updated Progress Bar Percentage
        pb.Update(18)

        '---Updating Progress Bar
        If lCheckOpenRequest Then
            pb.SetText("Checking if 'CAS Request Launcher' is running...")
            pb.BarRefresh()

            If IsRunning("OpenRequest.exe") Then
                '---Updating Progress Bar
                pb.SetText("Terminating 'CAS Request Launcher'...")
                pb.BarRefresh()

                '---Terminate CAS Request Launcher Application
                KillProcess("OpenRequest.exe")
            End If
        End If

        '---Updated Progress Bar Percentage
        pb.Update(20)



        '---Initalize the Continuation Flag
        lContinue = True

        '---Build the Data Resources Applications Folder
        '---Updating Progress Bar
        pb.SetText("Validating Destination Location...")
        pb.BarRefresh()

        'If Not FolderBuild("C:\", "C:\Data Resources Apps") Then
        '  MsgBox "* * * ERROR * * *" + Chr(13) + "Unable to create folder..."
        '  lContinue = False
        'End If

        '  CreateFolderRecursive("C:\Data Resources Apps")

        If (Not System.IO.Directory.Exists("C:\Data Resources Apps")) Then
            MsgBox("You do not currently have Goliath installed at C:\Data Resources Apps. You must have Goliath installed to update it.", vbOKOnly, "Directory Does not exist")
            Exit Function
        End If

        '---Pause briefly  
        Threading.Thread.Sleep(5000)

        '---Create File System Object
        ObjFileSystem = CreateObject("Scripting.FileSystemObject")

        '---Initalize Variables
        nStep = 0
        nCount = 0

        '---Check to see if the process should continue
        If lContinue Then

            '---Check and Copy Files for Goliath
            For nCount = 1 To 3
                Select Case nCount
                    Case 1
                        cSourceFileName = "Goliath.EXE"
                        cSourcePath = "\\msp09fil01\d906\casean\goliath"
                        cDestinationPath = "C:\Data Resources Apps"
                        cCopyType = "I"
        'cVersion = ??? - want to add this to the log

                    Case 2
                        cSourceFileName = "Goliath.exe.config"
                        cSourcePath = "\\msp09fil01\d906\casean\goliath"
                        cDestinationPath = "C:\Data Resources Apps"
                        cCopyType = "I"

                    Case 3
                        cSourceFileName = "Goliath.exe.manifest"
                        cSourcePath = "\\msp09fil01\d906\casean\goliath"
                        cDestinationPath = "C:\Data Resources Apps"
                        cCopyType = "I"

                End Select

                '---Update Progress Bar
                nStep = nCount + 2
                pb.SetTitle("Step " & nStep & " of " & nMaxSteps)
                pb.SetText("Authenticating '" & cSourceFileName & "'")
                pb.BarRefresh()

                '---Copy File if 'cCopyType = I' else delete it.
                If cCopyType = "I" Then
                    '---Check to see if the file can be copied.
                    If OkayToCopy(cSourcePath & "\" & cSourceFileName, cDestinationPath & "\" & cSourceFileName) Then
                        '---Check to see if the destination folder exists and copy the file.
                        If ObjFileSystem.FolderExists(cDestinationPath) Then
                            '---Update Progress Bar
                            pb.SetText("Installing '" & cSourceFileName & "'")
                            pb.BarRefresh()

                            '---Copying File
                            ObjFileSystem.CopyFile(cSourcePath & "\" & cSourceFileName, cDestinationPath & "\" & cSourceFileName, True)
                        End If
                    End If
                Else
                    '---Check to see if the file can be deleted
                    If OkayToDelete(cDestinationPath & "\", cSourceFileName) Then
                        '---Update Progress Bar
                        pb.SetText("Removing '" & cSourceFileName & "'")
                        pb.BarRefresh()

                        '---Removing File
                        ObjFileSystem.DeleteFile(cDestinationPath & "\" & cSourceFileName, True)
                    End If
                End If

                '---Updated Progress Bar
                pb.Update(nMaxSteps * nCount)
            Next
        End If

        If UpdateMe = True Then
            '---Write to the text file that the User ran the Script
            ObjFileSystem = CreateObject("Scripting.FileSystemObject")
            cFileName = "\\msp09fil01\d906\casean\Goliath\Logs\InstallScriptLogTest.txt"
            ObjTxtFile = ObjFileSystem.OpenTextFile(cFileName, 8, True)
            '----working on adding file version being loaded
            ObjTxtFile.WriteLine(cFullUserName & " (" & WShell.ExpandEnvironmentStrings("%UserName%") & ") - " & Now() & " on machine " & WShell.ExpandEnvironmentStrings("%COMPUTERNAME%") & " V" & cVersion)
            '---ObjTxtFile.WriteLine(cFullUserName & " (" & wShell.ExpandEnvironmentStrings("%UserName%") & ") - " & "Version " & cVersion & "on " & Now() & " on machine " & wShell.ExpandEnvironmentStrings("%COMPUTERNAME%"))
            ObjTxtFile.Close
            ObjTxtFile = Nothing
            ObjFileSystem = Nothing
        Else
            MsgBox("Update Failed.")
        End If

        '---Close Progress Bar
        pb.Close()

        '---Exit Script
        WShell = Nothing
        MsgBox("Your system has been updated!", vbInformation, "Process Complete")

    End Function
    '---LDAP Function for getting full user name
    Function GetFullNameFromLDAP()
        Dim ObjSysInfo
        Dim ObjUser
        'Dim WShell
        Dim StrUser
        Dim StrTempName
        Dim StrFullName
        Dim nPos

        '---Create System Info Object
        ObjSysInfo = CreateObject("ADSystemInfo")

        '---Create LDAP User Object
        ObjUser = GetObject("LDAP://" & ObjSysInfo.UserName & "")

        '---Initialize Varaibles
        nPos = 0

        '---Get the MS Domain User ID
        StrUser = Replace(ObjUser.Name, "CN=", "")

        '---Get the Stored LDAP Full User Name
        StrTempName = ObjUser.FullName

        '---Find the position of the Comma in the Full User Name from LDAP
        nPos = InStr(1, StrTempName, ",")

        '---If there is a Comma in the name then it is LN, FN MI format, reformat to FN MI LN Format
        If nPos > 0 Then
            '---Reformatting to FN MI LN format
            StrFullName = Trim(Mid(StrTempName, nPos + 1) & " " & Mid(StrTempName, 1, nPos - 1))
        Else
            '---Accepting Format as it is.
            StrFullName = Trim(StrTempName)
        End If

        '---Message Boxes for Testing
        'MsgBox "User: " & StrUser
        'Msgbox "FullName: " & StrFullName

        '---Clearing Objects
        ObjUser = Nothing
        ObjSysInfo = Nothing

        '---Setting Return Value
        GetFullNameFromLDAP = StrFullName
    End Function
    Function OkayToCopy(cSource, cDestination)
        Dim ObjFileSystem
        Dim demofile, date1, date2

        '---Initialize Return Variable
        OkayToCopy = False

        '---Create The File System Object
        ObjFileSystem = CreateObject("Scripting.FileSystemObject")

        '---Check to see if the Destination File Already Exists, if Not, return true that it is okay to copy file.
        If Not ObjFileSystem.FileExists(cDestination) Then
            OkayToCopy = True
            ObjFileSystem = Nothing
            Exit Function
        End If

        '---If Destination File Already Exists, then check to see if the source file is newer, if so, Return true that is it okay to copy file.
        demofile = ObjFileSystem.GetFile(cSource)
        date1 = demofile.DateLastModified
        demofile = ObjFileSystem.GetFile(cDestination)
        date2 = demofile.DateLastModified

        If DateDiff("d", date2, date1) > 0 Then
            OkayToCopy = True
            If Not ObjFileSystem Is Nothing Then ObjFileSystem = Nothing
            Exit Function
        End If
    End Function


    Function OkayToDelete(cDestination, cFileName)
        Dim ObjFileSystem

        '---Initialize Return Variable
        OkayToDelete = False

        '---Create The File System Object
        ObjFileSystem = CreateObject("Scripting.FileSystemObject")

        '---Check to see if the Destination File Already Exists, if so, return true that it is okay to delete file.
        If ObjFileSystem.FileExists(cDestination & cFileName) Then
            If IsRunning(cFileName) Then
                '---Terminate CAS Request Launcher Application
                KillProcess(cFileName)
            End If
            '---Set Return Variable based on if the file is still running
            OkayToDelete = Not IsRunning(cFileName)

            ObjFileSystem = Nothing
        End If
    End Function


    Function IsRunning(cProcessName)
        Dim colProcessList, ObjProcess

        '---Initialize Variables
        IsRunning = False

        '---Get Process list Object
        colProcessList = GetObject("Winmgmts:").ExecQuery("Select * from Win32_Process")

        '---Check all the processes
        For Each ObjProcess In colProcessList
            If ObjProcess.name = cProcessName Then
                IsRunning = True
            End If
        Next

    End Function


    Function KillProcess(cProcessName)
        Dim WShellProcess

        '---Create WScript Object
        WShellProcess = CreateObject("WScript.Shell")

        '---Check to see if it is running still.
        If IsRunning(cProcessName) Then WShellProcess.Run("taskkill /im " & cProcessName, , True)

        '---Exit Script
        WShellProcess = Nothing

    End Function


    Function GetVersion(cSourcePath, cSourceFileName)
        GetVersion = "0.0.0.0"
        cSourceFileName = "Goliath.EXE"
        cSourcePath = "\\msp09fil01\d906\casean\goliath"

        Dim objShell, objFolder, objFolderItem, i
        If FileIO.FileSystem.FileExists(cSourcePath & "\" & cSourceFileName) Then
            objShell = CreateObject("Shell.Application")
            objFolder = objShell.Namespace(cSourcePath)
            objFolderItem = objFolder.ParseName(cSourceFileName)
            Dim arrHeaders(300)
            For i = 0 To 300
                arrHeaders(i) = objFolder.GetDetailsOf(objFolder.Items, i)
                'WScript.Echo i &"- " & arrHeaders(i) & ": " & objFolder.GetDetailsOf(objFolderItem, i)
                If LCase(arrHeaders(i)) = "product version" Then
                    GetVersion = objFolder.GetDetailsOf(objFolderItem, i)
                    Exit For
                End If
            Next
        End If
    End Function

End Module
