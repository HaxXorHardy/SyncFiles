Imports System.IO
Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim dir1 As String = "C:\Users\*\Dropbox\DropsyncFiles"
        Dim dir2 As String = "C:\Users\*\Documents\my games\Fallout Shelter"
        Dim sync As New myDirMonitor(dir1, dir2)

    End Sub

End Class

Public Class myDirMonitor

    Dim Dir1 As DirectoryInfo
    Dim Dir2 As DirectoryInfo

    Public Enum SyncResults
        Successfull = 0
        NotSuccessfull = 1
    End Enum

    Public Sub New(ByVal dir1 As String, ByVal dir2 As String)
        Me.Dir1 = New DirectoryInfo(dir1)
        Me.Dir2 = New DirectoryInfo(dir2)
        'Both directories must exist... if not an exception is thrown
        If (Not Me.Dir1.Exists) OrElse (Not Me.Dir2.Exists) Then Throw New DirectoryNotFoundException
        Me.BeginSynchronization()
    End Sub

    Public Function BeginSynchronization() As SyncResults
        Rtb("Start sync!", 3, True)
        If Me.SyncProcess(Dir1, Dir2) = SyncResults.NotSuccessfull Then Return SyncResults.NotSuccessfull
        Rtb("Sync1 complete! Change direction -><-", 3, True)
        If Me.SyncProcess(Dir2, Dir1) = SyncResults.NotSuccessfull Then Return SyncResults.NotSuccessfull
        Rtb("Sync2 complete! all done", 3, True)
        Return SyncResults.Successfull
    End Function


    Private Function SyncProcess(
        ByVal sourceDir As DirectoryInfo,
        ByVal destinationDir As DirectoryInfo) As SyncResults
        Try
            Me.SyncDir(sourceDir, destinationDir)
        Catch ex As Exception
            Return SyncResults.NotSuccessfull
        End Try
        Return SyncResults.Successfull
    End Function

    Private Sub SyncDir(
        ByVal sourceDir As DirectoryInfo,
        ByVal destinationDir As DirectoryInfo)
        Dim files As FileInfo() = sourceDir.GetFiles
        Dim directories As DirectoryInfo() = sourceDir.GetDirectories
        'Sync files...
        For f As Integer = 0 To files.Length - 1
            If File.Exists(destinationDir.FullName & "\" & files(f).Name) Then
                Rtb("File: ", 0, False)
                Rtb(files(f).Name, 2, False)
                Rtb(" exists!", 1, False)
                If files(f).LastWriteTime > File.GetLastAccessTime(destinationDir.FullName & "\" & files(f).Name) Then
                    File.Copy(files(f).FullName, destinationDir.FullName & "\" & files(f).Name, True)
                    Rtb(" but is Newer ---> copy!", 3, True)
                    Rtb(files(f).LastWriteTime & " -- " & File.GetLastAccessTime(destinationDir.FullName & "\" & files(f).Name), 0, True)
                Else
                    Rtb(" and is up to date!", 1, True)
                End If
            Else
                File.Copy(files(f).FullName, destinationDir.FullName & "\" & files(f).Name)
                Rtb("File: ", 0, False)
                Rtb(files(f).Name, 2, False)
                Rtb(" not exist ---> copy!", 3, True)
            End If
        Next
        'Sync directories
        For d As Integer = 0 To directories.Length - 1
            If Not Directory.Exists(destinationDir.FullName & "\" & directories(d).Name) Then
                Rtb("Directory: ", 0, False)
                Rtb(directories(d).Name, 2, False)
                Rtb(" not exist ---> create!", 3, True)
                Directory.CreateDirectory(destinationDir.FullName & "\" & directories(d).Name)
            End If

            Me.SyncDir(directories(d), New DirectoryInfo(destinationDir.FullName & "\" & directories(d).Name))
        Next
    End Sub

    Private Sub Rtb(ByVal str As String, ByVal colour As Integer, ByVal newLine As Boolean)
        With Form1.RichTextBox1
            If colour = 1 Then
                .SelectionColor = Color.YellowGreen
            ElseIf colour = 2 Then
                .SelectionColor = Color.LimeGreen
            ElseIf colour = 3 Then
                .SelectionColor = Color.Red
            Else
                .SelectionColor = Color.Black
            End If
            .SelectionFont = New Font(Form1.RichTextBox1.Font, FontStyle.Bold)
            If newLine Then
                .AppendText(str & vbNewLine)
            Else
                .AppendText(str)
            End If
        End With
    End Sub

End Class