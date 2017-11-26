Imports System.IO
Imports System.Security.Cryptography

Public Class Form1

#Region " ---> Buttons"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim dir1 As String = "F:\Serie"
        Dim dir2 As String = "J:\Serie"
        Dim checkMD5 As Boolean = False
        StartSync(dir1, dir2, checkMD5)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim dir1 As String = "F:\Games"
        Dim dir2 As String = "J:\Games"
        Dim checkMD5 As Boolean = False
        StartSync(dir1, dir2, checkMD5)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim dir1 As String = "F:\Filme"
        Dim dir2 As String = "J:\Filme"
        Dim checkMD5 As Boolean = False
        StartSync(dir1, dir2, checkMD5)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim dir1 As String = "F:\Programme"
        Dim dir2 As String = "J:\Programme"
        Dim checkMD5 As Boolean = False
        StartSync(dir1, dir2, checkMD5)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Button1_Click(sender, e)
        Button2_Click(sender, e)
        Button3_Click(sender, e)
        Button5_Click(sender, e)
        Button6_Click(sender, e)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim dir1 As String = "C:\Program Files (x86)\Destiny 2"
        Dim dir2 As String = "J:\Destiny 2"
        Dim checkMD5 As Boolean = False
        StartSync(dir1, dir2, checkMD5)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim dir1 As String = "E:\Program Files (x86)\Steam\steamapps\common\PUBG"
        Dim dir2 As String = "J:\PUBG"
        Dim checkMD5 As Boolean = False
        StartSync(dir1, dir2, checkMD5)
    End Sub

#End Region

    'some personal improvements
    ' Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
    '  Dim dir1 As String = "C:\Users\Maurice\Dropbox\DropsyncFiles"
    '       Dim dir2 As String = "C:\Users\Maurice\Documents\my games\Fallout Shelter"
    '        Dim checkMD5 As Boolean = True
    'StartSync(dir1, dir2, checkMD5)
    'End Sub

    Private Sub StartSync(ByVal dir1 As String, ByVal dir2 As String, ByVal checkMD5 As Boolean)
        RichTextBox1.Text = Nothing
        Dim sync As New myDirMonitor(dir1, dir2, checkMD5)
    End Sub


End Class

Public Class myDirMonitor

    Dim Dir1 As DirectoryInfo
    Dim Dir2 As DirectoryInfo
    Dim checkMD5 As Boolean = False

    Public Enum SyncResults
        Successfull = 0
        NotSuccessfull = 1
    End Enum

    Public Sub New(ByVal dir1 As String, ByVal dir2 As String, ByVal checkmd5 As Boolean)
        Me.Dir1 = New DirectoryInfo(dir1)
        Me.Dir2 = New DirectoryInfo(dir2)
        Me.checkMD5 = checkmd5
        'Both directories must exist... if not an exception is thrown
        If (Not Me.Dir1.Exists) OrElse (Not Me.Dir2.Exists) Then Throw New DirectoryNotFoundException
        Me.BeginSynchronization()
    End Sub

    Public Function BeginSynchronization() As SyncResults
        Rtb("--------------------------------------------------------", 3, True)
        Rtb("Start Sync!", 3, True)
        Rtb(Dir1.ToString & " ---> " & Dir2.ToString, 0, True)
        Rtb("--------------------------------------------------------", 3, True)
        Rtb(vbNewLine, 0, True)
        If Me.SyncProcess(Dir1, Dir2) = SyncResults.NotSuccessfull Then Return SyncResults.NotSuccessfull
        Rtb(vbNewLine, 0, True)
        Rtb("--------------------------------------------------------", 3, True)
        Rtb("-> Sync1 complete! Change direction <-", 3, True)
        Rtb(Dir2.ToString & " ---> " & Dir1.ToString, 0, True)
        Rtb("--------------------------------------------------------", 3, True)
        Rtb(vbNewLine, 0, True)
        '
        'Uncomment for direction change!!!
        '
        'If Me.SyncProcess(Dir2, Dir1) = SyncResults.NotSuccessfull Then Return SyncResults.NotSuccessfull
        'Rtb(vbNewLine, 0, True)
        'Rtb("--------------------------------------------------------", 3, True)
        'Rtb("Sync2 complete! all done", 3, True)
        'Rtb("--------------------------------------------------------", 3, True)
        'Rtb(vbNewLine, 0, True)
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
                If files(f).LastWriteTime > File.GetLastWriteTime(destinationDir.FullName & "\" & files(f).Name) Then
                    Rtb(" but is Newer", 3, False)
                    If (MD5FileHash(files(f).FullName) <> MD5FileHash(destinationDir.FullName & "\" & files(f).Name)) Or checkMD5 = False Then
                        Rtb(" ---> copy!", 3, True)
                        'Rtb(MD5FileHash(files(f).FullName) & "  /=  " & MD5FileHash(destinationDir.FullName & "\" & files(f).Name), 0, True)
                        Rtb(files(f).LastWriteTime & " -- " & File.GetLastWriteTime(destinationDir.FullName & "\" & files(f).Name), 0, True)
                        File.Copy(files(f).FullName, destinationDir.FullName & "\" & files(f).Name, True)
                    Else
                        Rtb(" same Hash no need to copy!", 4, True)
                        Rtb(files(f).LastWriteTime & " -- " & File.GetLastWriteTime(destinationDir.FullName & "\" & files(f).Name), 0, True)
                        Rtb(MD5FileHash(files(f).FullName) & "  =  " & MD5FileHash(destinationDir.FullName & "\" & files(f).Name), 0, True)
                    End If
                ElseIf files(f).LastWriteTime < File.GetLastWriteTime(destinationDir.FullName & "\" & files(f).Name) Then
                        Rtb(" and is newer!", 4, True)
                        'Rtb(files(f).LastWriteTime & " -- " & File.GetLastWriteTime(destinationDir.FullName & "\" & files(f).Name), 0, True)
                        'Rtb(MD5FileHash(files(f).FullName) & "  -  " & MD5FileHash(destinationDir.FullName & "\" & files(f).Name), 0, True)
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
                .SelectionColor = Color.Green
            ElseIf colour = 3 Then
                .SelectionColor = Color.Red
            ElseIf colour = 4 Then
                .SelectionColor = Color.Blue
            Else
                .SelectionColor = Color.Black
            End If
            .SelectionFont = New Font(Form1.RichTextBox1.Font, FontStyle.Bold)
            If newLine Then
                .AppendText(str & vbNewLine)
            Else
                .AppendText(str)
            End If
            Form1.RichTextBox1.ScrollToCaret()
        End With
    End Sub

    Private Function MD5FileHash(ByVal sFile As String) As String
        Dim MD5 As New MD5CryptoServiceProvider
        Dim Hash As Byte()
        Dim Result As String = ""
        Dim Tmp As String = ""

        Dim FN As New FileStream(sFile, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        MD5.ComputeHash(FN)
        FN.Close()

        Hash = MD5.Hash
        For i As Integer = 0 To Hash.Length - 1
            Tmp = Hex(Hash(i))
            If Len(Tmp) = 1 Then Tmp = "0" & Tmp
            Result += Tmp
        Next
        Return Result
    End Function

End Class