Imports Microsoft.Win32

Public Class MainForm

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        checkWindowsVersion()
        checkCurrentStatus()
    End Sub

    Private Sub checkWindowsVersion()

        If Not My.Computer.Info.OSFullName.Contains("Windows 10") Then
            MsgBox("Invalid OS !" & vbCrLf & "Windows 10 is requiered !" & vbCrLf & "Installed operating system: " & My.Computer.Info.OSFullName, MsgBoxStyle.Critical, "OS-Exception...")
            Application.Exit()
        End If

    End Sub

    Private Sub RunPatcher_Click(sender As Object, e As EventArgs) Handles RunPatcher.Click

        Try
            Dim req As MsgBoxResult = MsgBox("The author of this tool assumes no responsibility for any resulting damage by this software, at Hard, software or files!" & vbCrLf & vbCrLf & "Do you agree ?", MsgBoxStyle.YesNo, "Securityinformation...")

            If req = MsgBoxResult.Yes Then

                Dim key As RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")

                If RegKeyExists(RegistryHive.CurrentUser, "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme") = True And key.GetValue("AppsUseLightTheme") = 0 Then

                    key.SetValue("AppsUseLightTheme", 1, RegistryValueKind.DWord)

                    MsgBox("Patch uninstalled.", MsgBoxStyle.Information, "Done...")

                    checkCurrentStatus()

                Else

                    key.SetValue("AppsUseLightTheme", 0, RegistryValueKind.DWord)

                    MsgBox("Patch installed.", MsgBoxStyle.Information, "Done...")

                    checkCurrentStatus()

                End If

            Else

                Application.Exit()

            End If
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Fehler...")
        End Try

    End Sub

    Private Sub checkCurrentStatus()

        Try

            Dim key As RegistryKey = My.Computer.Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")

            If RegKeyExists(RegistryHive.CurrentUser, "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme") And key.GetValue("AppsUseLightTheme") = 0 Then
                Status.Text = "Status: Patch active."
            Else
                Status.Text = "Status: Patch inactive."
            End If
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Critical, "Error...")
        End Try


    End Sub

    Private Function RegKeyExists(ByVal hive As RegistryHive, _
                                      ByVal path As String, _
                                      ByVal keyName As String) As Boolean
        Dim regKey As RegistryKey
        Select Case hive
            Case RegistryHive.CurrentUser
                regKey = Registry.CurrentUser.OpenSubKey(path)
            Case RegistryHive.CurrentUser
                regKey = Registry.CurrentUser.OpenSubKey(path)
            Case Else

                Return False
        End Select
        If regKey Is Nothing Then Return False
        For Each regKeyName As String In regKey.GetValueNames()
            If regKeyName.Trim.ToUpper = keyName.Trim.ToUpper Then Return True
        Next
        Return False
    End Function

    Private Sub About_Click(sender As Object, e As EventArgs) Handles About.Click
        MsgBox("Windows 10: AT-Patcher" & vbCrLf & "Author  Jens F." & vbCrLf & "Version  " & My.Application.Info.Version.ToString & vbCrLf & "Alle Rights Reserverd." & vbCrLf & "http://crafter629.de" & vbCrLf & vbCrLf & "This software is not related to Microsoft or its products", MsgBoxStyle.Information, "Über den AT-Patcher...")
    End Sub

    Private Sub IconCreditLink_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles IconCreditLink.LinkClicked
        Process.Start("http://www.iconsmind.com")
    End Sub
End Class
