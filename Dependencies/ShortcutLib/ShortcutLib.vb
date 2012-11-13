Imports System.IO
Imports System.Windows.Forms

Public Module ShortcutLib

    Private Function GetErrorText(ByVal ex As Exception) As String
        Dim err As String = ex.Message
        If ex.InnerException IsNot Nothing Then
            err &= " - More details: " & ex.InnerException.Message
        End If
        Return err
    End Function

    Public Sub CheckForShortcut()
        Try
            If System.Diagnostics.Debugger.IsAttached Then
                Return
            End If
            Dim ad As System.Deployment.Application.ApplicationDeployment
            ad = System.Deployment.Application.ApplicationDeployment.CurrentDeployment

            If (ad.IsFirstRun) Then
                Dim code As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim company As String = String.Empty
                Dim description As String = String.Empty

                If (Attribute.IsDefined(code, GetType(System.Reflection.AssemblyCompanyAttribute))) Then
                    Dim ascompany As System.Reflection.AssemblyCompanyAttribute
                    ascompany = Attribute.GetCustomAttribute(code, GetType(System.Reflection.AssemblyCompanyAttribute))
                    company = ascompany.Company
                End If

                If (Attribute.IsDefined(code, GetType(System.Reflection.AssemblyTitleAttribute))) Then
                    Dim asdescription As System.Reflection.AssemblyTitleAttribute
                    asdescription = Attribute.GetCustomAttribute(code, GetType(System.Reflection.AssemblyTitleAttribute))
                    description = asdescription.Title

                End If

                If (company <> String.Empty And description <> String.Empty) Then
                    'description = Replace(description, "_", " ")

                    Dim desktopPath As String = String.Empty
                    desktopPath = String.Concat(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "\", description, ".appref-ms")

                    Dim shortcutName As String = String.Empty
                    shortcutName = String.Concat(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "\", company, "\", description, ".appref-ms")

                    System.IO.File.Copy(shortcutName, desktopPath, True)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(GetErrorText(ex))
        End Try
    End Sub
End Module
