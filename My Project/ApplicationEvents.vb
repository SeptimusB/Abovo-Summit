Imports Microsoft.VisualBasic.ApplicationServices

Namespace My
    Partial Friend Class MyApplication
        Private Sub AuthoringStartup(sender As Object, e As StartupEventArgs) Handles Me.Startup
            If Not e.CommandLine.Contains("--structure-manager") Then Return
            'No normal main window, Debug auto-open, or shared models in this process.
            e.Cancel = True
            If Not ModelManagerForm.Authorize(Nothing) Then Return
            Abovo.ApplicationConfiguration.Initialize()
            Abovo.ModelCollection.Initialise()
            Abovo.FileManager.Initialise(Nothing)
            Abovo.AbovoAppCls.Initialise()
            Using manager As New StructureManagerForm()
                manager.ShowDialog()
            End Using
        End Sub
    End Class
End Namespace
