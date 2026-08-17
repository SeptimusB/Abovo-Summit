


Imports System.IO

Public Class VideoPlayer
    Dim FilePath As String
    Public Sub New(WhatToPlay As String)

        ' This call is required by the designer.
        InitializeComponent()

        Select Case WhatToPlay
            Case "NavGuide"
                FilePath = Path.Combine(Application.StartupPath, "OpenNav.mp4")
                If Not File.Exists(FilePath) Then
                    File.WriteAllBytes(FilePath, My.Resources.OpenNav)
                End If
        End Select

        With wmp

            .URL = FilePath
            .uiMode = "Full"
            .settings.setMode("loop", True)
            .Ctlcontrols.play()

        End With
        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Sub SetPlay()

    End Sub

End Class

