Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Windows.Forms
Imports DevExpress.UserSkins
Imports DevExpress.Skins
Imports DevExpress.LookAndFeel
Imports Abovo.GeneralFunctions

Namespace Abovo

	Friend Module Program
		''' <summary>
		''' The main entry point for the application.
		''' </summary>
		''' 
		<STAThread>
		Sub Main()

			Application.EnableVisualStyles()
			Application.SetCompatibleTextRenderingDefault(False)
			DevExpress.XtraEditors.WindowsFormsSettings.SetAccentColor(AbovoBlue)
			BonusSkins.Register()
			SkinManager.EnableFormSkins()

			ApplicationConfiguration.Initialize()
			ApplicationConfiguration.BaseApplicationTitle = "abovo summit"

			Using MainForm As New FormMainScreen()
				MainForm.Text = ApplicationConfiguration.BaseApplicationTitle
				Application.Run(MainForm)
			End Using

		End Sub

	End Module
	Public Class ApplicationConfiguration

		Public Shared PrivateBaseApplicationTitle As String
		Public Shared CurrentApplicationTitle As String
		Public Shared CurrentApplicationPath As String
		Public Shared CurrentWorkingDirectory As String
		Public Shared CopyrightMessage As String
        Private Shared WorkingModelID As Integer

		Public Shared ReadOnly Property DefaultTemplateFile As String
			Get
				Return IO.Path.Combine(CurrentApplicationPath, "Templates", "DefaultBPTemplate.xlsb")
			End Get
		End Property

		Public Shared Sub Initialize()

			CurrentApplicationPath = Application.StartupPath()
			CurrentWorkingDirectory = CurrentApplicationPath
			CopyrightMessage = "© " & DateTime.Now.Year.ToString & " Abovo Business Services Limited.  All rights reserved"

			'ExportServices.initialise()
		End Sub

		<Obsolete("Use Initialize instead.")>
		Public Shared Sub Initilise()
			Initialize()
		End Sub
		Public Shared Property ActiveModelID As Integer
			Set(value As Integer)
				WorkingModelID = value

			End Set
			Get
				Return WorkingModelID
			End Get
		End Property
		Public Shared Property BaseApplicationTitle As String
			Set(value As String)
				PrivateBaseApplicationTitle = value
				CurrentApplicationTitle = value
			End Set
			Get
				Return PrivateBaseApplicationTitle
			End Get
		End Property
		Public Shared Property ExtendedApplicationTitle As String
			Set(value As String)
				CurrentApplicationTitle = BaseApplicationTitle & value
			End Set
			Get
				Return CurrentApplicationTitle
			End Get
		End Property
	End Class

End Namespace

