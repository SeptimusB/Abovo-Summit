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

			DevExpress.XtraEditors.WindowsFormsSettings.SetAccentColor(AbovoBlue)
			Application.EnableVisualStyles()
			Application.SetCompatibleTextRenderingDefault(False)
			BonusSkins.Register()
			SkinManager.EnableFormSkins()

			ApplicationConfiguration.Initilise()

			ApplicationConfiguration.BaseApplicationTitle = "abovo summit"
			Application.Run(New FormMainScreen())

			FormMainScreen.Text = ApplicationConfiguration.BaseApplicationTitle

		End Sub

	End Module
	Public Class ApplicationConfiguration

		Public Shared PrivateBaseApplicationTitle As String
		Public Shared CurrentApplicationTitle As String
		Public Shared CurrentApplicationPath As String
		Public Shared CurrentWorkingDirectory As String
		Public Shared CopyrightMessage As String
        Private Shared WorkingModelID As Integer

		Public Shared DefaultTemplateFile As String = Application.StartupPath & "\Templates\DefaultBPTemplate.xlsb"

		Public Shared Sub Initilise()

			CurrentApplicationPath = Application.StartupPath()
			CurrentWorkingDirectory = CurrentApplicationPath
			CopyrightMessage = "© " & Year(Now()).ToString & " Abovo Business Services Limited.  All rights reserved"

			'ExportServices.initialise()
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

