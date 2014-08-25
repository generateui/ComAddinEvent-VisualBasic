Imports System
Imports System.Windows.Forms
Imports HasEvent
Imports Office = Microsoft.Office.Core

Public Class ThisAddIn
	Implements IObserver

	Private hasEvent As Office.COMAddIn

	Public Sub AfterSave() Implements IObserver.AfterSave  '' Implements IObserver.AfterSave
		MessageBox.Show("Before save event fired")
	End Sub

	Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Startup
		Dim comAddin As Office.COMAddIn = FindComAddinByName("HasEvent")
		Dim hd As ISubject = ObtainSubject(comAddin)
		hd.Listen(Me)
	End Sub

	Private Function FindComAddinByName(ByVal name As String) As Office.COMAddIn
		For Each comAddIn As Office.COMAddIn In Application.COMAddIns
			If (comAddIn.Description = name) Then
				Return comAddIn
			End If
		Next
		Return Nothing
	End Function

	''' <summary>
	''' Obtains a reference to a typed ComAddin project.
	''' The type is determined in the other ComAddin project.
	''' </summary>
	''' <seealso cref="http://blogs.msdn.com/b/andreww/archive/2008/08/13/comaddins-race-condition.aspx"/>
	''' <param name="comAddin"></param>
	''' <returns></returns>
	Private Function ObtainSubject(ByVal comAddin As Office.COMAddIn) As ISubject
		Dim obj As Object = Nothing
		Dim tries As Integer = 0
		' 50 * 100 miliseconds = 5000 milliseconds == 5 seconds

		While ((obj Is Nothing) _
					AndAlso (tries < 50))
			obj = comAddin.Object
			System.Threading.Thread.Sleep(100)
			tries = (tries + 1)

		End While
		Dim subject As ISubject = CType(obj, ISubject)
		Return subject
	End Function

	Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Shutdown

	End Sub
End Class
