Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Word
Imports Office = Microsoft.Office.Core

''' The observer pattern is needed to do eventing in COM. That's because
''' COM does not recognizes delegates and therefore does not recognize native 
''' .NET events.
''' El Zorko explains this better in 
''' http://stackoverflow.com/questions/1985451/exposing-a-net-class-which-has-events-to-com
''' <seealso cref="http://en.wikipedia.org/wiki/Observer_pattern"/>
''' <summary>
''' Third-party interested in the Save event
''' </summary>
<ComVisible(True), _
 Guid("12810C22-3FF2-4fc2-A7FD-7E1034462EB0")> _
Public Interface IObserver

	' Called by the first party when the first party handled the
	' Word interop BeforeSave event
	Sub AfterSave()
End Interface

''' <summary>
''' The first party needing to gain first-fire access to the BeforeSave event
''' </summary>
<ComVisible(True), _
 Guid("02810C22-3FF2-4fc2-A7FD-7E1034462EB0")> _
Public Interface ISubject

	Sub Listen(ByVal observer As IObserver)

	Sub Fire()
End Interface

<ComVisible(True), _
 Guid("02810C22-3FF2-4fc2-A7FD-5E1034462EB0"), _
 ClassInterface(ClassInterfaceType.None)> _
Partial Public Class ThisAddIn
	Implements ISubject

	Private _observers As List(Of IObserver) = New List(Of IObserver)

	Public Sub Listen(ByVal observer As IObserver) Implements ISubject.Listen
		_observers.Add(observer)
	End Sub

	Private Sub Application_DocumentBeforeSave(ByVal Doc As Document, ByRef SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Application.DocumentBeforeSave
		Fire()
	End Sub

	Private Sub ThisAddIn_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs)

	End Sub

	Public Sub Fire() Implements ISubject.Fire
		For Each observer As IObserver In _observers
			observer.AfterSave()
		Next
	End Sub

	''' <summary>
	''' Required (and undocumented on the COMAddin MSDN docs) override.
	''' </summary>
	''' <returns></returns>
	Protected Overrides Function RequestComAddInAutomationService() As Object
		Return Me
	End Function

End Class