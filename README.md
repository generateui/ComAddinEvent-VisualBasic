ComAddinEvent
=============

Proof of concept for communication between COM addins within VSTO context

Microsoft Office Word has a Document_Save event. The calling order of subscribers is indeterminate. 
When addin X needs access *before* addin Y to the save event, simply subscribing to the event won't solve the problem as addin Y might be called first.
A solution of this problem is to have addin Y subscribe to an event of addin X.

This proof of concept is written in VisualBasic.NET 10.0. For the C# version, check out https://github.com/generateui/ComAddinEvent.

Requirements:
* Visual Studio 2010
* VSTO april 2014 (newest at time of writing)

To debug:
1. Open and run HasEvent. Word should start and you should see the add-in load.
2. Close Word. The add-in HasEvent is still installed, and will load next time you run Word.
3. Open ConsumeEvent
4. Reference HasEvent.dll in ConsumeEvent. (References → Add reference → lookup HasEvent.dll, usually in HasEvent/bin/Debug)
5. Run ConsumeEvent. Word should start and you should se HasEvent and ConsumeEvent load.
6. Open or New a document, then attempt to save it. You should get a messagebox.
