' NX 10.0.0.24
' Journal created by Achyuthuni on Sun Jul 08 13:19:44 2018 India Standard Time
'
Option Strict Off
Imports System
Imports NXOpen

Module NXJournal
Sub Main (ByVal args() As String) 

Dim theSession As NXOpen.Session = NXOpen.Session.GetSession()
Dim workPart As NXOpen.Part = theSession.Parts.Work

Dim displayPart As NXOpen.Part = theSession.Parts.Display

' ----------------------------------------------
'   Menu: File->Show Dimensions
' ----------------------------------------------
Dim markId1 As NXOpen.Session.UndoMarkId
markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Show Dimensions")

Dim extrude1 As NXOpen.Features.Extrude = CType(workPart.Features.FindObject("EXTRUDE(2)"), NXOpen.Features.Extrude)

extrude1.ShowDimensions()

Dim nErrs1 As Integer
nErrs1 = theSession.UpdateManager.DoUpdate(markId1)

' ----------------------------------------------
'   Menu: Tools->Journal->Stop Recording
' ----------------------------------------------

End Sub
End Module