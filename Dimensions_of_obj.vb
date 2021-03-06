Option Strict Off
Imports System
Imports NXOpen
Imports NXOpen.UF
 
Module NXJournal
Sub Main
 
Dim theSession As Session = Session.GetSession()
Dim workPart As Part = theSession.Parts.Work
Dim displayPart As Part = theSession.Parts.Display
Dim lw As ListingWindow = theSession.ListingWindow
Dim mySelectedObject As NXObject
 
lw.Open
 
mySelectedObject = getGewaehlteVerbindung
mySelectedObject.ShowDimensions()
 
'these are the selections
lw.WriteLine("Object Tag: " & mySelectedObject.Tag)
lw.WriteLine("Object Type: " & mySelectedObject.GetType.ToString)
lw.WriteLine("")
For Each tempDim As Annotations.Dimension In workPart.Dimensions
 	Dim myDimText() As String
        Dim myDimDualText() As String
   	lw.WriteLine("dimension type is  " & tempDim.GetType.ToString)
	tempDim.GetDimensionText(myDimText, myDimDualText)
	lw.WriteLine("dimension value is  " & myDimText(0))
Next


lw.Close
 
End Sub
 
Public Function getGewaehlteVerbindung() as NXObject
	Dim returnGewaehlteVerbindung As NXObject ' Ut: War daov rMIcrosoft COllection
	SelectAnObject("Wählen Sie bitte die zu ändernde Verbindung aus", returnGewaehlteVerbindung)
Return returnGewaehlteVerbindung
 
End Function
 
    Function SelectAnObject(prompt As String, _
               ByRef selObj As NXObject) As Selection.Response
 
       Dim theUI As UI = UI.GetUI
       Dim cursor As Point3d
       Dim typeArray() As Selection.SelectionType = _
           {Selection.SelectionType.All, _
               Selection.SelectionType.Faces, _
               Selection.SelectionType.Edges, _
               Selection.SelectionType.Features}
 
       Dim resp As Selection.Response = theUI.SelectionManager.SelectObject( _
               prompt, "Selection", _
               Selection.SelectionScope.AnyInAssembly, _
               False, typeArray, selobj, cursor)
 
       If resp = Selection.Response.ObjectSelected Or _
               resp = Selection.Response.ObjectSelectedByName Then
           Return Selection.Response.Ok
       Else
           Return Selection.Response.Cancel
       End If
 
    End Function
 
End Module

