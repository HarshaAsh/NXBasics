Option Strict Off
Imports System
Imports NXOpen
Imports NXOpenUI
 
Module Module1
 
    Sub Main()
 
        Dim theSession As Session = Session.GetSession()
        Dim theUISession As UI = UI.GetUI
        Dim workPart As Part = theSession.Parts.Work
 
        Dim lw As ListingWindow = theSession.ListingWindow
        lw.Open()
 
        Dim markId1 As Session.UndoMarkId
        markId1 = theSession.SetUndoMark(Session.MarkVisibility.Visible, "journal")
 
        'change excelFileName to meet your needs
        Const excelFileName As String = "E:\Book1.xlsm"
        Dim row As Long = 1
        Dim column As Long = 1
 
        'create Excel object
        Dim objExcel = CreateObject("Excel.Application")
        If objExcel Is Nothing Then
            theUISession.NXMessageBox.Show("Error", NXMessageBox.DialogType.Error, "Could not start Excel, journal exiting")
            theSession.UndoToMark(markId1, "journal")
            Exit Sub
        End If
 
        'open Excel file
        Dim objWorkbook = objExcel.Workbooks.Open(excelFileName)
        If objWorkbook Is Nothing Then
            theUISession.NXMessageBox.Show("Error", NXMessageBox.DialogType.Error, "Could not open Excel file: " & excelFileName & ControlChars.NewLine & "journal exiting.")
            theSession.UndoToMark(markId1, "journal")
            Exit Sub
        End If
 
        objExcel.visible = True
 
        objExcel.Cells(row, 1) = workPart.FullPath
 
        Dim myDimText() As String
        Dim myDimDualText() As String
        For Each myDimension As Annotations.Dimension In workPart.Dimensions
            row += 1
            myDimension.GetDimensionText(myDimText, myDimDualText)
	    objExcel.Cells(row, column) = myDimension.GetType.ToString
            objExcel.Cells(row, column+1) = myDimText(0)
        Next
 
        'objExcel.Quit()
        objWorkbook = Nothing
        objExcel = Nothing
 
        lw.Close()
 
    End Sub
 
 
    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        'Unloads the image when the NX session terminates
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.AtTermination 
    End Function
 
End Module