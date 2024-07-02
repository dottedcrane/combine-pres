Attribute VB_Name = "PPT_AddInns"
Option Explicit
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function ClearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Function

Private Function PickDir() As String
Dim FD As FileDialog
    PickDir = ""
    Set FD = Application.FileDialog(msoFileDialogFolderPicker)
    With FD
        .Title = "Pick a directory to work on..."
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.Count <> 0 Then
            PickDir = .SelectedItems(1)
        End If
    End With
End Function

Private Function PickFile() As String
Dim FD As FileDialog

    PickFile = ""

    Set FD = Application.FileDialog(msoFileDialogFilePicker)
    With FD
        .Title = "Pick a file to work on"
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.Count <> 0 Then
            PickFile = .SelectedItems(1)
        End If
    End With
End Function

Sub Auto_Open()
' Last updated 03/23/2017

      Dim NewControl As Object
      ' Store an object reference to a command bar.
      Dim ToolsMenu As CommandBars

      ' Figure out where to place the menu choice.
      Set ToolsMenu = Application.CommandBars
      
            ' Create the menu choice. The choice is created in the first position in the Tools menu.
      Set NewControl = ToolsMenu("Tools").Controls.Add(Type:=msoControlButton, Before:=1)
      With NewControl
            .DescriptionText = "Combine Multiple Modular PowerPoints into One Large PowerPoint for Presentation Use by Instructors."
            ' Name the command.
            .Caption = "Combine Multiple PPTs"
            ' Connect the menu choice to your macro. The OnAction property
            ' should be set to the name of your macro.
            .OnAction = "CombineMultiplePPTs"
            .Tag = "CombineMultiplePPTs"
            .TooltipText = "Combine Multiple Modular PowerPoints into One Large PowerPoint for Presentation Use by Instructors."
      End With
End Sub

Sub Auto_Close()
      Dim oControl As Object
      Dim ToolsMenu As CommandBarControls
      ' Get an object reference to a command bar.
      Set ToolsMenu = Application.CommandBars("Tools").Controls
      ToolsMenu.Item("Combine Multiple PPTs").Delete
End Sub

Sub CombineMultiplePPTs()
' Last updated 06/16/2015
Dim names() As String
Dim SrcDir, SrcFile, SrcFile1 As String
Dim sMatch As Boolean
Dim Response As String

    ' Check to see whether a presentation is open.
    If Presentations.Count <> 0 Then
        ActiveWindow.ViewType = ppViewNormal
    Else
        MsgBox "No presentation open. Open a presentation and " _
            & "run the macro again.", vbExclamation
    End If
    
    If ActivePresentation.Saved <> msoTrue Then
        Response = MsgBox("Your presentation is not saved. Would you like to save it? ", vbYesNo)
        If Response = vbYes Then
            ActivePresentation.Save
        End If ' End of this one
    End If ' If not saved.

    Response = MsgBox("Did you put all the PowerPoint files, you need to combine in one folder? ", vbYesNo)
    If Response = vbNo Then
        MsgBox "Put all the PowerPoint files you want to combine into one folder and then restart this macro. Stopping macro..."
        Exit Sub
    End If ' End of this one

    Response = MsgBox("Did you number the PPTX files sequentially with 0 prefix, like 01_AboutThisCourse_Tempus_15_1.pptx? ", vbYesNo)
    If Response = vbNo Then
        MsgBox "Please renumber the PPTX files sequentially like 01_AboutThisCourse_Tempus_15_1.pptx, 02_Introduction_Tempus_15_1.pptx and then restart the macro. Stopping macro..."
        Exit Sub
    End If ' End of this one
        
    Response = MsgBox("Will this active presentation, be the first set of slides in the combined presentation? ", vbYesNo)
    If Response = vbNo Then
        MsgBox "Open the file with the first set of slides and restart this macro. Stopping macro..."
        Exit Sub
    End If ' End of this one
    

    SrcDir = PickDir()
    If SrcDir = "" Then Exit Sub
    
    Dim myFileName As String
    myFileName = Application.ActivePresentation.FullName
    
    SrcFile = Dir(SrcDir & "\*.pptx")

    SrcFile1 = SrcDir + "\" + SrcFile
    

    Do While SrcFile <> ""
        If SrcFile1 <> myFileName Then
            ImportSlidesFromPPT SrcDir + "\" + SrcFile
            'ImportSlidesFromPPT is a Sub copying slides into this active presentation.
        End If
        SrcFile = Dir()
        SrcFile1 = SrcDir + "\" + SrcFile
    Loop
    
    MsgBox ("Saving your combined presentation to your specified folder location. If you are not happy with it, delete it before rerunning the macro.")
    ActivePresentation.SaveAs (SrcDir + "\" + "Combined_Presentation.pptx")

End Sub

Private Sub ImportSlidesFromPPT(FileName As String)
' Last updated 09/21/2020

Dim SrcPPT As Presentation, SrcSld As Slide, Idx As Long, SldCnt As Long
Dim SlideFrom, SlideTo As Long
On Error Resume Next
    Set SrcPPT = Presentations.Open(FileName, , , msoFalse)
If Err <> 0 Then
        MsgBox ("Template was not found")
        Exit Sub
End If
    SldCnt = SrcPPT.Slides.Count
    SlideFrom = 1
    SlideTo = SldCnt
    For Idx = SlideFrom To SlideTo Step 1
        Set SrcSld = SrcPPT.Slides(Idx)
        SrcSld.Copy
        With ActivePresentation.Slides.Paste
        End With
    Next Idx
SrcPPT.Close

End Sub
