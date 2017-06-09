Attribute VB_Name = "Module11"
Sub Buttons()
  frmButtons.show vbModeless
End Sub


Sub D_BORDER()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Start a command
    CadInputQueue.SendCommand "DIALOG CELLMAINTENANCE"

    Dim modalHandler As New Macro1ModalHandler3
    AddModalDialogEventsHandler modalHandler

'   The following statement opens modal dialog "Attach Cell Library"

    CadInputQueue.SendCommand "ATTACH LIBRARY"

'   Set a variable associated with a dialog box
    SetCExpressionValue "tcb->activeCell", "BDR-D10", ""

'   Send a keyin that can be a command string
    CadInputQueue.SendKeyin "inputmanager currenttask"

    CadInputQueue.SendCommand "INPUTMANAGER MENU -609 2"

    CadInputQueue.SendCommand "DMSG ACTIVATETOOLBYPATH \Drawing\Cells\Place Active Cell"

    CadInputQueue.SendCommand "PLACE CELL ICON"

    CadInputQueue.SendKeyin "xy=0,0"

'   Send a reset to the current command
    CadInputQueue.SendReset

    RemoveModalDialogEventsHandler modalHandler
    CommandState.StartDefaultCommand
End Sub
Sub CRTS()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Start a command
    CadInputQueue.SendCommand "INPUTMANAGER MENU -609,7"

    CadInputQueue.SendCommand "DMSG ACTIVATETOOLBYPATH \Drawing\Text\Edit Text"

    CadInputQueue.SendCommand "EDIT TEXT"

'   Coordinates are in master units
    startPoint.x = 32.143094
    startPoint.Y = 4.92251
    startPoint.Z = 0#

'   Send a data point to the current command
    point.x = startPoint.x
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

'   Send a message string to an application
'   Content is defined by the application
    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine CURTIS SUBSTATION"

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "NextLine "

    point.x = startPoint.x
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine CURTIS SUBSTATION"

    point.x = startPoint.x + 1.191537
    point.Y = startPoint.Y - 2.839316
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine D-<<     >>"

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "NextLine "

'   Send a keyin that can be a command string
    CadInputQueue.SendKeyin Chr$(27)

    CadInputQueue.SendCommand "INPUTMANAGER MENU -609,7"

    CadInputQueue.SendCommand "DMSG ACTIVATETOOLBYPATH \Drawing\Text\Fill In Single Enter-Data Field"

    CadInputQueue.SendCommand "EDIT SINGLE DIALOG"

    point.x = startPoint.x + 1.173366
    point.Y = startPoint.Y - 2.799316
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine CON     "

    point.x = startPoint.x + 1.173366
    point.Y = startPoint.Y - 2.799316
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x + 0.144069
    point.Y = startPoint.Y + 10.062997
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine AWO 0000026594.                                                    "

  

    point.x = startPoint.x - 0.242137
    point.Y = startPoint.Y + 16.100209
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 0.268499
    point.Y = startPoint.Y - 3.622898
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine CRTS-         "

    point.x = startPoint.x + 0.268499
    point.Y = startPoint.Y - 3.622898
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    CommandState.StartDefaultCommand
End Sub



Sub Macro1()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Set a variable associated with a dialog box
    SetCExpressionValue "plotUI.uiPlotArea", 2, "PLOTDLG"

'   Coordinates are in master units
    startPoint.x = 35.175694
    startPoint.Y = 4.473955
    startPoint.Z = 0#

'   Send a data point to the current command
    point.x = startPoint.x
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 0.463557
    point.Y = startPoint.Y - 0.42816
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

'   Start a command
    CadInputQueue.SendCommand "MDL SILENTLOAD USTNVBA IDE"

    SetCExpressionValue "plotUI.uiPlotArea", 3, "PLOTDLG"

'   Send a keyin that can be a command string
    CadInputQueue.SendKeyin "VBA RUN BUTTONS"

    point.x = startPoint.x + 2.448534
    point.Y = startPoint.Y + 1.35584
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 0.237722
    point.Y = startPoint.Y + 1.28448
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 2.115724
    point.Y = startPoint.Y + 2.652213
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.33281
    point.Y = startPoint.Y + 1.71264
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1
End Sub
