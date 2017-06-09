Attribute VB_Name = "Module1_old"
Sub Macro1()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

    Dim modalHandler As New Macro1ModalHandler5
    AddModalDialogEventsHandler modalHandler

'   The following statement opens modal dialog "Preferences [descartes]"

'   Start a command
    CadInputQueue.SendCommand "MDL SILENTLOAD USERPREF"

    CadInputQueue.SendCommand "MDL SILENTUNLOAD SPELLCHECK"

    RemoveModalDialogEventsHandler modalHandler
    CommandState.StartDefaultCommand
End Sub
Sub Macro2()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

    Dim modalHandler As New Macro2ModalHandler1
    AddModalDialogEventsHandler modalHandler

'   The following statement opens modal dialog "Preferences [descartes]"

'   Start a command
    CadInputQueue.SendCommand "MDL SILENTLOAD USERPREF"

    CadInputQueue.SendCommand "MDL SILENTUNLOAD SPELLCHECK"

    RemoveModalDialogEventsHandler modalHandler
    CommandState.StartDefaultCommand
End Sub
Sub Macro3()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Start a command
    CadInputQueue.SendCommand "DIALOG PLOT"

    Dim modalHandler As New Macro3ModalHandler
    AddModalDialogEventsHandler modalHandler

'   The following statement opens modal dialog "Print - Raster Options"

    CadInputQueue.SendCommand "PRINT ROPTSDIALOG"

    RemoveModalDialogEventsHandler modalHandler
    CommandState.StartDefaultCommand
End Sub
Sub Macro4()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Send a keyin that can be a command string
    CadInputQueue.SendKeyin "level purge all"

    Dim modalHandler As New Macro4ModalHandler1
    AddModalDialogEventsHandler modalHandler

'   The following statement opens modal dialog "Design File Settings"

'   Start a command
    CadInputQueue.SendCommand "MDL SILENTLOAD DGNSET"

    CadInputQueue.SendCommand "FILEDESIGN"

    CadInputQueue.SendCommand "MDL SILENTUNLOAD DGNSET"

    RemoveModalDialogEventsHandler modalHandler
    CommandState.StartDefaultCommand
End Sub
Sub Macro5()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

    Dim modalHandler As New Macro5ModalHandler0
    AddModalDialogEventsHandler modalHandler

'   The following statement opens modal dialog "Color Table"

'   Start a command
    CadInputQueue.SendCommand "DIALOG COLOR"

'   Coordinates are in master units
    startPoint.x = 2.95957877203563
    startPoint.Y = 0.120543355820554
    startPoint.Z = 0.083333333333315

'   Send a data point to the current command
    point.x = startPoint.x
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.178317424247461
    point.Y = startPoint.Y + 0.414411330316334
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "DELETE ELEMENT"

    point.x = startPoint.x - 0.111822281528084
    point.Y = startPoint.Y - 5.83859518712345E-02
    point.Z = startPoint.Z - 2.3592E-16
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.241407279381207
    point.Y = startPoint.Y + 0.130217249126066
    point.Z = startPoint.Z - 2.3592E-16
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "DELETE ELEMENT"

    CadInputQueue.SendCommand "MDL SILENTUNLOAD VBAPM"

'   The following statement opens modal dialog "Color Table"

    CadInputQueue.SendCommand "DIALOG COLOR"

'   The following statement opens modal dialog "Design File Settings"

    CadInputQueue.SendCommand "MDL SILENTLOAD DGNSET"

    CadInputQueue.SendCommand "FILEDESIGN"

    CadInputQueue.SendCommand "MDL SILENTUNLOAD DGNSET"

'   The following statement opens modal dialog "Level/Filter Import"

'   The following statement opens modal dialog "Import Levels"

    CadInputQueue.SendCommand "LEVELMANAGER LIBRARY IMPORT"

    point.x = startPoint.x + 0.40227260475232
    point.Y = startPoint.Y - 0.312869042962989
    point.Z = startPoint.Z + 0#
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 0.713249564377427
    point.Y = startPoint.Y - 2.95800476295584E-02
    point.Z = startPoint.Z + 0#
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

'   Set a variable associated with a dialog box
    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    point.x = startPoint.x - 2.55371618863213
    point.Y = startPoint.Y + 1.78409905360514
    point.Z = startPoint.Z + 1.80411E-15
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 0.213079974691378
    point.Y = startPoint.Y - 0.414583282239851
    point.Z = startPoint.Z + 1.80411E-15
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 3.5085772346024
    point.Y = startPoint.Y + 2.19851038392148
    point.Z = startPoint.Z + 1.80411E-15
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 0.207327799715654
    point.Y = startPoint.Y - 0.35702615302925
    point.Z = startPoint.Z + 1.80411E-15
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "Change Attributes"

'   Send a keyin that can be a command string
    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES USEACTIVE ON"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE LEVEL"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET LEVEL ""New or Revisions"""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE COLOR"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET COLOR ""0"""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE LINESTYLE"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET LINESTYLE ""Continuous"""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE WEIGHT"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET WEIGHT 0"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE TRANSPARENCY"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET TRANSPARENCY 0"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE PRIORITY"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET PRIORITY 0"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE ELEMENTCLASS"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET ELEMENTCLASS PRIMARY"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE TEMPLATE"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET TEMPLATE """""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES MAKECOPY OFF"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENTIREELEMENT OFF"

    SetCExpressionValue "tcb->msToolSettings.general.useFence", 0, "CHANGEATTRIBS"

    CadInputQueue.SendCommand "LOCK FENCE INSIDE"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET LEVEL ""Vendor"""

    point.x = startPoint.x + 0.115293000104061
    point.Y = startPoint.Y + 0.926497828367175
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    point.x = startPoint.x - 3.20191440620909
    point.Y = startPoint.Y + 1.94525901539483
    point.Z = startPoint.Z + 2.34535E-15
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.631411213932183
    point.Y = startPoint.Y - 6.28739690214429E-03
    point.Z = startPoint.Z + 2.34535E-15
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    CadInputQueue.SendCommand "NEWFILE U:\New folder\bellSWGRTITLEBLOCK.dgn"

    CadInputQueue.SendKeyin "task sendtaskchangedasync"

    CadInputQueue.SendKeyin "task sendtaskchangedasync ""\Drawing"""

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    point.x = startPoint.x - -39.6470393651277
    point.Y = startPoint.Y - 5.9953998274684
    point.Z = startPoint.Z + 2.45666666666633
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -17.8879261258157
    point.Y = startPoint.Y + 13.083591904566
    point.Z = startPoint.Z + 2.45666666666633
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "MDL LOAD CLIPBRD COPY"

    CadInputQueue.SendCommand "NEWFILE ""P:\Active Projects\PGE\Substation\6446 BELL\2000 Substation\2300 Engineering\2310 Electrical\2311 Drawings\Indoor\bell7313a0.dgn"",""~4683"""

    CadInputQueue.SendKeyin "task sendtaskchangedasync"

    CadInputQueue.SendKeyin "task sendtaskchangedasync ""\Drawing"""

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    CadInputQueue.SendCommand "MDL KEYIN CLIPBRD CLIPBOARD PASTE"

    point.x = startPoint.x - 28.631953394003
    point.Y = startPoint.Y + 16.2454221816907
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

'   Send a reset to the current command
    CadInputQueue.SendReset

    point.x = startPoint.x - -9.72067325664033
    point.Y = startPoint.Y - 10.5493999921612
    point.Z = startPoint.Z + 2.45666666666622
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "MDL SILENTLOAD USTNVBA MACROS"

    CadInputQueue.SendReset

    point.x = startPoint.x - 6.70546566878871
    point.Y = startPoint.Y + 56.2127948024225
    point.Z = startPoint.Z + 2.45666666666622
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 93.0284853312103
    point.Y = startPoint.Y - 26.450259538265
    point.Z = startPoint.Z + 2.45666666666622
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "SCALE ICON"

    CadInputQueue.SendCommand "ACTIVE XSCALE 0.3900"

    CadInputQueue.SendCommand "ACTIVE SCALE"

    point.x = startPoint.x - -15.9807978633999
    point.Y = startPoint.Y - 9.31802029667172
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "MOVE ICON"

'   Send a tentative point
    CadInputQueue.SendTentativePoint Point3dFromXYZ(43.3222605049795, 15.6762988369673, 2.53999999999963), 1

    point.x = startPoint.x - -40.2932509755801
    point.Y = startPoint.Y + 15.5668957102602
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendAdjustedDataPoint point, 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-10.8485751025206, 26.7805644149428, 0#), 1

    point.x = startPoint.x - 13.757003394003
    point.Y = startPoint.Y + 26.7454221816907
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendAdjustedDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    point.x = startPoint.x - 17.8581020925771
    point.Y = startPoint.Y + 6.35711818150786
    point.Z = startPoint.Z + 2.45666666666631
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 18.0026820661479
    point.Y = startPoint.Y + 6.73084617122634
    point.Z = startPoint.Z + 2.45666666666631
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 2, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR MODE REMOVE"

    point.x = startPoint.x - 19.2263406966295
    point.Y = startPoint.Y + 5.90013022633797
    point.Z = startPoint.Z + 2.45666666666631
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 19.4851915607698
    point.Y = startPoint.Y + 7.26581970489595
    point.Z = startPoint.Z + 2.45666666666631
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "MOVE ICON"

    point.x = startPoint.x - 18.457003394003
    point.Y = startPoint.Y + 6.8954221816907
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 17.7338536777898
    point.Y = startPoint.Y + 7.65405225873153
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "MOVE ICON"

    point.x = startPoint.x - 17.4179055416471
    point.Y = startPoint.Y + 7.88136182113021
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 17.0881472557046
    point.Y = startPoint.Y + 6.44033105421862
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    point.x = startPoint.x - 18.0520137461763
    point.Y = startPoint.Y + 5.93398990685809
    point.Z = startPoint.Z + 2.45666666666631
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 22.0994999854615
    point.Y = startPoint.Y + 11.4532591098579
    point.Z = startPoint.Z + 2.45666666666631
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 1, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR MODE ADD"

    point.x = startPoint.x - 17.9014096070401
    point.Y = startPoint.Y + 5.91515281060212
    point.Z = startPoint.Z + 2.45666666666631
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 22.1936275724216
    point.Y = startPoint.Y + 11.2837252435542
    point.Z = startPoint.Z + 2.45666666666631
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "MOVE ICON"

    point.x = startPoint.x - 21.8924192941492
    point.Y = startPoint.Y + 11.1330284735064
    point.Z = startPoint.Z + 2.45666666666631
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 21.9213352888634
    point.Y = startPoint.Y + 11.8660175630188
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    CadInputQueue.SendCommand "EDIT SINGLE DIALOG"

    point.x = startPoint.x - 12.979263944635
    point.Y = startPoint.Y + 8.66307073823081
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

'   Send a message string to an application
'   Content is defined by the application
    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine BELL SUBSTATION #3                                      "

    point.x = startPoint.x - 12.979263944635
    point.Y = startPoint.Y + 8.66307073823081
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - 12.8346839710643
    point.Y = startPoint.Y + 8.8197953790805
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine 15KV OUTDOOR SWITCHGEAR                                      "

    point.x = startPoint.x - 12.8105873088025
    point.Y = startPoint.Y + 8.8197953790805
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - 12.4491373748756
    point.Y = startPoint.Y + 8.85596260389197
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine UNIT 4 - FDR. BKR R304 CONTROL SCHEM.                                      "

    point.x = startPoint.x - 12.4491373748756
    point.Y = startPoint.Y + 8.85596260389197
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - 11.1479176127389
    point.Y = startPoint.Y + 8.94035279511872
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 12.7764830014951
    point.Y = startPoint.Y + 7.54338620193733
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine 3 "

    point.x = startPoint.x - 12.7764830014951
    point.Y = startPoint.Y + 7.54338620193733
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - 13.3702247596256
    point.Y = startPoint.Y + 7.41221973328774
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine BELL-7313"

    point.x = startPoint.x - 14.4343333651063
    point.Y = startPoint.Y + 7.12673977210923
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 5, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR ALL"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    CadInputQueue.SendKeyin "VBA RUN BUTTONS"

    point.x = startPoint.x - 57.5043080250155
    point.Y = startPoint.Y + 33.4859425904581
    point.Z = startPoint.Z + 2.45666666666631
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 8.85056034016884
    point.Y = startPoint.Y - -0.431345583671012
    point.Z = startPoint.Z + 2.45666666666631
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-44.1353407261628, 5.16875384750476, 2.53999999999963), 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-43.043773250429, 4.99629508051283, 2.53999999999963), 1

    point.x = startPoint.x - 45.757003394003
    point.Y = startPoint.Y + 5.2454221816907
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendAdjustedDataPoint point, 1

    point.x = startPoint.x - 48.0141342146059
    point.Y = startPoint.Y + 7.92252327488308
    point.Z = startPoint.Z + 2.45666666666631
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendKeyin Chr$(27)

    CadInputQueue.SendCommand "MOVE ICON"

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-42.9458768831316, 4.95628464657069, 2.53999999999963), 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-42.7105579157313, 5.42721205296999, 2.53999999999963), 1

    point.x = startPoint.x - 45.757003394003
    point.Y = startPoint.Y + 5.2454221816907
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendAdjustedDataPoint point, 1

    CadInputQueue.SendKeyin "xy=0,0"

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "VIEW ON 5"

    CadInputQueue.SendKeyin "dialog viewsettings popup"

    CadInputQueue.SendKeyin "MDL KEYIN BENTLEY.VIEWATTRIBUTESDIALOG,VAD VIEWATTRIBUTESDIALOG SETATTRIBUTE 0 DataFields False"

    point.x = startPoint.x - -31.2800461097629
    point.Y = startPoint.Y + 4.42010573726455
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -32.3080604127794
    point.Y = startPoint.Y + 6.25881152715817
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "PRINT EXECUTE"

    point.x = startPoint.x - -32.1410080885392
    point.Y = startPoint.Y + 6.49025701120072
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    point.x = startPoint.x - -31.7426525461203
    point.Y = startPoint.Y + 6.2973857744986
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

'   The following statement opens modal dialog "Open"

    CadInputQueue.SendCommand "DIALOG OPENFILE"

    point.x = startPoint.x - -28.1960032007134
    point.Y = startPoint.Y + 4.6386931388603
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    point.x = startPoint.x - -27.3221910431493
    point.Y = startPoint.Y - -2.54282570003051
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

'   The following statement opens modal dialog "Open"

    CadInputQueue.SendCommand "DIALOG OPENFILE"

    CadInputQueue.SendKeyin "task sendtaskchangedasync"

    CadInputQueue.SendKeyin "task sendtaskchangedasync ""\Drawing"""

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    point.x = startPoint.x + 0.193136870243662
    point.Y = startPoint.Y - 0.277476094811245
    point.Z = startPoint.Z + 0#
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.285178612605209
    point.Y = startPoint.Y + 0.491060597448877
    point.Z = startPoint.Z + 0#
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "DELETE ELEMENT"

'   The following statement opens modal dialog "Design File Settings"

    CadInputQueue.SendCommand "MDL SILENTLOAD DGNSET"

    CadInputQueue.SendCommand "MDL SILENTUNLOAD DGNSET"

'   The following statement opens modal dialog "Color Table"

    CadInputQueue.SendCommand "DIALOG COLOR"

    point.x = startPoint.x - 3.6622037756937
    point.Y = startPoint.Y + 2.70297650225928
    point.Z = startPoint.Z + 0#
    CadInputQueue.SendDataPoint point, 1

'   The following statement opens modal dialog "Design File Settings"

    CadInputQueue.SendCommand "MDL SILENTLOAD USTNVBA MACROS"

    CadInputQueue.SendCommand "FILEDESIGN"

    CadInputQueue.SendCommand "MDL SILENTUNLOAD DGNSET"

    CadInputQueue.SendKeyin "level purge all"

'   The following statement opens modal dialog "Design File Settings"

    CadInputQueue.SendCommand "MDL SILENTLOAD DGNSET"

    CadInputQueue.SendCommand "FILEDESIGN"

    CadInputQueue.SendCommand "MDL SILENTUNLOAD DGNSET"

'   The following statement opens modal dialog "Alert"

    CadInputQueue.SendCommand "UNDO ALL"

    point.x = startPoint.x + 1.02681143702998
    point.Y = startPoint.Y + 0.608699634550071
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 0.99088028597463
    point.Y = startPoint.Y + 0.698582761919565
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 1.04477701255766
    point.Y = startPoint.Y + 0.761500951078211
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    point.x = startPoint.x + 0.936983559391602
    point.Y = startPoint.Y + 0.700380424466956
    point.Z = startPoint.Z + 0#
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 0.784276167406347
    point.Y = startPoint.Y + 0.853181740995096
    point.Z = startPoint.Z + 0#
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.125320921559938
    point.Y = startPoint.Y - 0.256515349508679
    point.Z = startPoint.Z + 0#
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.280543494119067
    point.Y = startPoint.Y + 0.117398460348417
    point.Z = startPoint.Z + 0#
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "DELETE ELEMENT"

'   The following statement opens modal dialog "Color Table"

    CadInputQueue.SendCommand "DIALOG COLOR"

    CadInputQueue.SendKeyin "level purge all"

    point.x = startPoint.x - 2.81984285313975
    point.Y = startPoint.Y + 1.81724569437845
    point.Z = startPoint.Z - 1.388E-17
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 0.672215890052462
    point.Y = startPoint.Y - 1.3736053272386
    point.Z = startPoint.Z - 1.388E-17
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendKeyin Chr$(27)

    CadInputQueue.SendCommand "SCALE ICON"

    CadInputQueue.SendCommand "ACTIVE XSCALE 0.3900"

    CadInputQueue.SendCommand "ACTIVE SCALE"

    point.x = startPoint.x - 2.00016346968949
    point.Y = startPoint.Y - 0.800600390258071
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

'   The following statement opens modal dialog "Design File Settings"

    CadInputQueue.SendCommand "MDL SILENTLOAD DGNSET"

    CadInputQueue.SendCommand "FILEDESIGN"

    CadInputQueue.SendCommand "MDL SILENTUNLOAD DGNSET"

'   The following statement opens modal dialog "Color Table"

    CadInputQueue.SendCommand "DIALOG COLOR"

    point.x = startPoint.x - 0.385769666542224
    point.Y = startPoint.Y - 0.422631053694063
    point.Z = startPoint.Z + 0#
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.459356663903589
    point.Y = startPoint.Y - 0.290092989400102
    point.Z = startPoint.Z + 0#
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    point.x = startPoint.x - 2.44252624279237
    point.Y = startPoint.Y + 0.265830558055125
    point.Z = startPoint.Z + 0#
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.830971000578482
    point.Y = startPoint.Y - 0.919648794796418
    point.Z = startPoint.Z + 0#
    CadInputQueue.SendDataPoint point, 1

'   The following statement opens modal dialog "Level/Filter Import"

'   The following statement opens modal dialog "Import Levels"

    CadInputQueue.SendCommand "LEVELMANAGER LIBRARY IMPORT"

    point.x = startPoint.x - 0.301144619576654
    point.Y = startPoint.Y - 0.595666859855624
    point.Z = startPoint.Z + 0#
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "Change Attributes"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES USEACTIVE ON"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE LEVEL"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET LEVEL ""Vendor"""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE COLOR"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET COLOR ""0"""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE LINESTYLE"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET LINESTYLE ""Continuous"""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE WEIGHT"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET WEIGHT 0"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE TRANSPARENCY"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET TRANSPARENCY 0"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE PRIORITY"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET PRIORITY 0"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE ELEMENTCLASS"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET ELEMENTCLASS PRIMARY"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE TEMPLATE"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET TEMPLATE """""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES MAKECOPY OFF"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENTIREELEMENT OFF"

    SetCExpressionValue "tcb->msToolSettings.general.useFence", 0, "CHANGEATTRIBS"

    CadInputQueue.SendCommand "LOCK FENCE INSIDE"

    point.x = startPoint.x - 0.551340410605294
    point.Y = startPoint.Y - 0.367406860238246
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    point.x = startPoint.x - 2.21647618527293
    point.Y = startPoint.Y + 0.190587594471574
    point.Z = startPoint.Z + 7.0777E-16
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 1.18165903487874
    point.Y = startPoint.Y - 0.913896274644771
    point.Z = startPoint.Z + 7.0777E-16
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    CadInputQueue.SendCommand "NEWFILE U:\New folder\bellSWGRTITLEBLOCK.dgn"

    CadInputQueue.SendKeyin "task sendtaskchangedasync"

    CadInputQueue.SendKeyin "task sendtaskchangedasync ""\Drawing"""

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    point.x = startPoint.x - 6.22655504146303
    point.Y = startPoint.Y + 26.8878035695086
    point.Z = startPoint.Z + 2.45666666666633
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -37.5159921922054
    point.Y = startPoint.Y - 8.57667517944953
    point.Z = startPoint.Z + 2.45666666666633
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "MDL LOAD CLIPBRD COPY"

    point.x = startPoint.x - -33.2538978463608
    point.Y = startPoint.Y - 11.6068679839491
    point.Z = startPoint.Z + 2.45666666666633
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -32.6930959587497
    point.Y = startPoint.Y - 10.5968037157826
    point.Z = startPoint.Z + 2.45666666666633
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -25.5148317973271
    point.Y = startPoint.Y - 11.1579505314307
    point.Z = startPoint.Z + 2.45666666666633
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "NEWFILE P:\Active Projects\PGE\Substation\6446 BELL\2000 Substation\2300 Engineering\2310 Electrical\2311 Drawings\Indoor\bell7313b0.dgn"

    CadInputQueue.SendKeyin "task sendtaskchangedasync"

    CadInputQueue.SendKeyin "task sendtaskchangedasync ""\Drawing"""

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    CadInputQueue.SendCommand "MDL KEYIN CLIPBRD CLIPBOARD PASTE"

    CadInputQueue.SendCommand "ACTIVE ANGLE 0.0000°"

    CadInputQueue.SendCommand "ACTIVE ANGLE"

    CadInputQueue.SendCommand "ACTIVE XSCALE 1.0000"

    CadInputQueue.SendCommand "ACTIVE SCALE"

    point.x = startPoint.x - 4.70169394288028
    point.Y = startPoint.Y - 1.99909755251633
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    point.x = startPoint.x - -15.792365326435
    point.Y = startPoint.Y + 10.7901239732489
    point.Z = startPoint.Z + 2.45666666666614
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -51.9735448359856
    point.Y = startPoint.Y - 19.4560095039993
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "MOVE ICON"

    CadInputQueue.SendTentativePoint Point3dFromXYZ(49.5509086694719, 8.81060678696583, 2.53999999999945), 1

    point.x = startPoint.x - -46.5778380614257
    point.Y = startPoint.Y + 8.53319221428959
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendAdjustedDataPoint point, 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(13.1579450193187, 8.54097098153532, 0#), 1

    point.x = startPoint.x - -10.1732560571197
    point.Y = startPoint.Y + 8.50090244748367
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendAdjustedDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    CadInputQueue.SendKeyin "VBA RUN BUTTONS"

    CadInputQueue.SendCommand "VIEW ON 5"

    CadInputQueue.SendCommand "PLACE FENCE ICON"

    point.x = startPoint.x - 25.7595787720356
    point.Y = startPoint.Y + 12.5794566441794
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    point.x = startPoint.x - 29.0959011922633
    point.Y = startPoint.Y + 11.981172036036
    point.Z = startPoint.Z - -2.45666666666612
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -15.9025063967995
    point.Y = startPoint.Y - 22.2194920782503
    point.Z = startPoint.Z - -2.45666666666612
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "MOVE ICON"

    point.x = startPoint.x - 21.8267439428803
    point.Y = startPoint.Y - 12.9990975525163
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendKeyin "xy=0,0"

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    CadInputQueue.SendCommand "NEWFILE U:\New folder\bellSWGRTITLEBLOCK.dgn"

    CadInputQueue.SendKeyin "task sendtaskchangedasync"

    CadInputQueue.SendKeyin "task sendtaskchangedasync ""\Drawing"""

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

'   The following statement opens modal dialog "Compress Options"

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

'   The following statement opens modal dialog "Level/Filter Import"

'   The following statement opens modal dialog "Import Levels"

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    point.x = startPoint.x - -31.410299380563
    point.Y = startPoint.Y + 4.80849955322854
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -30.6764212172336
    point.Y = startPoint.Y + 5.59931598527109
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -27.1952042886195
    point.Y = startPoint.Y - -1.43811523571386
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -26.7059521797332
    point.Y = startPoint.Y - -2.66199780911304
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "Change Attributes"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES USEACTIVE ON"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE LEVEL"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET LEVEL ""Vendor"""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE COLOR"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET COLOR ""0"""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE LINESTYLE"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET LINESTYLE ""Continuous"""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE WEIGHT"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET WEIGHT 0"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE TRANSPARENCY"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET TRANSPARENCY 0"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE PRIORITY"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET PRIORITY 0"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE ELEMENTCLASS"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET ELEMENTCLASS PRIMARY"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE TEMPLATE"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET TEMPLATE """""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES MAKECOPY OFF"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENTIREELEMENT OFF"

    SetCExpressionValue "tcb->msToolSettings.general.useFence", 0, "CHANGEATTRIBS"

    CadInputQueue.SendCommand "LOCK FENCE INSIDE"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET LEVEL ""Border-titleblock"""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE COLOR"

    point.x = startPoint.x - -34.7268641571481
    point.Y = startPoint.Y + 8.65431517810213
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    CadInputQueue.SendKeyin "level purge all"

    point.x = startPoint.x - -37.7780284085457
    point.Y = startPoint.Y - -2.10891197072276
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -32.9542348920698
    point.Y = startPoint.Y - -3.44201723772157
    point.Z = startPoint.Z + 2.45666666666632
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    point.x = startPoint.x - -31.9062083245021
    point.Y = startPoint.Y - 0.844605732886673
    point.Z = startPoint.Z + 2.45666666666631
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -27.8145977524912
    point.Y = startPoint.Y + 4.68605728214073
    point.Z = startPoint.Z + 2.45666666666631
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    CadInputQueue.SendCommand "NEWFILE P:\Active Projects\PGE\Substation\6446 BELL\2000 Substation\2300 Engineering\2310 Electrical\2311 Drawings\Indoor\bell7313b0.dgn"

    CadInputQueue.SendKeyin "task sendtaskchangedasync"

    CadInputQueue.SendKeyin "task sendtaskchangedasync ""\Drawing"""

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    CadInputQueue.SendCommand "EDIT SINGLE DIALOG"

    point.x = startPoint.x - -29.4404212279644
    point.Y = startPoint.Y - -1.27945664417945
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine BELL-7313"

    point.x = startPoint.x - -31.7404212279644
    point.Y = startPoint.Y - -0.779456644179446
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - -30.0404212279644
    point.Y = startPoint.Y - -1.77945664417945
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine 21"

    point.x = startPoint.x - -30.0404212279644
    point.Y = startPoint.Y - -1.77945664417945
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - -29.7404212279644
    point.Y = startPoint.Y - -3.07945664417945
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine BELL SUBSTATION #3                                      "

    point.x = startPoint.x - -29.5404212279644
    point.Y = startPoint.Y + 3.67945664417945
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - -30.0404212279644
    point.Y = startPoint.Y + 3.57945664417945
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine 15KV OUTDOOR SWITCHGEAR                                      "

    point.x = startPoint.x - -30.0404212279644
    point.Y = startPoint.Y + 3.57945664417945
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - -30.3404212279644
    point.Y = startPoint.Y - -3.27945664417945
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine UNIT 4 - FDR. BKR R304 CONTROL SCHEM.                                      "

    point.x = startPoint.x - -30.3404212279644
    point.Y = startPoint.Y - -3.27945664417945
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - -31.2404212279644
    point.Y = startPoint.Y - -2.97945664417945
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendKeyin "dialog viewsettings popup"

    CadInputQueue.SendKeyin "MDL KEYIN BENTLEY.VIEWATTRIBUTESDIALOG,VAD VIEWATTRIBUTESDIALOG SETATTRIBUTE 0 DataFields False"

    CadInputQueue.SendKeyin "VBA RUN BUTTONS"

    CadInputQueue.SendCommand "PRINT EXECUTE"

    point.x = startPoint.x - -31.5756002218801
    point.Y = startPoint.Y + 6.28452769205179
    point.Z = startPoint.Z + 2.45666666666623
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "msDialogState.gridInfo.roundoffUnit", (ActiveModelReference.UORsPerMasterUnit * 0.05), "MGDSHOOK"

    CadInputQueue.SendCommand "ACTIVE UNITROUND"

    point.x = startPoint.x - -31.2800461097629
    point.Y = startPoint.Y + 4.80584821066881
    point.Z = startPoint.Z + 2.45666666666623
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    point.x = startPoint.x - -28.8031741484325
    point.Y = startPoint.Y - -2.92696091312891
    point.Z = startPoint.Z + 2.45666666666622
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -26.650079006342
    point.Y = startPoint.Y - -2.52282284325535
    point.Z = startPoint.Z + 2.45666666666623
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -27.4721049433915
    point.Y = startPoint.Y - -1.33987925814897
    point.Z = startPoint.Z + 2.45666666666622
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -27.439208485695
    point.Y = startPoint.Y - -1.66904616878726
    point.Z = startPoint.Z + 2.45666666666622
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -27.2683011078185
    point.Y = startPoint.Y - -1.92929375751067
    point.Z = startPoint.Z + 2.45666666666622
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -27.3094216799392
    point.Y = startPoint.Y - -2.45583223370748
    point.Z = startPoint.Z + 2.45666666666622
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendKeyin Chr$(27)

    CadInputQueue.SendCommand "MOVE ICON"

    point.x = startPoint.x - -27.1904212279644
    point.Y = startPoint.Y - -2.17945664417945
    point.Z = startPoint.Z + 2.45666666666622
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -27.8404212279644
    point.Y = startPoint.Y - -1.57945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    point.x = startPoint.x - -22.7476082103034
    point.Y = startPoint.Y - -0.687594762848514
    point.Z = startPoint.Z + 2.45666666666622
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -21.366213990625
    point.Y = startPoint.Y + 4.56109209994958
    point.Z = startPoint.Z + 2.45666666666622
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "MOVE ICON"

    point.x = startPoint.x - -21.4883140025202
    point.Y = startPoint.Y + 4.29160664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -21.4404212279644
    point.Y = startPoint.Y + 5.52945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    CadInputQueue.SendCommand "PRINT EXECUTE"

    point.x = startPoint.x - -31.3956977188522
    point.Y = startPoint.Y + 6.1559468675837
    point.Z = startPoint.Z + 2.45666666666623
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -31.3956977188522
    point.Y = startPoint.Y + 6.16880495003051
    point.Z = startPoint.Z + 2.45666666666623
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -31.3956977188522
    point.Y = startPoint.Y + 6.18166303247732
    point.Z = startPoint.Z + 2.45666666666623
    CadInputQueue.SendDataPoint point, 1

'   The following statement opens modal dialog "Open"

    CadInputQueue.SendCommand "DIALOG OPENFILE"

    CadInputQueue.SendKeyin "task sendtaskchangedasync"

    CadInputQueue.SendKeyin "task sendtaskchangedasync ""\Drawing"""

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    point.x = startPoint.x - 7.39583659943892E-02
    point.Y = startPoint.Y - 8.62494185738905E-02
    point.Z = startPoint.Z - 6.10623E-15
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.363088088016569
    point.Y = startPoint.Y + 0.207879928787902
    point.Z = startPoint.Z - 6.10623E-15
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    point.x = startPoint.x - 7.87771946947586E-02
    point.Y = startPoint.Y - 0.115180174052099
    point.Z = startPoint.Z - 6.10623E-15
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.184791426102891
    point.Y = startPoint.Y + 0.101800492034469
    point.Z = startPoint.Z - 6.10623E-15
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "DELETE ELEMENT"

    point.x = startPoint.x - 0.160697282601043
    point.Y = startPoint.Y - 5.24968705159797E-02
    point.Z = startPoint.Z - 6.10623E-15
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.247436199207697
    point.Y = startPoint.Y + 9.21569068750659E-02
    point.Z = startPoint.Z - 6.10623E-15
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "DELETE ELEMENT"

    CadInputQueue.SendCommand "MDL SILENTLOAD USTNVBA MACROS"

'   The following statement opens modal dialog "Color Table"

    CadInputQueue.SendCommand "DIALOG COLOR"

    CadInputQueue.SendKeyin "level purge all"

    CadInputQueue.SendReset

'   The following statement opens modal dialog "Color Table"

    CadInputQueue.SendCommand "DIALOG COLOR"

    CadInputQueue.SendKeyin "level purge all"

    point.x = startPoint.x - 3.87185420612206
    point.Y = startPoint.Y - 0.87389677051904
    point.Z = startPoint.Z - 2.388367E-14
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 3.93067779865587
    point.Y = startPoint.Y - 0.638457679713302
    point.Z = startPoint.Z - 2.388367E-14
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

'   The following statement opens modal dialog "Level/Filter Import"

'   The following statement opens modal dialog "Import Levels"

    CadInputQueue.SendCommand "LEVELMANAGER LIBRARY IMPORT"

    point.x = startPoint.x - 3.61450098878665
    point.Y = startPoint.Y + 2.44358716836557
    point.Z = startPoint.Z - 3.001766E-14
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 3.22374164326869
    point.Y = startPoint.Y - 1.80535267351924
    point.Z = startPoint.Z - 3.001766E-14
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendKeyin Chr$(27)

    CadInputQueue.SendCommand "Change Attributes"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES USEACTIVE ON"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE LEVEL"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET LEVEL ""Vendor"""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE COLOR"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET COLOR ""0"""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE LINESTYLE"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET LINESTYLE ""Continuous"""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE WEIGHT"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET WEIGHT 0"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE TRANSPARENCY"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET TRANSPARENCY 0"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE PRIORITY"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET PRIORITY 0"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE ELEMENTCLASS"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET ELEMENTCLASS PRIMARY"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES DISABLE TEMPLATE"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES SET TEMPLATE """""

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES MAKECOPY OFF"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENTIREELEMENT OFF"

    SetCExpressionValue "tcb->msToolSettings.general.useFence", 0, "CHANGEATTRIBS"

    CadInputQueue.SendCommand "LOCK FENCE INSIDE"

    CadInputQueue.SendKeyin "CHANGE ATTRIBUTES ENABLE COLOR"

    point.x = startPoint.x + 6.19735445764391E-02
    point.Y = startPoint.Y + 1.48711586196725
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

'   The following statement opens modal dialog "Design File Settings"

    CadInputQueue.SendCommand "MDL SILENTLOAD DGNSET"

    CadInputQueue.SendCommand "FILEDESIGN"

    CadInputQueue.SendCommand "MDL SILENTUNLOAD DGNSET"

    point.x = startPoint.x + 0.135503035243701
    point.Y = startPoint.Y - 0.77530665124414
    point.Z = startPoint.Z - 2.445266E-14
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 3.52258912545257
    point.Y = startPoint.Y + 3.51042054857907
    point.Z = startPoint.Z - 2.445266E-14
    CadInputQueue.SendDataPoint point, 1

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

'   The following statement opens modal dialog "Design File Settings"

    CadInputQueue.SendCommand "MDL SILENTLOAD DGNSET"

    CadInputQueue.SendCommand "FILEDESIGN"

    CadInputQueue.SendCommand "MDL SILENTUNLOAD DGNSET"

    CadInputQueue.SendCommand "NEWFILE U:\New folder\bellSWGRTITLEBLOCK.dgn"

    CadInputQueue.SendKeyin "task sendtaskchangedasync"

    CadInputQueue.SendKeyin "task sendtaskchangedasync ""\Drawing"""

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    point.x = startPoint.x - -33.6289917232435
    point.Y = startPoint.Y - 4.6783607773943
    point.Z = startPoint.Z + 2.4566666666663
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -21.0670294407541
    point.Y = startPoint.Y + 12.8294198708256
    point.Z = startPoint.Z + 2.4566666666663
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "MDL LOAD CLIPBRD COPY"

    CadInputQueue.SendCommand "NEWFILE ""P:\Active Projects\PGE\Substation\6446 BELL\2000 Substation\2300 Engineering\2310 Electrical\2311 Drawings\Indoor\bell7313c0.dgn"",""~9308"""

    CadInputQueue.SendKeyin "task sendtaskchangedasync"

    CadInputQueue.SendKeyin "task sendtaskchangedasync ""\Drawing"""

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    CadInputQueue.SendCommand "COMPONENTVIEW COMPONENTSETOVERRIDE SUSPEND"

    CadInputQueue.SendCommand "MDL KEYIN CLIPBRD CLIPBOARD PASTE"

    point.x = startPoint.x - 37.4892706296046
    point.Y = startPoint.Y + 30.9404508829546
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    point.x = startPoint.x - 5.32835376563225
    point.Y = startPoint.Y + 69.9609485976437
    point.Z = startPoint.Z - -2.45666666666545
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 114.798833963212
    point.Y = startPoint.Y - 44.4094757385139
    point.Z = startPoint.Z - -2.45666666666545
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "SCALE ICON"

    CadInputQueue.SendCommand "ACTIVE XSCALE 0.3900"

    CadInputQueue.SendCommand "ACTIVE SCALE"

    point.x = startPoint.x + 105.834118461059
    point.Y = startPoint.Y + 28.249382075045
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "MOVE ICON"

    CadInputQueue.SendTentativePoint Point3dFromXYZ(98.2705269009586, 38.3111785306598, 2.5399999999987), 1

    point.x = startPoint.x + 95.1037765401523
    point.Y = startPoint.Y + 38.4830111570074
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendAdjustedDataPoint point, 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-19.8272901125179, 41.5777451127749, 0#), 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-19.7878628767806, 41.5777451127749, 0#), 1

    CadInputQueue.SendReset

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-19.6892947874374, 41.5382936268357, 2.5399999999987), 1

    point.x = startPoint.x - 21.1149206296046
    point.Y = startPoint.Y + 41.4404508829546
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendAdjustedDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "MOVE ICON"

    CadInputQueue.SendTentativePoint Point3dFromXYZ(98.0953691155843, 38.4380474080459, 2.53999999999871), 1

    point.x = startPoint.x + 95.1037765401523
    point.Y = startPoint.Y + 38.4830111570074
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendAdjustedDataPoint point, 1

    CadInputQueue.SendCommand "LOCK ASSOCIATION OFF"

    CadInputQueue.SendCommand "LOCK UNIT ON"

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-19.7282985463777, 41.402149975998, 0#), 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-19.7282985463777, 41.402149975998, 0#), 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-19.7282985463777, 41.402149975998, 0#), 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-19.7899036022172, 41.5254358695581, 0#), 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-19.7899036022172, 41.5254358695581, 0#), 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-19.7899036022172, 41.5254358695581, 0#), 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-19.7899036022172, 41.5254358695581, 0#), 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-19.6520087280777, 41.5632255108319, 0#), 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(-19.6520087280777, 41.5632255108319, 0#), 1

    point.x = startPoint.x - 22.6143206296046
    point.Y = startPoint.Y + 41.4404508829546
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendAdjustedDataPoint point, 1

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    point.x = startPoint.x - 66.0404762882776
    point.Y = startPoint.Y + 46.3408407936845
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 9.49248525517938
    point.Y = startPoint.Y + 6.9533530554312
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "MOVE ICON"

    point.x = startPoint.x - 54.6143206296046
    point.Y = startPoint.Y + 19.9404508829546
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendKeyin "xy=0,0"

    CadInputQueue.SendReset

    CadInputQueue.SendCommand "CHOOSE ELEMENT"

    SetCExpressionValue "powerSelectInfo.prefs.currMode", 4, "PSELECT"

    CadInputQueue.SendCommand "POWERSELECTOR DESELECT"

    CadInputQueue.SendKeyin "dialog viewsettings popup"

    CadInputQueue.SendKeyin "MDL KEYIN BENTLEY.VIEWATTRIBUTESDIALOG,VAD VIEWATTRIBUTESDIALOG SETATTRIBUTE 0 DataFields False"

    CadInputQueue.SendKeyin "dialog viewsettings popup"

    CadInputQueue.SendKeyin "MDL KEYIN BENTLEY.VIEWATTRIBUTESDIALOG,VAD VIEWATTRIBUTESDIALOG SETATTRIBUTE 0 DataFields True"

    CadInputQueue.SendCommand "EDIT SINGLE DIALOG"

    point.x = startPoint.x - -29.4404212279644
    point.Y = startPoint.Y - -2.07945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine BELL-XXXX"

    point.x = startPoint.x - -29.4404212279644
    point.Y = startPoint.Y - -2.07945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine BELL-7314"

    point.x = startPoint.x - -30.9404212279644
    point.Y = startPoint.Y - -1.57945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - -30.0404212279644
    point.Y = startPoint.Y - -2.27945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -30.0404212279644
    point.Y = startPoint.Y - -2.17945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine 2 "

    point.x = startPoint.x - -30.0404212279644
    point.Y = startPoint.Y - -2.17945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - -29.9404212279644
    point.Y = startPoint.Y - -2.47945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - -29.6404212279644
    point.Y = startPoint.Y - -3.27945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine BELL SUBSTATION #3                                      "

    point.x = startPoint.x - -29.6404212279644
    point.Y = startPoint.Y - -3.27945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - -30.0404212279644
    point.Y = startPoint.Y - -3.27945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine 15KV OUTDOOR SWITCHGEAR                                      "

    point.x = startPoint.x - -30.0404212279644
    point.Y = startPoint.Y - -3.27945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - -30.4404212279644
    point.Y = startPoint.Y - -3.17945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine UNIT 5 - MAIN BKR R302 CONTROL SCHEM.                                      "

    point.x = startPoint.x - -30.4404212279644
    point.Y = startPoint.Y - -3.17945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - -32.1404212279644
    point.Y = startPoint.Y - -2.37945664417945
    point.Z = startPoint.Z - 0.083333333333315
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendKeyin "dialog viewsettings popup"

    CadInputQueue.SendKeyin "MDL KEYIN BENTLEY.VIEWATTRIBUTESDIALOG,VAD VIEWATTRIBUTESDIALOG SETATTRIBUTE 0 DataFields False"

'   The following statement opens modal dialog "Design File Settings"

    CadInputQueue.SendCommand "MDL SILENTLOAD DGNSET"

    CadInputQueue.SendCommand "FILEDESIGN"

    CadInputQueue.SendCommand "MDL SILENTUNLOAD DGNSET"

'   The following statement opens modal dialog "Color Table"

    CadInputQueue.SendCommand "DIALOG COLOR"

    CadInputQueue.SendCommand "EXIT"

    CadInputQueue.SendCommand "PRINT EXIT PLOTDLG"

    RemoveModalDialogEventsHandler modalHandler
    CommandState.StartDefaultCommand
End Sub
Sub Macro6()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

    Dim modalHandler As New Macro6ModalHandler
    AddModalDialogEventsHandler modalHandler

'   The following statement opens modal dialog "Print Attributes"

'   Start a command
    CadInputQueue.SendCommand "PRINT ATTRIBDIALOG"

    RemoveModalDialogEventsHandler modalHandler
    CommandState.StartDefaultCommand
End Sub
Sub Macro7()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

    CommandState.StartDefaultCommand
End Sub
Sub Macro8()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

    Dim modalHandler As New Macro8ModalHandler
    AddModalDialogEventsHandler modalHandler

'   The following statement opens modal dialog "Print Attributes"

'   Start a command
    CadInputQueue.SendCommand "PRINT ATTRIBDIALOG"

    RemoveModalDialogEventsHandler modalHandler
    CommandState.StartDefaultCommand
End Sub
Sub Macro9()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Start a command
    CadInputQueue.SendCommand "PLACE FENCE ICON"

'   Send a tentative point
'   Coordinates are in master units
    CadInputQueue.SendTentativePoint Point3dFromXYZ(0.098639241090924, 14.4992497162524, 1.4111111111138), 1

'   Coordinates are in master units
    startPoint.x = 0#
    startPoint.Y = 14.6666666666666
    startPoint.Z = 0#

'   Send a data point to the current command
    point.x = startPoint.x
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendAdjustedDataPoint point, 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(22.6630469909838, -4.14684863753751E-03, 1.41111111111402), 1

    point.x = startPoint.x + 22.6666666666667
    point.Y = startPoint.Y - 14.6666666666666
    point.Z = startPoint.Z
    CadInputQueue.SendAdjustedDataPoint point, 1

    point.x = startPoint.x + 23.1448617967697
    point.Y = startPoint.Y - 6.5622857142857
    point.Z = startPoint.Z + 1.4111111111123
    CadInputQueue.SendDataPoint point, 5

    CadInputQueue.SendCommand "FIT VIEW EXTENDED 5"

    CadInputQueue.SendCommand "WINDOW AREA EXTENDED 1"

    point.x = startPoint.x + 21.7245776434224
    point.Y = startPoint.Y - 8.34410349586891
    point.Z = startPoint.Z + 1.41111111111421
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 21.6370483654411
    point.Y = startPoint.Y - 14.7128887948723
    point.Z = startPoint.Z + 1.41111111111426
    CadInputQueue.SendDataPoint point, 1

    CommandState.StartDefaultCommand
End Sub

Sub Macro10()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Start a command
    CadInputQueue.SendCommand "PLACE FENCE ICON"

'   Send a tentative point
'   Coordinates are in master units
    CadInputQueue.SendTentativePoint Point3dFromXYZ(0.365598749417673, 43.5815800805484, 4.23333333333436), 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(0.102272749327261, 44.0333269768638, 4.23333333333436), 1

'   Coordinates are in master units
    startPoint.x = 0#
    startPoint.Y = 43.9999999999999
    startPoint.Z = 0#

'   Send a data point to the current command
    point.x = startPoint.x
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendAdjustedDataPoint point, 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(67.9812786286272, 3.29940720802703E-02, 4.23333333333444), 1

    point.x = startPoint.x + 68#
    point.Y = startPoint.Y - 43.9999999999999
    point.Z = startPoint.Z
    CadInputQueue.SendAdjustedDataPoint point, 1

    point.x = startPoint.x + 8.83333333333333
    point.Y = startPoint.Y + 56.1666666666667
    point.Z = startPoint.Z + 4.23333333333333
    CadInputQueue.SendDataPoint point, 5

    CadInputQueue.SendCommand "FIT VIEW EXTENDED 5"

    CadInputQueue.SendCommand "WINDOW AREA EXTENDED 1"

    point.x = startPoint.x + 65.6749694500775
    point.Y = startPoint.Y - 24.7501961741185
    point.Z = startPoint.Z + 4.23333333333444
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 65.7208898602495
    point.Y = startPoint.Y - 44.117079717327
    point.Z = startPoint.Z + 4.23333333333444
    CadInputQueue.SendDataPoint point, 1

    CommandState.StartDefaultCommand
End Sub
Sub Macro11()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Coordinates are in master units
    startPoint.x = -2.38719521710856
    startPoint.Y = 5.40902255639098
    startPoint.Z = 0#

'   Send a data point to the current command
    point.x = startPoint.x
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 5

'   Send a keyin that can be a command string
    CadInputQueue.SendKeyin "dialog viewsettings popup"

    CadInputQueue.SendKeyin "MDL KEYIN BENTLEY.VIEWATTRIBUTESDIALOG,VAD VIEWATTRIBUTESDIALOG SETATTRIBUTE 4 DataFields False"

    CommandState.StartDefaultCommand
End Sub
Sub Macro12()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Start a command
    CadInputQueue.SendCommand "PLACE FENCE ICON"

'   Send a tentative point
'   Coordinates are in master units
    CadInputQueue.SendTentativePoint Point3dFromXYZ(0.493826490298015, 30.0246028833885, 0#), 1

'   Coordinates are in master units
    startPoint.x = 0#
    startPoint.Y = 30#
    startPoint.Z = 0#

'   Send a data point to the current command
    point.x = startPoint.x
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendAdjustedDataPoint point, 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(41.97131768619, 1.17293206833481E-02, 0#), 1

    point.x = startPoint.x + 42.0000000000001
    point.Y = startPoint.Y - 30#
    point.Z = startPoint.Z
    CadInputQueue.SendAdjustedDataPoint point, 1

    point.x = startPoint.x + 1.79290742839691
    point.Y = startPoint.Y + 3.46071529917275
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 5

    CadInputQueue.SendCommand "PRINT MAXIMIZE"

    CadInputQueue.SendCommand "WINDOW AREA EXTENDED 1"

    point.x = startPoint.x + 40.9646372074583
    point.Y = startPoint.Y - 30.3439460111004
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 40.801902248909
    point.Y = startPoint.Y - 20.3922465283965
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CommandState.StartDefaultCommand
End Sub
Sub Macro13()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Start a command
    CadInputQueue.SendCommand "PLACE FENCE ICON"

'   Send a tentative point
'   Coordinates are in master units
    CadInputQueue.SendTentativePoint Point3dFromXYZ(-7.68330441137402, 29.5718401832303, 0#), 1

'   Coordinates are in master units
    startPoint.x = -8.00000000000005
    startPoint.Y = 30#
    startPoint.Z = 0#

'   Send a data point to the current command
    point.x = startPoint.x
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendAdjustedDataPoint point, 1

    CadInputQueue.SendTentativePoint Point3dFromXYZ(33.9605377145485, 2.06433146363044E-02, 0#), 1

    point.x = startPoint.x + 42#
    point.Y = startPoint.Y - 30#
    point.Z = startPoint.Z
    CadInputQueue.SendAdjustedDataPoint point, 1

    point.x = startPoint.x - 4.59999999999996
    point.Y = startPoint.Y - 14.05
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 5

    CadInputQueue.SendCommand "FIT VIEW EXTENDED 5"

    CadInputQueue.SendCommand "WINDOW AREA EXTENDED 1"

    point.x = startPoint.x + 40.7100968867561
    point.Y = startPoint.Y - 20.425106943707
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 40.9609209462061
    point.Y = startPoint.Y - 30.0283192740974
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CommandState.StartDefaultCommand
End Sub
Sub Macro14()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Coordinates are in master units
    startPoint.x = -6.56545142075177
    startPoint.Y = 54.3998934213841
    startPoint.Z = 0#

'   Send a data point to the current command
    point.x = startPoint.x
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 5.95906729282748
    point.Y = startPoint.Y + 4.23707626929689
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

'   Start a command
    CadInputQueue.SendCommand "PLACE FENCE ICON"

    point.x = startPoint.x - 0.439089168945145
    point.Y = startPoint.Y + 8.69041997317247
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 6.20997538936757
    point.Y = startPoint.Y + 13.4610539948993
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendCommand "PLACE FENCE ICON"

    point.x = startPoint.x + 0.354799730263835
    point.Y = startPoint.Y - 3.12043012749257
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 5.91790268323876
    point.Y = startPoint.Y + 4.72600740824241
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CommandState.StartDefaultCommand
End Sub
Sub Macro15()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Coordinates are in master units
    startPoint.x = -7.0696445597244
    startPoint.Y = 57.3839304543931
    startPoint.Z = 0#

'   Send a data point to the current command
    point.x = startPoint.x
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 6.42324727142667
    point.Y = startPoint.Y + 6.62867043018892
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

'   Start a command
    CadInputQueue.SendCommand "ORDER ELEMENT FRONT"

    CommandState.StartDefaultCommand
End Sub
Sub Macro16()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Coordinates are in master units
    startPoint.x = 33.1527004748686
    startPoint.Y = 0.754570801973472
    startPoint.Z = 0#

'   Send a data point to the current command
    point.x = startPoint.x
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

'   Send a message string to an application
'   Content is defined by the application
    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 16"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 setColor 3"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 16"

    point.x = startPoint.x + 2.02202552933635
    point.Y = startPoint.Y + 0.511180800000016
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.474140430326564
    point.Y = startPoint.Y + 11.9115087470084
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 27"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 setColor 3"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 27"

    point.x = startPoint.x + 0.491130115385928
    point.Y = startPoint.Y + 11.7356298362372
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.531604437698213
    point.Y = startPoint.Y + 12.735398608026
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 83"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 setColor 3"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 83"

    point.x = startPoint.x + 1.09579141498126
    point.Y = startPoint.Y + 13.787661622426
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.486300933930082
    point.Y = startPoint.Y + 18.8328585963654
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 18"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 setColor 3"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 18"

    point.x = startPoint.x + 0.966672293459965
    point.Y = startPoint.Y + 19.4093287712902
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x - 0.483975647140291
    point.Y = startPoint.Y + 19.2116423764102
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 18"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 setColor 3"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 18"

    point.x = startPoint.x - 0.465645794156544
    point.Y = startPoint.Y + 20.0158674528391
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 18"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 setColor 3"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 18"

    point.x = startPoint.x - 0.427927708092867
    point.Y = startPoint.Y + 19.6215636852625
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 18"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 setColor 3"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 18"

    point.x = startPoint.x - 0.435867267566309
    point.Y = startPoint.Y + 20.4314206033249
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 18"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 setColor 3"

    CadInputQueue.SendMessageToApplication "WORDPROC", "1 selection 13 18"

    point.x = startPoint.x - 0.274746316731445
    point.Y = startPoint.Y + 20.4222726665818
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CommandState.StartDefaultCommand
End Sub
Sub Macro17()
    Dim startPoint As Point3d
    Dim point As Point3d, point2 As Point3d
    Dim lngTemp As Long

'   Coordinates are in master units
    startPoint.x = 32.5440839946826
    startPoint.Y = 13.6218545103832
    startPoint.Z = 0#

'   Send a data point to the current command
    point.x = startPoint.x
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

'   Send a message string to an application
'   Content is defined by the application
    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine REVISED PRIOR TO CONSTRUCTION, AWO 1000001215.                                                                    "

    point.x = startPoint.x
    point.Y = startPoint.Y
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 0.857899547846586
    point.Y = startPoint.Y - 0.754601892977787
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1, 2

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - 8.19649249535317E-02
    point.Y = startPoint.Y + 6.9992059638519
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    point.x = startPoint.x + 9.84945181524211E-02
    point.Y = startPoint.Y + 5.95595517648597
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine PES   "

    point.x = startPoint.x + 9.84945181524211E-02
    point.Y = startPoint.Y + 5.95595517648597
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x + 8.10086674956736E-02
    point.Y = startPoint.Y + 6.33216249704301
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine DDB   "

    point.x = startPoint.x + 8.10086674956736E-02
    point.Y = startPoint.Y + 6.33216249704301
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x + 5.91513541747446E-02
    point.Y = startPoint.Y + 6.75648935860153
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine RCL   "

    point.x = startPoint.x + 5.91513541747446E-02
    point.Y = startPoint.Y + 6.74774035114672
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x + 7.22657421673034E-02
    point.Y = startPoint.Y + 7.1501946940682
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine DDB   "

    point.x = startPoint.x + 7.22657421673034E-02
    point.Y = startPoint.Y + 7.1501946940682
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x + 8.53801301598622E-02
    point.Y = startPoint.Y + 7.52202751089783
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine REJ   "

    point.x = startPoint.x + 8.53801301598622E-02
    point.Y = startPoint.Y + 7.52202751089783
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CadInputQueue.SendMessageToApplication "TEXTEDIT", "FirstLine "

    point.x = startPoint.x - 0.771426552020685
    point.Y = startPoint.Y + 7.2114377462519
    point.Z = startPoint.Z
    CadInputQueue.SendDataPoint point, 1

    CommandState.StartDefaultCommand
End Sub
