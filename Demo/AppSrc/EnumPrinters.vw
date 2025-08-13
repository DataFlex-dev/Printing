Use Windows.pkg
Use DFClient.pkg
Use cDefaultPrinter.pkg
Use cCJGridColumnRowIndicator.pkg
Use cCJGridColumn.pkg
Use cSplitterContainer.pkg

Use cPrintersHandler.pkg
Use cPrinterDriversHandler.pkg
Use cPrinterJobsHandler.pkg
Use cPrinterFormsHandler.pkg
Use cPrinterPortsHandler.pkg

Use Classes\cDateTimeHandler.pkg

Use PrinterInfoDialog.dg
Use DriverInfoDialog.dg
Use cCJGrid.pkg

Activate_View Activate_oEnumPrintersView for oEnumPrintersView
Object oEnumPrintersView is a dbView
    Set Label To "Printer Information"
    Set Border_Style to Border_Thick
    Set Size to 265 407
    Set Location to 2 2
    Set Icon to "Printing.Ico"
    Set pbAutoActivate to True

    Object oPrintersHandler is a cPrintersHandler
    End_Object

    Object oPrinterDriversHandler is a cPrinterDriversHandler
    End_Object

    Object oPrinterJobsHandler is a cPrinterJobsHandler
    End_Object

    Object oPrinterFormsHandler is a cPrinterFormsHandler
    End_Object

    Object oPrinterPortsHandler is a cPrinterPortsHandler
    End_Object

    Object oInfoTabDialog is a TabDialog
        Set Size to 254 399
        Set Location to 5 5
        Set peAnchors to anAll

        Object oPrintersTabPage is a TabPage
            Set Label to "Printers"

            Object oSplitterContainer1 is a cSplitterContainer
                Set pbSplitVertical to False
                Set piSplitterLocation to 114

                Object oPrintersSection is a cSplitterContainerChild
                    // Shows the enumerated printernames
                    Object oPrintersList is a cCJGrid
                        Set Size to 105 300
                        Set Location to 5 5
                        Set peAnchors to anAll
                        Set pbAllowEdit to False
                        Set pbAllowDeleteRow to False
                        Set pbAllowAppendRow to False
                        Set pbAutoAppend to False
                        Set pbAllowInsertRow to False
                        Set pbAutoSave to False
                        Set pbEditOnTyping to False
                        Set piAlternateRowBackgroundColor to clBtnFace
                        Set pbUseAlternateRowBackgroundColor to True
                        Set psNoItemsText to "No Printers Enumerated"

                        Object oRowIndicator is a cCJGridColumnRowIndicator
                            Set piWidth to 23
                        End_Object

                        Object oPrinterNameColumn is a cCJGridColumn
                            Set piWidth to 421
                            Set psCaption to "PrinterName"
                        End_Object

                        Object oDefaultPrinterColumn is a cCJGridColumn
                            Set piWidth to 77
                            Set pbCheckbox to True
                            Set psCaption to "Default"
                        End_Object

                        // Lists the printernames stored in the oWindowsPrinters object
                        Procedure FillPrintersList
                            tPrinterInfo[] PrintersInfo
                            Integer iPrinters iPrinter iPrinterNameColumnID iDefaultPrinterColumnID
                            String sDefaultPrinterName
                            Handle hoDefaultPrinter
                            tDataSourceRow[] PrintersGridData

                            Get Create (RefClass (cDefaultPrinter)) to hoDefaultPrinter
                            Get DefaultPrinter of hoDefaultPrinter to sDefaultPrinterName
                            Send Destroy of hoDefaultPrinter

                            Get pPrintersInfo of oPrintersHandler to PrintersInfo

                            Move (SizeOfArray (PrintersInfo)) to iPrinters
                            If (iPrinters > 0) Begin
                                Decrement iPrinters

                                Get piColumnId of oPrinterNameColumn to iPrinterNameColumnID
                                Get piColumnId of oDefaultPrinterColumn to iDefaultPrinterColumnID

                                For iPrinter from 0 to iPrinters
                                    Move PrintersInfo[iPrinter].sPrinterName to PrintersGridData[iPrinter].sValue[iPrinterNameColumnID]
                                    Move (PrintersInfo[iPrinter].sPrinterName = sDefaultPrinterName) to PrintersGridData[iPrinter].sValue[iDefaultPrinterColumnID]
                                Loop
                            End

                            Send InitializeData PrintersGridData
                        End_Procedure

                        Function SelectedPrinterName Returns String
                            Handle hoDataSource
                            Integer iRow
                            String sPrinterName

                            Get phoDataSource to hoDataSource
                            Get SelectedRow of hoDataSource to iRow
                            Get RowValue of oPrinterNameColumn iRow to sPrinterName

                            Function_Return sPrinterName
                        End_Function

                        Procedure OnRowDoubleClick Integer iRow Integer iCol
                            Forward Send OnRowDoubleClick iRow iCol

                            Send ShowInfo of oPrinterInfoDialog oPrintersHandler iRow
                        End_Procedure
                    End_Object

                    // Makes it possible to scan the windows printers on this machine
                    Object oScanButton is a Button
                        Set Size to 14 80
                        Set Location to 21 309
                        Set Label to "Scan printers"
                        Set peAnchors to anTopRight

                        // Starts the printer scan method
                        Procedure OnClick
                            // ToDo: For now we hardcode the enum operation. Should be listening to the flags grid and use other structs
                            Send EnumPrintersOnLocalMachine of oPrintersHandler True
                            Send FillPrintersList of oPrintersList
                            Send EnableDisableButton of oSetDefaultPrinterButton
                        End_Procedure
                    End_Object

                    Object oSetDefaultPrinterButton is a Button
                        Set Size to 14 80
                        Set Location to 5 309
                        Set Label to "Set Default Printer"
                        Set peAnchors to anTopRight

                        // Starts the printer scan method
                        Procedure OnClick
                            Handle hoDefaultPrinter
                            Boolean bChanged
                            String sDefaultPrinterName

                            Get Create (RefClass (cDefaultPrinter)) to hoDefaultPrinter

                            Get Value of oPrintersList to sDefaultPrinterName
                            Get ChangeDefaultPrinter of hoDefaultPrinter sDefaultPrinterName to bChanged
                            If (not (bChanged)) Begin
                                Send Stop_Box "Change to default printer failed"
                            End
                            Else Begin
                                Send KeyAction of oScanButton
                            End

                            Send Destroy of hoDefaultPrinter
                        End_Procedure

                        // Code to enable or disable the button based on the presence of printers in
                        // the list of printers
                        Procedure Activating
                            Forward Send Activating
                            Send EnableDisableButton
                        End_Procedure

                        Procedure EnableDisableButton
                            Handle hoDataSource
                            Integer iPrinters

                            Get phoDataSource of oPrintersList to hoDataSource
                            Get RowCount of hoDataSource to iPrinters

                            Set Enabled_State to (iPrinters > 0)
                        End_Procedure
                    End_Object

                    Object oFlagsGrid is a cCJGrid
                        Set Size to 67 77
                        Set Location to 40 311
                        Set pbAllowAppendRow to False
                        Set pbAllowDeleteRow to False
                        Set pbAllowInsertRow to False
                        Set pbAutoAppend to False
                        Set peAnchors to anTopRight

                        Object oPrinterFlagColumn is a cCJGridColumn
                            Set pbVisible to False
                            Set pbShowInFieldChooser to False
                            Set pbEditable to False
                        End_Object

                        Object oFlagSelectedColumn is a cCJGridColumn
                            Set piWidth to 27
                            Set pbCheckbox to True
                        End_Object

                        Object oFlagNameColumn is a cCJGridColumn
                            Set piWidth to 101
                            Set psCaption to "Flag"

                            Function OnGetTooltip Integer iRow String sValue String sText Returns String
                                String sRetVal

                                Get RowValue of oFlagToolTipColumn iRow to sRetVal

                                Function_Return sRetVal
                            End_Function
                        End_Object

                        Object oFlagToolTipColumn is a cCJGridColumn
                            Set pbVisible to False
                            Set pbShowInFieldChooser to False
                            Set pbEditable to False
                        End_Object

                        Procedure AddFlagInfo tDataSourceRow[] ByRef PrinterFlagRows Boolean bSelected Integer eFlag String sName String sDescription
                            Integer iFlagColumnId iFlagSelectedColumnId iFlagNameColumnId iFlagToolTipColumnId iRow

                            Get piColumnId of oPrinterFlagColumn to iFlagColumnId
                            Get piColumnId of oFlagSelectedColumn to iFlagSelectedColumnId
                            Get piColumnId of oFlagNameColumn to iFlagNameColumnId
                            Get piColumnId of oFlagToolTipColumn to iFlagToolTipColumnId

                            Move (SizeOfArray (PrinterFlagRows)) to iRow
                            Move eFlag to PrinterFlagRows[iRow].sValue[iFlagColumnId]
                            Move bSelected to PrinterFlagRows[iRow].sValue[iFlagSelectedColumnId]
                            Move sName to PrinterFlagRows[iRow].sValue[iFlagNameColumnId]
                            Move sDescription to PrinterFlagRows[iRow].sValue[iFlagToolTipColumnId]
                        End_Procedure

                        Procedure Activating
                            tDataSourceRow[] PrinterFlags

                            Forward Send Activating

                            Send AddFlagInfo (&PrinterFlags) False PRINTER_ENUM_LOCAL "Local" "If the PRINTER_ENUM_NAME flag is not also passed, the Function ignores the Name parameter, and enumerates the locally installed printers. If PRINTER_ENUM_NAME is also passed, the Function enumerates the local printers on Name."
                            Send AddFlagInfo (&PrinterFlags) False PRINTER_ENUM_NAME "Name" "The Function enumerates the printer identified by Name. This can be a server, a domain, or a Print provider. If Name is NULL, the Function enumerates available Print providers."
                            Send AddFlagInfo (&PrinterFlags) False PRINTER_ENUM_SHARED "Shared" "The Function enumerates printers that have the shared attribute. Cannot be used in isolation; Use an OR operation to combine with another PRINTER_ENUM type."
                            Send AddFlagInfo (&PrinterFlags) True PRINTER_ENUM_CONNECTIONS "Connections" "The Function enumerates the list of printers to which the user has made previous connections."
                            Send AddFlagInfo (&PrinterFlags) False PRINTER_ENUM_NETWORK "Network" "The Function enumerates network printers in the computer's domain. This value is valid only if Level is 1."
                            Send AddFlagInfo (&PrinterFlags) False PRINTER_ENUM_REMOTE "Remote" "The Function enumerates network printers and Print servers in the computer's domain. This value is valid only if Level is 1."
                            Send AddFlagInfo (&PrinterFlags) False PRINTER_ENUM_CATEGORY_3D "3D" "The Function enumerates only 3D printers."
                            Send AddFlagInfo (&PrinterFlags) False PRINTER_ENUM_CATEGORY_ALL "All" "The Function enumerates All Print devices, including 3D printers."

                            Send InitializeData PrinterFlags
                        End_Procedure
                    End_Object
                End_Object

                Object oPrintersDetailSection is a cSplitterContainerChild
                    Object oPrinterDetailsDialog is a TabDialog
                        Set Size to 113 384
                        Set Location to 5 5
                        Set peAnchors to anAll

                        Object oJobsTabPage is a TabPage
                            Set Label to "Jobs"

                            Object oJobsList is a cCJGrid
                                Set Size to 85 300
                                Set Location to 5 5
                                Set peAnchors to anAll
                                Set pbAllowEdit to False
                                Set pbAllowDeleteRow to False
                                Set pbAllowAppendRow to False
                                Set pbAutoAppend to False
                                Set pbAllowInsertRow to False
                                Set pbAutoSave to False
                                Set pbEditOnTyping to False
                                Set piAlternateRowBackgroundColor to clBtnFace
                                Set pbUseAlternateRowBackgroundColor to True
                                Set psNoItemsText to "No Jobs Enumerated"

                                Object oRowIndicator is a cCJGridColumnRowIndicator
                                    Set piWidth to 42
                                End_Object

                                Object oJobIDColumn is a cCJGridColumn
                                    Set piWidth to 67
                                    Set psCaption to "Job ID"
                                End_Object

                                Object oJobStatusColumn is a cCJGridColumn
                                    Set piWidth to 67
                                    Set psCaption to "Status"
                                End_Object

                                Object oJobStatusTextColumn is a cCJGridColumn
                                    Set piWidth to 343
                                    Set psCaption to "Status Description"
                                End_Object

                                // This method lists the jobs for the currently selected printer
                                Procedure FillJobsList
                                    Integer eType iJobIDColumnID iJobStatusColumnID iJobStatusTextColumnID iJob iJobs
                                    Integer iElements iElement iRow
                                    String sPrinterName
                                    tPrinterJobInfo[] JobsInfo
                                    tDataSourceRow[] JobsGridData

                                    Get SelectedPrinterName of oPrintersList to sPrinterName
                                    Get InfoType of oJobsTypeComboForm to eType
                                    Send EnumJobs of oPrinterJobsHandler sPrinterName eType
                                    Get pJobsInfo of oPrinterJobsHandler to JobsInfo
                                    Move (SizeOfArray (JobsInfo)) to iJobs
                                    If (iJobs > 0) Begin
                                        Get piColumnId of oJobIDColumn to iJobIDColumnID
                                        Get piColumnId of oJobStatusColumn to iJobStatusColumnID
                                        Get piColumnId of oJobStatusTextColumn to iJobStatusTextColumnID

                                        Decrement iJobs
                                        For iJob from 0 to iJobs
                                            Move JobsInfo[iJob].uiJobId to JobsGridData[iRow].sValue[iJobIDColumnID]
                                            Move JobsInfo[iJob].uiStatus to JobsGridData[iRow].sValue[iJobStatusColumnID]
                                            Move (SizeOfArray (JobsInfo[iJob].sStatus)) to iElements
                                            If (iElements > 0) Begin
                                                Decrement iElements
                                                For iElement from 0 to iElements
                                                    Move JobsInfo[iJob].sStatus[iElement] to JobsGridData[iRow].sValue[iJobStatusTextColumnID]
                                                    Increment iRow
                                                Loop
                                            End
                                            Else Begin
                                                Increment iRow
                                            End
                                        Loop
                                    End

                                    Send InitializeData JobsGridData
                                End_Procedure

                                Procedure OnIdle
                                    String sPrinterName
                                    Boolean bEnabled

                                    Forward Send OnIdle

                                    Get SelectedPrinterName of oPrintersList to sPrinterName
                                    Move (sPrinterName <> "") to bEnabled
                                    Set Enabled_State of oScanButton to bEnabled
                                    Set Enabled_State of oJobsTypeComboForm to bEnabled
                                    Set Enabled_State of oJobsList to bEnabled
                                End_Procedure
                            End_Object

                            Object oJobsTypeComboForm is a ComboForm
                                Set Size to 12 65
                                Set Location to 5 310
                                Set Entry_State to False
                                Set Combo_Sort_State to False
                                Set peAnchors to anTopRight

                                Procedure Combo_Fill_List
                                    Send Combo_Add_Item "Type 1"
                                    Send Combo_Add_Item "Type 2"
                                End_Procedure

                                Function InfoType Returns Integer
                                    String sValue
                                    Integer eValue

                                    Get Value to sValue
                                    Get Combo_Item_Matching sValue to eValue

                                    Function_Return (eValue + 1)
                                End_Function
                            End_Object

                            Object oScanButton is a Button
                                Set Size to 14 65
                                Set Location to 20 310
                                Set Label to "Scan Jobs"
                                Set peAnchors to anTopRight

                                // Starts the Job scan method
                                Procedure OnClick
                                    Send FillJobsList of oJobsList
                                End_Procedure
                            End_Object
                        End_Object

                        Object oFormsTabPage is a TabPage
                            Set Label to "Forms"

                            Object oFormsList is a cCJGrid
                                Set Size to 85 300
                                Set Location to 5 5
                                Set peAnchors to anAll
                                Set pbAllowEdit to False
                                Set pbAllowDeleteRow to False
                                Set pbAllowAppendRow to False
                                Set pbAutoAppend to False
                                Set pbAllowInsertRow to False
                                Set pbAutoSave to False
                                Set pbEditOnTyping to False
                                Set piAlternateRowBackgroundColor to clBtnFace
                                Set pbUseAlternateRowBackgroundColor to True
                                Set psNoItemsText to "No Forms Enumerated"

                                Object oRowIndicator is a cCJGridColumnRowIndicator
                                    Set piWidth to 40
                                End_Object

                                Object oFormNameColumn is a cCJGridColumn
                                    Set piWidth to 150
                                    Set psCaption to "Form Name"
                                End_Object

                                Object oFormSizeColumn is a cCJGridColumn
                                    Set piWidth to 150
                                    Set psCaption to "Size"
                                End_Object

                                Object oFormImageableAreaColumn is a cCJGridColumn
                                    Set piWidth to 150
                                    Set psCaption to "Imageable"
                                End_Object

                                Object oFormMUIDllColumn is a cCJGridColumn
                                    Set piWidth to 150
                                    Set psCaption to "MUI DLL"
                                    Set pbVisible to False
                                End_Object

                                Object oFormDisplayNameColumn is a cCJGridColumn
                                    Set piWidth to 150
                                    Set psCaption to "DisplayName"
                                    Set pbVisible to False
                                End_Object

                                Object oFormLanguageIdColumn is a cCJGridColumn
                                    Set piWidth to 150
                                    Set psCaption to "LanguageID"
                                    Set pbVisible to False
                                End_Object

                                // This method lists the forms for the currently selected printer
                                Procedure FillFormsList
                                    Integer iFormNameColumnId iFormSizeColumnId iFormImageableAreaColumnId
                                    Integer iFormMUIDllColumnId iFormDisplayNameColumnId iFormLanguageIdColumnId
                                    Integer eFlags eInfoType iForm iForms
                                    String sPrinterName
                                    tPrinterFormInfo[] FormsInfo
                                    tDataSourceRow[] FormsGridData

                                    Get SelectedPrinterName of oPrintersList to sPrinterName
                                    Get FilterValue of oFormsFilterComboForm to eFlags
                                    Get InfoType of oFormsTypeComboForm to eInfoType
                                    Send EnumForms of oPrinterFormsHandler sPrinterName eInfoType eFlags
                                    Get pFormsInfo of oPrinterFormsHandler to FormsInfo
                                    Move (SizeOfArray (FormsInfo)) to iForms
                                    If (iForms > 0) Begin
                                        Get piColumnId of oFormNameColumn to iFormNameColumnId
                                        Get piColumnId of oFormSizeColumn to iFormSizeColumnId
                                        Get piColumnId of oFormImageableAreaColumn to iFormImageableAreaColumnId
                                        Get piColumnId of oFormMUIDllColumn to iFormMUIDllColumnId
                                        Get piColumnId of oFormDisplayNameColumn to iFormDisplayNameColumnId
                                        Get piColumnId of oFormLanguageIdColumn to iFormLanguageIdColumnId

                                        Decrement iForms
                                        For iForm from 0 to iForms
                                            Move FormsInfo[iForm].sName to FormsGridData[iForm].sValue[iFormNameColumnID]
                                            Move (SFormat ("Width: %1mm, Height: %2mm", FormsInfo[iForm].Size.cx / 1000.0, FormsInfo[iForm].Size.cy / 1000.0)) to FormsGridData[iForm].sValue[iFormSizeColumnId]
                                            Move (SFormat ("Width: %1mm, Height: %2mm", FormsInfo[iForm].ImageableArea.right - FormsInfo[iForm].ImageableArea.left / 1000.0, FormsInfo[iForm].ImageableArea.bottom - FormsInfo[iForm].ImageableArea.top / 1000.0)) to FormsGridData[iForm].sValue[iFormImageableAreaColumnId]
                                            If (eInfoType = C_ENUMFORMS_TYPE_2) Begin
                                                Move FormsInfo[iForm].sMuiDll to FormsGridData[iForm].sValue[iFormMUIDllColumnId]
                                                Move FormsInfo[iForm].sDisplayName to FormsGridData[iForm].sValue[iFormDisplayNameColumnId]
                                                Move FormsInfo[iForm].uiLangId to FormsGridData[iForm].sValue[iFormLanguageIdColumnId]
                                            End
                                        Loop
                                    End

                                    Send InitializeData FormsGridData
                                End_Procedure

                                Procedure OnIdle
                                    String sPrinterName
                                    Boolean bEnabled

                                    Forward Send OnIdle

                                    Get SelectedPrinterName of oPrintersList to sPrinterName
                                    Move (sPrinterName <> "") to bEnabled
                                    Set Enabled_State of oScanButton to bEnabled
                                    Set Enabled_State of oFormsFilterComboForm to bEnabled
                                    Set Enabled_State of oFormsTypeComboForm to bEnabled
                                    Set Enabled_State of oFormsList to bEnabled
                                End_Procedure
                            End_Object

                            Object oFormsFilterComboForm is a ComboForm
                                Set Size to 12 65
                                Set Location to 5 309
                                Set Entry_State to False
                                Set Combo_Sort_State to False
                                Set peAnchors to anTopRight

                                Procedure Combo_Fill_List
                                    Send Combo_Add_Item "All"
                                    Send Combo_Add_Item "User"
                                    Send Combo_Add_Item "Builtin"
                                    Send Combo_Add_Item "Printer"
                                End_Procedure

                                Function FilterValue Returns Integer
                                    String sValue
                                    Integer eValue

                                    Get Value to sValue
                                    Get Combo_Item_Matching sValue to eValue

                                    Function_Return (eValue - 1)
                                End_Function
                            End_Object

                            Object oFormsTypeComboForm is a ComboForm
                                Set Size to 12 65
                                Set Location to 19 309
                                Set Entry_State to False
                                Set Combo_Sort_State to False
                                Set peAnchors to anTopRight

                                Procedure Combo_Fill_List
                                    Send Combo_Add_Item "Type 1"
                                    Send Combo_Add_Item "Type 2"
                                End_Procedure

                                Procedure OnChange
                                    Integer eType
                                    Boolean bShowColumns

                                    Get InfoType to eType
                                    Move (eType = C_ENUMFORMS_TYPE_2) to bShowColumns
                                    Set pbVisible of oFormMUIDllColumn to bShowColumns
                                    Set pbVisible of oFormDisplayNameColumn to bShowColumns
                                    Set pbVisible of oFormLanguageIdColumn to bShowColumns
                                End_Procedure

                                Function InfoType Returns Integer
                                    String sValue
                                    Integer eValue

                                    Get Value to sValue
                                    Get Combo_Item_Matching sValue to eValue

                                    Function_Return (eValue + 1)
                                End_Function
                            End_Object

                            Object oScanButton is a Button
                                Set Size to 14 65
                                Set Location to 33 309
                                Set Label to "Scan Forms"
                                Set peAnchors to anTopRight

                                // Starts the Job scan method
                                Procedure OnClick
                                    Send FillFormsList of oFormsList
                                End_Procedure
                            End_Object
                        End_Object

                        Object oPortsTabPage is a TabPage
                            Set Label to "Ports"

                            Object oPortsList is a cCJGrid
                                Set Size to 85 300
                                Set Location to 5 5
                                Set peAnchors to anAll
                                Set pbAllowEdit to False
                                Set pbAllowDeleteRow to False
                                Set pbAllowAppendRow to False
                                Set pbAutoAppend to False
                                Set pbAllowInsertRow to False
                                Set pbAutoSave to False
                                Set pbEditOnTyping to False
                                Set piAlternateRowBackgroundColor to clBtnFace
                                Set pbUseAlternateRowBackgroundColor to True
                                Set psNoItemsText to "No Ports Enumerated"

                                Object oRowIndicator is a cCJGridColumnRowIndicator
                                    Set piWidth to 40
                                End_Object

                                Object oPortNameColumn is a cCJGridColumn
                                    Set piWidth to 150
                                    Set psCaption to "Name"
                                End_Object

                                Object oMonitorNameColumn is a cCJGridColumn
                                    Set piWidth to 150
                                    Set psCaption to "Monitor"
                                    Set pbVisible to False
                                End_Object

                                Object oDescriptionColumn is a cCJGridColumn
                                    Set piWidth to 150
                                    Set psCaption to "Description"
                                    Set pbVisible to False
                                End_Object

                                Object oTypeColumn is a cCJGridColumn
                                    Set piWidth to 50
                                    Set psCaption to "Type"
                                    Set pbVisible to False
                                End_Object

                                // This method lists the ports for the currently selected printer
                                Procedure FillPortsList
                                    Integer eType iPort iPorts
                                    Integer iPortNameColumnID iMonitorNameColumnID iDescriptionColumnID iTypeColumnID
                                    Integer iElements iElement iRow
                                    Boolean bAllPorts
                                    String sPrinterName
                                    tPrinterPortInfo[] PortsInfo
                                    tDataSourceRow[] PortsGridData

                                    Get AllPortsState of oAllPortsComboForm to bAllPorts
                                    If (not (bAllPorts)) Begin
                                        Get SelectedPrinterName of oPrintersList to sPrinterName
                                    End
                                    Get InfoType of oPortsTypeComboForm to eType
                                    Send EnumPorts of oPrinterPortsHandler sPrinterName eType
                                    Get pPortsInfo of oPrinterPortsHandler to PortsInfo
                                    Move (SizeOfArray (PortsInfo)) to iPorts
                                    If (iPorts > 0) Begin
                                        Get piColumnId of oPortNameColumn to iPortNameColumnID
                                        Get piColumnId of oMonitorNameColumn to iMonitorNameColumnID
                                        Get piColumnId of oDescriptionColumn to iDescriptionColumnID
                                        Get piColumnId of oTypeColumn to iTypeColumnID

                                        Decrement iPorts
                                        For iPort from 0 to iPorts
                                            Move PortsInfo[iPort].uiPortType to PortsGridData[iRow].sValue[iTypeColumnID]
                                            Move PortsInfo[iPort].sPortName to PortsGridData[iRow].sValue[iPortNameColumnID]
                                            Move PortsInfo[iPort].sMonitorName to PortsGridData[iRow].sValue[iMonitorNameColumnID]
                                            Move PortsInfo[iPort].sDescription to PortsGridData[iRow].sValue[iDescriptionColumnID]
                                            Increment iRow
                                        Loop
                                    End

                                    Send InitializeData PortsGridData
                                End_Procedure

                                Procedure OnIdle
                                    String sPrinterName
                                    Boolean bAllPorts bEnabled

                                    Forward Send OnIdle

                                    Get SelectedPrinterName of oPrintersList to sPrinterName
                                    Get AllPortsState of oAllPortsComboForm to bAllPorts
                                    Move (sPrinterName <> "" or bAllPorts) to bEnabled
                                    Set Enabled_State of oScanButton to bEnabled
                                    Set Enabled_State of oPortsTypeComboForm to bEnabled
                                    Set Enabled_State of oPortsList to bEnabled
                                End_Procedure
                            End_Object

                            Object oPortsTypeComboForm is a ComboForm
                                Set Size to 12 65
                                Set Location to 5 310
                                Set Entry_State to False
                                Set Combo_Sort_State to False
                                Set peAnchors to anTopRight

                                Procedure Combo_Fill_List
                                    Send Combo_Add_Item "Type 1"
                                    Send Combo_Add_Item "Type 2"
                                End_Procedure

                                Procedure OnChange
                                    Integer eType
                                    Boolean bShowColumns

                                    Get InfoType to eType
                                    Move (eType = C_ENUMPORTS_TYPE_2) to bShowColumns
                                    Set pbVisible of oMonitorNameColumn to bShowColumns
                                    Set pbVisible of oDescriptionColumn to bShowColumns
                                    Set pbVisible of oTypeColumn to bShowColumns
                                End_Procedure

                                Function InfoType Returns Integer
                                    String sValue
                                    Integer eValue

                                    Get Value to sValue
                                    Get Combo_Item_Matching sValue to eValue

                                    Function_Return (eValue + 1)
                                End_Function
                            End_Object

                            Object oAllPortsComboForm is a ComboForm
                                Set Size to 12 65
                                Set Location to 19 310
                                Set Entry_State to False
                                Set Combo_Sort_State to False
                                Set peAnchors to anTopRight

                                Procedure Combo_Fill_List
                                    Send Combo_Add_Item "All Ports"
                                    Send Combo_Add_Item "Current Printer"
                                End_Procedure

                                Function AllPortsState Returns Boolean
                                    String sValue
                                    Integer eValue

                                    Get Value to sValue
                                    Get Combo_Item_Matching sValue to eValue

                                    Function_Return (eValue = 0)
                                End_Function
                            End_Object

                            Object oScanButton is a Button
                                Set Size to 14 65
                                Set Location to 33 310
                                Set Label to "Scan Ports"
                                Set peAnchors to anTopRight

                                // Starts the Job scan method
                                Procedure OnClick
                                    Send FillPortsList of oPortsList
                                End_Procedure
                            End_Object
                        End_Object
                    End_Object
                End_Object
            End_Object
        End_Object

        Object oDriversTabPage is a TabPage
            Set Label to "Drivers"

            Object oDriversList is a cCJGrid
                Set Size to 224 298
                Set Location to 5 5
                Set peAnchors to anAll
                Set pbAllowEdit to False
                Set pbAllowDeleteRow to False
                Set pbAllowAppendRow to False
                Set pbAutoAppend to False
                Set pbAllowInsertRow to False
                Set pbAutoSave to False
                Set pbEditOnTyping to False
                Set piAlternateRowBackgroundColor to clBtnFace
                Set pbUseAlternateRowBackgroundColor to True
                Set psNoItemsText to "No Drivers Enumerated"

                Object oRowIndicator is a cCJGridColumnRowIndicator
                    Set piWidth to 25
                End_Object

                Object oDriverNameColumn is a cCJGridColumn
                    Set piWidth to 495
                    Set psCaption to "DriverName"
                End_Object

                Procedure OnRowDoubleClick Integer iRow Integer iCol
                    Forward Send OnRowDoubleClick iRow iCol

                    Send ShowInfo of oDriverInfoDialog oPrinterDriversHandler iRow
                End_Procedure

                // This method lists the printerdrivers stored in the oWindowsPrinters object
                Procedure FillDriversList
                    tPrinterDriverInfo[] DriversInfo
                    Integer iDrivers iDriver iDriverNameColumnID
                    tDataSourceRow[] DriversGridData

                    Get pDriversInfo of oPrinterDriversHandler to DriversInfo

                    Move (SizeOfArray (DriversInfo)) to iDrivers
                    If (iDrivers > 0) Begin
                        Decrement iDrivers

                        Get piColumnId of oDriverNameColumn to iDriverNameColumnID

                        For iDriver from 0 to iDrivers
                            Move DriversInfo[iDriver].sName to DriversGridData[iDriver].sValue[iDriverNameColumnID]
                        Loop
                    End

                    Send InitializeData DriversGridData
                End_Procedure
            End_Object

            Object oDriversInfoComboForm is a ComboForm
                Set Size to 12 80
                Set Location to 5 310
                Set Entry_State to False
                Set Combo_Sort_State to False
                Set peAnchors to anTopRight

                Procedure Combo_Fill_List
                    Send Combo_Add_Item "Info 1"
                    Send Combo_Add_Item "Info 2"
                    Send Combo_Add_Item "Info 3"
                    Send Combo_Add_Item "Info 4"
                    Send Combo_Add_Item "Info 5"
                    Send Combo_Add_Item "Info 6"
                    Send Combo_Add_Item "Info 8"
                End_Procedure

                Function InfoType Returns Integer
                    String sValue
                    Integer eInfo

                    Get Value to sValue
                    Move (Right (sValue, 1)) to eInfo

                    Function_Return eInfo
                End_Function
            End_Object

            Object oScanButton Is A Button
                Set Size to 14 80
                Set Location to 19 310
                Set Label to "Scan drivers"
                Set peAnchors to anTopRight

                // Starts the driver scan method
                Procedure OnClick
                    Integer eInfoType

                    Get InfoType of oDriversInfoComboForm to eInfoType
                    Send EnumPrinterDrivers of oPrinterDriversHandler "" "All" eInfoType
                    Send FillDriversList Of oDriversList
                End_Procedure
            End_Object
        End_Object
    End_Object
End_Object
