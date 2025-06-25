namespace ExcelEssentials
{
    partial class MiscRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MiscRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MiscRibbon));
            this.miscTab = this.Factory.CreateRibbonTab();
            this.importGroup = this.Factory.CreateRibbonGroup();
            this.importSheetOrTxtFileSplitButton = this.Factory.CreateRibbonSplitButton();
            this.importSheetOrTxtFileAdvButton = this.Factory.CreateRibbonButton();
            this.importTxtFileLegacyButton = this.Factory.CreateRibbonButton();
            this.modifiersGroup = this.Factory.CreateRibbonGroup();
            this.changeToTextButton = this.Factory.CreateRibbonButton();
            this.changeToValueButton = this.Factory.CreateRibbonButton();
            this.evaluateFormulaButton = this.Factory.CreateRibbonButton();
            this.repasteAsValuesButton = this.Factory.CreateRibbonButton();
            this.sortingAbsButton = this.Factory.CreateRibbonButton();
            this.copyAsPictureButton = this.Factory.CreateRibbonButton();
            this.separator11 = this.Factory.CreateRibbonSeparator();
            this.removeEmptyButton = this.Factory.CreateRibbonButton();
            this.removeErrSplitBtn = this.Factory.CreateRibbonSplitButton();
            this.removeNaButton = this.Factory.CreateRibbonButton();
            this.removeFormattingButton = this.Factory.CreateRibbonButton();
            this.removeHiddenRowsSplitButton = this.Factory.CreateRibbonSplitButton();
            this.removeHiddenColumnsButton = this.Factory.CreateRibbonButton();
            this.clearRangeOutsideButton = this.Factory.CreateRibbonButton();
            this.removeDuplicatesSplitButton = this.Factory.CreateRibbonSplitButton();
            this.removeDuplicatesInColumnsButton = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.prependTextSplitButton = this.Factory.CreateRibbonSplitButton();
            this.appendTextButton = this.Factory.CreateRibbonButton();
            this.trimButton = this.Factory.CreateRibbonButton();
            this.formatNumberSplitButton = this.Factory.CreateRibbonSplitButton();
            this.formatStringToDateButton = this.Factory.CreateRibbonButton();
            this.separator13 = this.Factory.CreateRibbonSeparator();
            this.selectWithoutHeadersButton = this.Factory.CreateRibbonButton();
            this.fillEmptyWithAboveValueButton = this.Factory.CreateRibbonButton();
            this.copyDelimitedValuesButton = this.Factory.CreateRibbonButton();
            this.filterGroup = this.Factory.CreateRibbonGroup();
            this.filterColumnSplitButton = this.Factory.CreateRibbonSplitButton();
            this.filterColumnNotInRangeBtn = this.Factory.CreateRibbonButton();
            this.filterColumnFromRangeInRangeButton = this.Factory.CreateRibbonButton();
            this.filterColumnFromRangeNotInRangeButton = this.Factory.CreateRibbonButton();
            this.filterColumnInRegexButton = this.Factory.CreateRibbonButton();
            this.filterColumnNotInRegexButton = this.Factory.CreateRibbonButton();
            this.filterColumnFlipFilterBtn = this.Factory.CreateRibbonButton();
            this.hideRowsWithTextSplitButton = this.Factory.CreateRibbonSplitButton();
            this.hideColumnsWithTextButton = this.Factory.CreateRibbonButton();
            this.takeRowsWithTextButton = this.Factory.CreateRibbonButton();
            this.validationGroup = this.Factory.CreateRibbonGroup();
            this.splitButton1 = this.Factory.CreateRibbonSplitButton();
            this.colorRowsUniqueButton = this.Factory.CreateRibbonButton();
            this.colorRowsWithTextSplitButton = this.Factory.CreateRibbonSplitButton();
            this.colorColumnsWithTextButton = this.Factory.CreateRibbonButton();
            this.colorCellsWithTextButton = this.Factory.CreateRibbonButton();
            this.colorRowsButton = this.Factory.CreateRibbonButton();
            this.formatTrueFalseButton = this.Factory.CreateRibbonButton();
            this.searchGroup = this.Factory.CreateRibbonGroup();
            this.searchDialogButton = this.Factory.CreateRibbonButton();
            this.fileAndExportGroup = this.Factory.CreateRibbonGroup();
            this.saveSelectedWorksheetsAsXlsxSplitBtn = this.Factory.CreateRibbonSplitButton();
            this.saveAllWorksheetsAsXlsxButton = this.Factory.CreateRibbonButton();
            this.saveSelectedWorksheetsAsTxtButton = this.Factory.CreateRibbonButton();
            this.saveAllWorksheetsAsTxtButton = this.Factory.CreateRibbonButton();
            this.duplicateWorksheetsSplitBtn = this.Factory.CreateRibbonSplitButton();
            this.duplicateWorksheetsToNewWorkbookBtn = this.Factory.CreateRibbonButton();
            this.duplicateWorkbookBtn = this.Factory.CreateRibbonButton();
            this.saveThisWorksheetAsTxt = this.Factory.CreateRibbonButton();
            this.divideTableToPartsAndSaveButton = this.Factory.CreateRibbonButton();
            this.getFilePathButton = this.Factory.CreateRibbonButton();
            this.exportMacrosButton = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.deleteWorksheetButton = this.Factory.CreateRibbonButton();
            this.deleteOtherWorksheetsButton = this.Factory.CreateRibbonButton();
            this.deleteWorkbookButton = this.Factory.CreateRibbonButton();
            this.macroGroup = this.Factory.CreateRibbonGroup();
            this.runMacroButton = this.Factory.CreateRibbonButton();
            this.runCustomFormButton = this.Factory.CreateRibbonButton();
            this.pivotToolsTab = this.Factory.CreateRibbonTab();
            this.pivotTemplatesGroup = this.Factory.CreateRibbonGroup();
            this.createPivotFromTemplateMenu = this.Factory.CreateRibbonMenu();
            this.createPvFromLoadedButton = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.createPvFromInboundButton = this.Factory.CreateRibbonButton();
            this.createPvFromInboundPlusButton = this.Factory.CreateRibbonButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.createPvFromOutboundButton = this.Factory.CreateRibbonButton();
            this.createPvFromOutboundPlusButton = this.Factory.CreateRibbonButton();
            this.separator5 = this.Factory.CreateRibbonSeparator();
            this.createPvFromMNTXButton = this.Factory.CreateRibbonButton();
            this.separator6 = this.Factory.CreateRibbonSeparator();
            this.createPvFromYottaButton = this.Factory.CreateRibbonButton();
            this.separator7 = this.Factory.CreateRibbonSeparator();
            this.createPvFromSAPCButton = this.Factory.CreateRibbonButton();
            this.createPvFromSAPCLoadedButton = this.Factory.CreateRibbonButton();
            this.separator8 = this.Factory.CreateRibbonSeparator();
            this.createPvFromVBAKButton = this.Factory.CreateRibbonButton();
            this.createPvFromWBRKButton = this.Factory.CreateRibbonButton();
            this.createPvFromVBPAButton = this.Factory.CreateRibbonButton();
            this.createPvFromVBAPButton = this.Factory.CreateRibbonButton();
            this.createPvFromVBPAandKNA1Button = this.Factory.CreateRibbonButton();
            this.separator9 = this.Factory.CreateRibbonSeparator();
            this.createPvFromQlikButton = this.Factory.CreateRibbonButton();
            this.separator10 = this.Factory.CreateRibbonSeparator();
            this.customPivotsTemplatesMenu = this.Factory.CreateRibbonMenu();
            this.createPvFromCustom1Button = this.Factory.CreateRibbonButton();
            this.createPvFromCustom2Button = this.Factory.CreateRibbonButton();
            this.createPvFromCustom3Button = this.Factory.CreateRibbonButton();
            this.createPvFromCustom4Button = this.Factory.CreateRibbonButton();
            this.createPvFromCustom5Button = this.Factory.CreateRibbonButton();
            this.createPvFromCustom6Button = this.Factory.CreateRibbonButton();
            this.createPvFromCustom7Button = this.Factory.CreateRibbonButton();
            this.createPvFromCustom8Button = this.Factory.CreateRibbonButton();
            this.createPvFromCustom9Button = this.Factory.CreateRibbonButton();
            this.generatePivotTemlateCodeButton = this.Factory.CreateRibbonButton();
            this.runPvTemplateBtn = this.Factory.CreateRibbonButton();
            this.pivotFormatGroup = this.Factory.CreateRibbonGroup();
            this.formatClickedPivotButton = this.Factory.CreateRibbonButton();
            this.formatAllPivotButton = this.Factory.CreateRibbonButton();
            this.grandTotalsToggleButton = this.Factory.CreateRibbonToggleButton();
            this.subtotalsToggleButton = this.Factory.CreateRibbonToggleButton();
            this.pivotToolsGroup = this.Factory.CreateRibbonGroup();
            this.changePivotTableSourceButton = this.Factory.CreateRibbonButton();
            this.updatePivotTableSourceButton = this.Factory.CreateRibbonButton();
            this.refreshPivotsButton = this.Factory.CreateRibbonButton();
            this.combinedTableFromPvValuesButton = this.Factory.CreateRibbonButton();
            this.dataImportTab = this.Factory.CreateRibbonTab();
            this.sqlImportGroup = this.Factory.CreateRibbonGroup();
            this.loadToDataTableButton = this.Factory.CreateRibbonButton();
            this.sapImportGroup = this.Factory.CreateRibbonGroup();
            this.runS4ExtractButton = this.Factory.CreateRibbonButton();
            this.sdeImportGroup = this.Factory.CreateRibbonGroup();
            this.runSdeButton = this.Factory.CreateRibbonButton();
            this.sdeQueryComboBox = this.Factory.CreateRibbonComboBox();
            this.sdeInstancesEditBox = this.Factory.CreateRibbonEditBox();
            this.browserGroup = this.Factory.CreateRibbonGroup();
            this.browserButton = this.Factory.CreateRibbonButton();
            this.browserWebsitesComboBox = this.Factory.CreateRibbonComboBox();
            this.importFromBrowserCheckBox = this.Factory.CreateRibbonCheckBox();
            this.goToPropertiesButton = this.Factory.CreateRibbonButton();
            this.updateMacrosButton = this.Factory.CreateRibbonButton();
            this.createMacroUpdateButton = this.Factory.CreateRibbonButton();
            this.checkMacrosButton = this.Factory.CreateRibbonButton();
            this.excelEssentialsPackInfoBtn = this.Factory.CreateRibbonButton();
            this.miscTab.SuspendLayout();
            this.importGroup.SuspendLayout();
            this.modifiersGroup.SuspendLayout();
            this.filterGroup.SuspendLayout();
            this.validationGroup.SuspendLayout();
            this.searchGroup.SuspendLayout();
            this.fileAndExportGroup.SuspendLayout();
            this.macroGroup.SuspendLayout();
            this.pivotToolsTab.SuspendLayout();
            this.pivotTemplatesGroup.SuspendLayout();
            this.pivotFormatGroup.SuspendLayout();
            this.pivotToolsGroup.SuspendLayout();
            this.dataImportTab.SuspendLayout();
            this.sqlImportGroup.SuspendLayout();
            this.sapImportGroup.SuspendLayout();
            this.sdeImportGroup.SuspendLayout();
            this.browserGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // miscTab
            // 
            this.miscTab.Groups.Add(this.importGroup);
            this.miscTab.Groups.Add(this.modifiersGroup);
            this.miscTab.Groups.Add(this.filterGroup);
            this.miscTab.Groups.Add(this.validationGroup);
            this.miscTab.Groups.Add(this.searchGroup);
            this.miscTab.Groups.Add(this.fileAndExportGroup);
            this.miscTab.Groups.Add(this.macroGroup);
            this.miscTab.Label = "Misc";
            this.miscTab.Name = "miscTab";
            // 
            // importGroup
            // 
            this.importGroup.Items.Add(this.importSheetOrTxtFileSplitButton);
            this.importGroup.Label = "Importing";
            this.importGroup.Name = "importGroup";
            // 
            // importSheetOrTxtFileSplitButton
            // 
            this.importSheetOrTxtFileSplitButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.importSheetOrTxtFileSplitButton.Items.Add(this.importSheetOrTxtFileAdvButton);
            this.importSheetOrTxtFileSplitButton.Items.Add(this.importTxtFileLegacyButton);
            this.importSheetOrTxtFileSplitButton.Label = "Import worksheet or txt file";
            this.importSheetOrTxtFileSplitButton.Name = "importSheetOrTxtFileSplitButton";
            this.importSheetOrTxtFileSplitButton.OfficeImageId = "ImportOpml";
            this.importSheetOrTxtFileSplitButton.ScreenTip = "Import worksheet or txt file";
            this.importSheetOrTxtFileSplitButton.SuperTip = "Will create window that will accept txt/csv or Excel file and will import delimit" +
    "ed table or sheet/sheets";
            this.importSheetOrTxtFileSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.importSheetOrTxtFile_Click);
            // 
            // importSheetOrTxtFileAdvButton
            // 
            this.importSheetOrTxtFileAdvButton.Label = "Import worksheet or txt file advanced";
            this.importSheetOrTxtFileAdvButton.Name = "importSheetOrTxtFileAdvButton";
            this.importSheetOrTxtFileAdvButton.OfficeImageId = "ImportOpml";
            this.importSheetOrTxtFileAdvButton.ScreenTip = "Advanced Import worksheet or txt file (beta)";
            this.importSheetOrTxtFileAdvButton.ShowImage = true;
            this.importSheetOrTxtFileAdvButton.SuperTip = "Will create window that will accept txt/csv file and will import delimited table " +
    "to Excel and it should handle all text in \'\"\' quotes properly";
            this.importSheetOrTxtFileAdvButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.importSheetOrTxtFileAdv_Click);
            // 
            // importTxtFileLegacyButton
            // 
            this.importTxtFileLegacyButton.Label = "Legacy Import worksheet or txt file";
            this.importTxtFileLegacyButton.Name = "importTxtFileLegacyButton";
            this.importTxtFileLegacyButton.OfficeImageId = "ImportExcel";
            this.importTxtFileLegacyButton.ScreenTip = "Legacy Import worksheet or txt file";
            this.importTxtFileLegacyButton.ShowImage = true;
            this.importTxtFileLegacyButton.SuperTip = "Will create window that will accept txt/csv file and will import delimited table " +
    "as text";
            this.importTxtFileLegacyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.importTxtFileLegacyButton_Click);
            // 
            // modifiersGroup
            // 
            this.modifiersGroup.Items.Add(this.changeToTextButton);
            this.modifiersGroup.Items.Add(this.changeToValueButton);
            this.modifiersGroup.Items.Add(this.evaluateFormulaButton);
            this.modifiersGroup.Items.Add(this.repasteAsValuesButton);
            this.modifiersGroup.Items.Add(this.sortingAbsButton);
            this.modifiersGroup.Items.Add(this.copyAsPictureButton);
            this.modifiersGroup.Items.Add(this.separator11);
            this.modifiersGroup.Items.Add(this.removeEmptyButton);
            this.modifiersGroup.Items.Add(this.removeErrSplitBtn);
            this.modifiersGroup.Items.Add(this.removeFormattingButton);
            this.modifiersGroup.Items.Add(this.removeHiddenRowsSplitButton);
            this.modifiersGroup.Items.Add(this.clearRangeOutsideButton);
            this.modifiersGroup.Items.Add(this.removeDuplicatesSplitButton);
            this.modifiersGroup.Items.Add(this.separator1);
            this.modifiersGroup.Items.Add(this.prependTextSplitButton);
            this.modifiersGroup.Items.Add(this.trimButton);
            this.modifiersGroup.Items.Add(this.formatNumberSplitButton);
            this.modifiersGroup.Items.Add(this.separator13);
            this.modifiersGroup.Items.Add(this.selectWithoutHeadersButton);
            this.modifiersGroup.Items.Add(this.fillEmptyWithAboveValueButton);
            this.modifiersGroup.Items.Add(this.copyDelimitedValuesButton);
            this.modifiersGroup.Label = "Modifiers";
            this.modifiersGroup.Name = "modifiersGroup";
            // 
            // changeToTextButton
            // 
            this.changeToTextButton.Label = "To text";
            this.changeToTextButton.Name = "changeToTextButton";
            this.changeToTextButton.OfficeImageId = "AsianLayoutPhoneticGuide";
            this.changeToTextButton.ScreenTip = "Convert to text";
            this.changeToTextButton.ShowImage = true;
            this.changeToTextButton.SuperTip = "Select range and it will change it to text in the same way as Text to Columns doe" +
    "s.";
            this.changeToTextButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.changeToTextButton_Click);
            // 
            // changeToValueButton
            // 
            this.changeToValueButton.Label = "To value";
            this.changeToValueButton.Name = "changeToValueButton";
            this.changeToValueButton.OfficeImageId = "EquationMatrixGallery";
            this.changeToValueButton.ScreenTip = "Convert to value";
            this.changeToValueButton.ShowImage = true;
            this.changeToValueButton.SuperTip = "Select range and it will make it general value type, so it will remove text forma" +
    "t from it.";
            this.changeToValueButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.changeToValueButton_Click);
            // 
            // evaluateFormulaButton
            // 
            this.evaluateFormulaButton.Enabled = false;
            this.evaluateFormulaButton.Label = "Ev formula";
            this.evaluateFormulaButton.Name = "evaluateFormulaButton";
            this.evaluateFormulaButton.OfficeImageId = "ShowFormulas";
            this.evaluateFormulaButton.ScreenTip = "Evaluate formula";
            this.evaluateFormulaButton.ShowImage = true;
            this.evaluateFormulaButton.SuperTip = "Select rectagular range and it should evaluate and replace formulas to value. For" +
    " arrays select first cell. (May be slower than \"Repaste as values\")";
            this.evaluateFormulaButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.evaluateFormulaButton_Click);
            // 
            // repasteAsValuesButton
            // 
            this.repasteAsValuesButton.Label = "Repaste as val";
            this.repasteAsValuesButton.Name = "repasteAsValuesButton";
            this.repasteAsValuesButton.OfficeImageId = "PasteValuesAndNumberFormatting";
            this.repasteAsValuesButton.ScreenTip = "Repaste as values";
            this.repasteAsValuesButton.ShowImage = true;
            this.repasteAsValuesButton.SuperTip = "Select range or ranges and it will repaste it as values in the same place";
            this.repasteAsValuesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.repasteAsValuesButton_Click);
            // 
            // sortingAbsButton
            // 
            this.sortingAbsButton.Label = "Sort by abs(val)";
            this.sortingAbsButton.Name = "sortingAbsButton";
            this.sortingAbsButton.OfficeImageId = "Sort";
            this.sortingAbsButton.ScreenTip = "Sort by abs(values)";
            this.sortingAbsButton.ShowImage = true;
            this.sortingAbsButton.SuperTip = "Click cell in column that have numbers and it should insert extra column with abs" +
    "olute values and sort from biggest to smallest and remove column automatically (" +
    "do not use for pivot tables)";
            this.sortingAbsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sortingAbsButton_Click);
            // 
            // copyAsPictureButton
            // 
            this.copyAsPictureButton.Label = "Copy as picture";
            this.copyAsPictureButton.Name = "copyAsPictureButton";
            this.copyAsPictureButton.OfficeImageId = "Camera";
            this.copyAsPictureButton.ScreenTip = "Copy as picture";
            this.copyAsPictureButton.ShowImage = true;
            this.copyAsPictureButton.SuperTip = "Copy selected rectagle range as picture (try again in case or error as it appear " +
    "randomly)";
            this.copyAsPictureButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.copyAsPictureButton_Click);
            // 
            // separator11
            // 
            this.separator11.Name = "separator11";
            // 
            // removeEmptyButton
            // 
            this.removeEmptyButton.Label = "Remove empty";
            this.removeEmptyButton.Name = "removeEmptyButton";
            this.removeEmptyButton.OfficeImageId = "DataViewConditionalFormatting";
            this.removeEmptyButton.ScreenTip = "Remove empty cells";
            this.removeEmptyButton.ShowImage = true;
            this.removeEmptyButton.SuperTip = "Select range and it will remove empty cells from it (try to use on smaller data o" +
    "r save everything before)";
            this.removeEmptyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeEmptyButton_Click);
            // 
            // removeErrSplitBtn
            // 
            this.removeErrSplitBtn.Items.Add(this.removeNaButton);
            this.removeErrSplitBtn.Label = "Remove Err";
            this.removeErrSplitBtn.Name = "removeErrSplitBtn";
            this.removeErrSplitBtn.OfficeImageId = "ConditionalFormattingClearMenu";
            this.removeErrSplitBtn.ScreenTip = "Blank Error cells";
            this.removeErrSplitBtn.SuperTip = "Select range and it will blank all cells with error";
            this.removeErrSplitBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeErrSplitBtn_Click);
            // 
            // removeNaButton
            // 
            this.removeNaButton.Label = "Remove #N/A";
            this.removeNaButton.Name = "removeNaButton";
            this.removeNaButton.OfficeImageId = "ConditionalFormattingClearMenu";
            this.removeNaButton.ScreenTip = "Blank #N/A cells";
            this.removeNaButton.ShowImage = true;
            this.removeNaButton.SuperTip = "Select range and it will blank N/A values also work for text \"#N/A\"";
            this.removeNaButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeNaButton_Click);
            // 
            // removeFormattingButton
            // 
            this.removeFormattingButton.Label = "Remove format";
            this.removeFormattingButton.Name = "removeFormattingButton";
            this.removeFormattingButton.OfficeImageId = "HighlightClear";
            this.removeFormattingButton.ScreenTip = "Remove formatting from cells";
            this.removeFormattingButton.ShowImage = true;
            this.removeFormattingButton.SuperTip = "Select range and it will remove formatting from it";
            this.removeFormattingButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeFormattingButton_Click);
            // 
            // removeHiddenRowsSplitButton
            // 
            this.removeHiddenRowsSplitButton.Items.Add(this.removeHiddenColumnsButton);
            this.removeHiddenRowsSplitButton.Label = "Remove hidden rows";
            this.removeHiddenRowsSplitButton.Name = "removeHiddenRowsSplitButton";
            this.removeHiddenRowsSplitButton.OfficeImageId = "DeleteRows";
            this.removeHiddenRowsSplitButton.ScreenTip = "Remove hidden rows";
            this.removeHiddenRowsSplitButton.SuperTip = "Remove hidden rows from selected range (or region when only one cell selected) (c" +
    "an be slow on large tables)";
            this.removeHiddenRowsSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeHiddenRowsSplitButton_Click);
            // 
            // removeHiddenColumnsButton
            // 
            this.removeHiddenColumnsButton.Label = "Remove hidden columns";
            this.removeHiddenColumnsButton.Name = "removeHiddenColumnsButton";
            this.removeHiddenColumnsButton.OfficeImageId = "DeleteColumns";
            this.removeHiddenColumnsButton.ScreenTip = "Remove hidden columns";
            this.removeHiddenColumnsButton.ShowImage = true;
            this.removeHiddenColumnsButton.SuperTip = "Remove hidden columns from selected range (or region when only one cell selected)" +
    " (can be slow on large tables)";
            this.removeHiddenColumnsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeHiddenColumnsButton_Click);
            // 
            // clearRangeOutsideButton
            // 
            this.clearRangeOutsideButton.Label = "Clear outside rng/region";
            this.clearRangeOutsideButton.Name = "clearRangeOutsideButton";
            this.clearRangeOutsideButton.OfficeImageId = "CellStyleNew";
            this.clearRangeOutsideButton.ScreenTip = "Clear outside range/region";
            this.clearRangeOutsideButton.ShowImage = true;
            this.clearRangeOutsideButton.SuperTip = "Clear everything outside region when one cell selected or outside range when sele" +
    "cted rectangular range";
            this.clearRangeOutsideButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clearRangeOutsideButton_Click);
            // 
            // removeDuplicatesSplitButton
            // 
            this.removeDuplicatesSplitButton.Items.Add(this.removeDuplicatesInColumnsButton);
            this.removeDuplicatesSplitButton.Label = "Remove duplicates";
            this.removeDuplicatesSplitButton.Name = "removeDuplicatesSplitButton";
            this.removeDuplicatesSplitButton.OfficeImageId = "RemoveDuplicates";
            this.removeDuplicatesSplitButton.ScreenTip = "Remove duplicates";
            this.removeDuplicatesSplitButton.SuperTip = "Select recatangular range and it will remove duplicate rows from it";
            this.removeDuplicatesSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeDuplicatesButton_Click);
            // 
            // removeDuplicatesInColumnsButton
            // 
            this.removeDuplicatesInColumnsButton.Label = "Remove duplicates in columns";
            this.removeDuplicatesInColumnsButton.Name = "removeDuplicatesInColumnsButton";
            this.removeDuplicatesInColumnsButton.OfficeImageId = "RemoveDuplicates";
            this.removeDuplicatesInColumnsButton.ScreenTip = "Remove duplicates in each column";
            this.removeDuplicatesInColumnsButton.ShowImage = true;
            this.removeDuplicatesInColumnsButton.SuperTip = "Select recatangular range and it will remove duplicate cells from every column";
            this.removeDuplicatesInColumnsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeDuplicatesInColumnsButton_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // prependTextSplitButton
            // 
            this.prependTextSplitButton.Items.Add(this.appendTextButton);
            this.prependTextSplitButton.Label = "Prepend";
            this.prependTextSplitButton.Name = "prependTextSplitButton";
            this.prependTextSplitButton.OfficeImageId = "OutlineDemoteToBodyText";
            this.prependTextSplitButton.ScreenTip = "Prepend values with text";
            this.prependTextSplitButton.SuperTip = "Select range and it will prepend all values with given text from form that will a" +
    "ppear";
            this.prependTextSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.prependTextSplitButton_Click);
            // 
            // appendTextButton
            // 
            this.appendTextButton.Label = "Append";
            this.appendTextButton.Name = "appendTextButton";
            this.appendTextButton.OfficeImageId = "LeftArrow2";
            this.appendTextButton.ScreenTip = "Append values with text";
            this.appendTextButton.ShowImage = true;
            this.appendTextButton.SuperTip = "Select range and it will append all values with given text from form that will ap" +
    "pear (for numbers formatted as text it may convert back to number after appendin" +
    "g with number)";
            this.appendTextButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.appendTextButton_Click);
            // 
            // trimButton
            // 
            this.trimButton.Label = "Trim";
            this.trimButton.Name = "trimButton";
            this.trimButton.OfficeImageId = "TextDirectionContext";
            this.trimButton.ScreenTip = "Trim leading ant trailing spaces";
            this.trimButton.ShowImage = true;
            this.trimButton.SuperTip = "Select range and it will remove leading and trailing spaces, tabs and new lines f" +
    "rom values";
            this.trimButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.trimButton_Click);
            // 
            // formatNumberSplitButton
            // 
            this.formatNumberSplitButton.Items.Add(this.formatStringToDateButton);
            this.formatNumberSplitButton.Label = "Format 123";
            this.formatNumberSplitButton.Name = "formatNumberSplitButton";
            this.formatNumberSplitButton.OfficeImageId = "DollarSign";
            this.formatNumberSplitButton.ScreenTip = "Custom format";
            this.formatNumberSplitButton.SuperTip = "Select any range and it will apply number format for numbers, date, if field is f" +
    "ormatted as date, text is not affected. Apply again to change displayed precisio" +
    "n (just format, does not modify values)";
            this.formatNumberSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.formatNumberSplitButton_Click);
            // 
            // formatStringToDateButton
            // 
            this.formatStringToDateButton.Label = "Format string to date";
            this.formatStringToDateButton.Name = "formatStringToDateButton";
            this.formatStringToDateButton.OfficeImageId = "DateInsert";
            this.formatStringToDateButton.ScreenTip = "Format date that is text or number to date value";
            this.formatStringToDateButton.ShowImage = true;
            this.formatStringToDateButton.SuperTip = "Select single column range of string date values and it will show form to input f" +
    "ormat of date to parse the data and it will paste it into selected cell after cl" +
    "icking OK";
            this.formatStringToDateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.formatStringToDateButton_Click);
            // 
            // separator13
            // 
            this.separator13.Name = "separator13";
            // 
            // selectWithoutHeadersButton
            // 
            this.selectWithoutHeadersButton.Label = "Select w/o headers";
            this.selectWithoutHeadersButton.Name = "selectWithoutHeadersButton";
            this.selectWithoutHeadersButton.OfficeImageId = "SelectWholeLayout";
            this.selectWithoutHeadersButton.ScreenTip = "Select region without headers";
            this.selectWithoutHeadersButton.ShowImage = true;
            this.selectWithoutHeadersButton.SuperTip = "Click on cell inside your table and click it will select rectangular range withou" +
    "t first row";
            this.selectWithoutHeadersButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.selectWithoutHeadersButton_Click);
            // 
            // fillEmptyWithAboveValueButton
            // 
            this.fillEmptyWithAboveValueButton.Label = "Fill with above value";
            this.fillEmptyWithAboveValueButton.Name = "fillEmptyWithAboveValueButton";
            this.fillEmptyWithAboveValueButton.OfficeImageId = "TablePropertiesHeightInfoPath";
            this.fillEmptyWithAboveValueButton.ScreenTip = "Fill empty cells with above cell\'s content";
            this.fillEmptyWithAboveValueButton.ShowImage = true;
            this.fillEmptyWithAboveValueButton.SuperTip = "Select rectangular range and it will fill empty cells with above cell\'s formula o" +
    "r value";
            this.fillEmptyWithAboveValueButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.fillEmptyWithAboveValueButton_Click);
            // 
            // copyDelimitedValuesButton
            // 
            this.copyDelimitedValuesButton.Label = "Copy delimited";
            this.copyDelimitedValuesButton.Name = "copyDelimitedValuesButton";
            this.copyDelimitedValuesButton.OfficeImageId = "MultiplicationSign";
            this.copyDelimitedValuesButton.ScreenTip = "Copy selected values delimited";
            this.copyDelimitedValuesButton.ShowImage = true;
            this.copyDelimitedValuesButton.SuperTip = "Will display form to set format to copy and will prepend and append and delimit w" +
    "ith choosen values and copy formatted string to clipboard";
            this.copyDelimitedValuesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.copyDelimitedValuesButton_Click);
            // 
            // filterGroup
            // 
            this.filterGroup.Items.Add(this.filterColumnSplitButton);
            this.filterGroup.Items.Add(this.hideRowsWithTextSplitButton);
            this.filterGroup.Items.Add(this.takeRowsWithTextButton);
            this.filterGroup.Label = "Filtering";
            this.filterGroup.Name = "filterGroup";
            // 
            // filterColumnSplitButton
            // 
            this.filterColumnSplitButton.Items.Add(this.filterColumnNotInRangeBtn);
            this.filterColumnSplitButton.Items.Add(this.filterColumnFromRangeInRangeButton);
            this.filterColumnSplitButton.Items.Add(this.filterColumnFromRangeNotInRangeButton);
            this.filterColumnSplitButton.Items.Add(this.filterColumnInRegexButton);
            this.filterColumnSplitButton.Items.Add(this.filterColumnNotInRegexButton);
            this.filterColumnSplitButton.Items.Add(this.filterColumnFlipFilterBtn);
            this.filterColumnSplitButton.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.filterColumnSplitButton.Label = "Filter column";
            this.filterColumnSplitButton.Name = "filterColumnSplitButton";
            this.filterColumnSplitButton.OfficeImageId = "FilterBySelection";
            this.filterColumnSplitButton.ScreenTip = "Filter column";
            this.filterColumnSplitButton.SuperTip = resources.GetString("filterColumnSplitButton.SuperTip");
            this.filterColumnSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.filterColumnInRangeSplitButton_Click);
            // 
            // filterColumnNotInRangeBtn
            // 
            this.filterColumnNotInRangeBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.filterColumnNotInRangeBtn.Label = "Filter column not in range";
            this.filterColumnNotInRangeBtn.Name = "filterColumnNotInRangeBtn";
            this.filterColumnNotInRangeBtn.OfficeImageId = "FilterClear";
            this.filterColumnNotInRangeBtn.ScreenTip = "Filter column not in range";
            this.filterColumnNotInRangeBtn.ShowImage = true;
            this.filterColumnNotInRangeBtn.SuperTip = resources.GetString("filterColumnNotInRangeBtn.SuperTip");
            this.filterColumnNotInRangeBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.filterColumnNotInRangeButton_Click);
            // 
            // filterColumnFromRangeInRangeButton
            // 
            this.filterColumnFromRangeInRangeButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.filterColumnFromRangeInRangeButton.Label = "Filter column from range";
            this.filterColumnFromRangeInRangeButton.Name = "filterColumnFromRangeInRangeButton";
            this.filterColumnFromRangeInRangeButton.OfficeImageId = "FilterAdvancedByForm";
            this.filterColumnFromRangeInRangeButton.ScreenTip = "Filter column from range";
            this.filterColumnFromRangeInRangeButton.ShowImage = true;
            this.filterColumnFromRangeInRangeButton.SuperTip = resources.GetString("filterColumnFromRangeInRangeButton.SuperTip");
            this.filterColumnFromRangeInRangeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.filterColumnFromRangeInRangeButton_Click);
            // 
            // filterColumnFromRangeNotInRangeButton
            // 
            this.filterColumnFromRangeNotInRangeButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.filterColumnFromRangeNotInRangeButton.Label = "Filter column from range not in";
            this.filterColumnFromRangeNotInRangeButton.Name = "filterColumnFromRangeNotInRangeButton";
            this.filterColumnFromRangeNotInRangeButton.OfficeImageId = "FilterAdvancedByForm";
            this.filterColumnFromRangeNotInRangeButton.ScreenTip = "Filter column from range not in";
            this.filterColumnFromRangeNotInRangeButton.ShowImage = true;
            this.filterColumnFromRangeNotInRangeButton.SuperTip = resources.GetString("filterColumnFromRangeNotInRangeButton.SuperTip");
            this.filterColumnFromRangeNotInRangeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.filterColumnFromRangeNotInRangeButton_Click);
            // 
            // filterColumnInRegexButton
            // 
            this.filterColumnInRegexButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.filterColumnInRegexButton.Label = "Filter column by Regex pattern";
            this.filterColumnInRegexButton.Name = "filterColumnInRegexButton";
            this.filterColumnInRegexButton.OfficeImageId = "SortFilterMenu";
            this.filterColumnInRegexButton.ScreenTip = "Filter column by Regex pattern";
            this.filterColumnInRegexButton.ShowImage = true;
            this.filterColumnInRegexButton.SuperTip = resources.GetString("filterColumnInRegexButton.SuperTip");
            this.filterColumnInRegexButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.filterColumnInRegexButton_Click);
            // 
            // filterColumnNotInRegexButton
            // 
            this.filterColumnNotInRegexButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.filterColumnNotInRegexButton.Label = "Filter column not in Regex pattern";
            this.filterColumnNotInRegexButton.Name = "filterColumnNotInRegexButton";
            this.filterColumnNotInRegexButton.OfficeImageId = "SortFilterMenu";
            this.filterColumnNotInRegexButton.ScreenTip = "Filter column not in Regex pattern";
            this.filterColumnNotInRegexButton.ShowImage = true;
            this.filterColumnNotInRegexButton.SuperTip = resources.GetString("filterColumnNotInRegexButton.SuperTip");
            this.filterColumnNotInRegexButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.filterColumnNotInRegexButton_Click);
            // 
            // filterColumnFlipFilterBtn
            // 
            this.filterColumnFlipFilterBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.filterColumnFlipFilterBtn.Label = "Flip filter in column";
            this.filterColumnFlipFilterBtn.Name = "filterColumnFlipFilterBtn";
            this.filterColumnFlipFilterBtn.OfficeImageId = "FilterReapply";
            this.filterColumnFlipFilterBtn.ScreenTip = "Flip filter in column";
            this.filterColumnFlipFilterBtn.ShowImage = true;
            this.filterColumnFlipFilterBtn.SuperTip = resources.GetString("filterColumnFlipFilterBtn.SuperTip");
            this.filterColumnFlipFilterBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.filterColumnFlipFilterBtn_Click);
            // 
            // hideRowsWithTextSplitButton
            // 
            this.hideRowsWithTextSplitButton.Items.Add(this.hideColumnsWithTextButton);
            this.hideRowsWithTextSplitButton.Label = "Hide rows with text";
            this.hideRowsWithTextSplitButton.Name = "hideRowsWithTextSplitButton";
            this.hideRowsWithTextSplitButton.OfficeImageId = "GroupTableMerge";
            this.hideRowsWithTextSplitButton.ScreenTip = "Hide rows with text";
            this.hideRowsWithTextSplitButton.SuperTip = "Will show dialog that will ask for text and it will hide the entire rows from sel" +
    "ected rectangular range ";
            this.hideRowsWithTextSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.hideRowsWithTextSplitButton_Click);
            // 
            // hideColumnsWithTextButton
            // 
            this.hideColumnsWithTextButton.Label = "Hide columns with text";
            this.hideColumnsWithTextButton.Name = "hideColumnsWithTextButton";
            this.hideColumnsWithTextButton.OfficeImageId = "GroupTableMerge";
            this.hideColumnsWithTextButton.ScreenTip = "Hide columns with text";
            this.hideColumnsWithTextButton.ShowImage = true;
            this.hideColumnsWithTextButton.SuperTip = "Will show dialog that will ask for text and it will hide the entire columns from " +
    "selected rectangular range ";
            this.hideColumnsWithTextButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.hideColumnsWithTextButton_Click);
            // 
            // takeRowsWithTextButton
            // 
            this.takeRowsWithTextButton.Label = "Take rows with text";
            this.takeRowsWithTextButton.Name = "takeRowsWithTextButton";
            this.takeRowsWithTextButton.OfficeImageId = "TableInsertRowsBelow";
            this.takeRowsWithTextButton.ScreenTip = "Take rows with text";
            this.takeRowsWithTextButton.ShowImage = true;
            this.takeRowsWithTextButton.SuperTip = "Select rectangular range and it will show dialog with text input that will take r" +
    "ange rows that contain that text with headers to new sheet";
            this.takeRowsWithTextButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.takeRowsWithTextButton_Click);
            // 
            // validationGroup
            // 
            this.validationGroup.Items.Add(this.splitButton1);
            this.validationGroup.Items.Add(this.colorRowsWithTextSplitButton);
            this.validationGroup.Items.Add(this.formatTrueFalseButton);
            this.validationGroup.Label = "Validation";
            this.validationGroup.Name = "validationGroup";
            // 
            // splitButton1
            // 
            this.splitButton1.Items.Add(this.colorRowsUniqueButton);
            this.splitButton1.Label = "Color rows unique";
            this.splitButton1.Name = "splitButton1";
            this.splitButton1.OfficeImageId = "GroupResourceGraphFormat";
            this.splitButton1.ScreenTip = "Color rows unique";
            this.splitButton1.SuperTip = "Color rows in selected rectangular range based on first column unique values in u" +
    "nique colors (~2000 colors and then thay will repeat and may end up next to each" +
    " other)";
            this.splitButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.colorRowsUniqueSplitButton_Click);
            // 
            // colorRowsUniqueButton
            // 
            this.colorRowsUniqueButton.Label = "Color cells unique";
            this.colorRowsUniqueButton.Name = "colorRowsUniqueButton";
            this.colorRowsUniqueButton.OfficeImageId = "FormatBarStylesMenu";
            this.colorRowsUniqueButton.ScreenTip = "Color cells unique";
            this.colorRowsUniqueButton.ShowImage = true;
            this.colorRowsUniqueButton.SuperTip = "Color cells in selected range based on unique values in unique colors (~2000 colo" +
    "rs and then thay will repeat and may end up next to each other)";
            this.colorRowsUniqueButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.colorCellsUniqueButton_Click);
            // 
            // colorRowsWithTextSplitButton
            // 
            this.colorRowsWithTextSplitButton.Items.Add(this.colorColumnsWithTextButton);
            this.colorRowsWithTextSplitButton.Items.Add(this.colorCellsWithTextButton);
            this.colorRowsWithTextSplitButton.Items.Add(this.colorRowsButton);
            this.colorRowsWithTextSplitButton.Label = "Color rows with text";
            this.colorRowsWithTextSplitButton.Name = "colorRowsWithTextSplitButton";
            this.colorRowsWithTextSplitButton.OfficeImageId = "ColorPickerTable";
            this.colorRowsWithTextSplitButton.ScreenTip = "Color rows with text";
            this.colorRowsWithTextSplitButton.SuperTip = "Select rectangular range and it will show text input dialog that will color rows " +
    "with that contains it";
            this.colorRowsWithTextSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.colorRowsWithTextSplitButton_Click);
            // 
            // colorColumnsWithTextButton
            // 
            this.colorColumnsWithTextButton.Label = "Color columns with text";
            this.colorColumnsWithTextButton.Name = "colorColumnsWithTextButton";
            this.colorColumnsWithTextButton.OfficeImageId = "ColorPickerTable";
            this.colorColumnsWithTextButton.ScreenTip = "Color columns with text";
            this.colorColumnsWithTextButton.ShowImage = true;
            this.colorColumnsWithTextButton.SuperTip = "Select rectangular range and it will show text input dialog that will color colum" +
    "ns with that contains it";
            this.colorColumnsWithTextButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.colorColumnsWithTextButton_Click);
            // 
            // colorCellsWithTextButton
            // 
            this.colorCellsWithTextButton.Label = "Color cells with text";
            this.colorCellsWithTextButton.Name = "colorCellsWithTextButton";
            this.colorCellsWithTextButton.OfficeImageId = "GroupResourceGraphFormat";
            this.colorCellsWithTextButton.ScreenTip = "Color cells with text";
            this.colorCellsWithTextButton.ShowImage = true;
            this.colorCellsWithTextButton.SuperTip = "Select rectangular range and it will show text input dialog that will color cells" +
    " with that contains it";
            this.colorCellsWithTextButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.colorCellsWithTextButton_Click);
            // 
            // colorRowsButton
            // 
            this.colorRowsButton.Label = "Color rows";
            this.colorRowsButton.Name = "colorRowsButton";
            this.colorRowsButton.OfficeImageId = "GroupNetworkDiagramFormat";
            this.colorRowsButton.ScreenTip = "Color rows";
            this.colorRowsButton.ShowImage = true;
            this.colorRowsButton.SuperTip = "Select rectangular range and it will color rows in random colors";
            this.colorRowsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.colorRowsButton_Click);
            // 
            // formatTrueFalseButton
            // 
            this.formatTrueFalseButton.Label = "Format TRUE/FALSE";
            this.formatTrueFalseButton.Name = "formatTrueFalseButton";
            this.formatTrueFalseButton.OfficeImageId = "DataValidation";
            this.formatTrueFalseButton.ScreenTip = "Format cells for TRUE/FALSE";
            this.formatTrueFalseButton.ShowImage = true;
            this.formatTrueFalseButton.SuperTip = resources.GetString("formatTrueFalseButton.SuperTip");
            this.formatTrueFalseButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.formatTrueFalseButton_Click);
            // 
            // searchGroup
            // 
            this.searchGroup.Items.Add(this.searchDialogButton);
            this.searchGroup.Label = "Search";
            this.searchGroup.Name = "searchGroup";
            // 
            // searchDialogButton
            // 
            this.searchDialogButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.searchDialogButton.Image = global::ExcelEssentials.Properties.Resources.table_lookup_512;
            this.searchDialogButton.Label = "Search dialog";
            this.searchDialogButton.Name = "searchDialogButton";
            this.searchDialogButton.OfficeImageId = "DrawingExplorer";
            this.searchDialogButton.ScreenTip = "Search dialog";
            this.searchDialogButton.ShowImage = true;
            this.searchDialogButton.SuperTip = "Will show column search form that helps find columns faster, clear them and find " +
    "columns with specific text";
            this.searchDialogButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.searchDialogButton_Click);
            // 
            // fileAndExportGroup
            // 
            this.fileAndExportGroup.Items.Add(this.saveSelectedWorksheetsAsXlsxSplitBtn);
            this.fileAndExportGroup.Items.Add(this.duplicateWorksheetsSplitBtn);
            this.fileAndExportGroup.Items.Add(this.saveThisWorksheetAsTxt);
            this.fileAndExportGroup.Items.Add(this.divideTableToPartsAndSaveButton);
            this.fileAndExportGroup.Items.Add(this.getFilePathButton);
            this.fileAndExportGroup.Items.Add(this.exportMacrosButton);
            this.fileAndExportGroup.Items.Add(this.separator2);
            this.fileAndExportGroup.Items.Add(this.deleteWorksheetButton);
            this.fileAndExportGroup.Items.Add(this.deleteOtherWorksheetsButton);
            this.fileAndExportGroup.Items.Add(this.deleteWorkbookButton);
            this.fileAndExportGroup.Label = "File && Export";
            this.fileAndExportGroup.Name = "fileAndExportGroup";
            // 
            // saveSelectedWorksheetsAsXlsxSplitBtn
            // 
            this.saveSelectedWorksheetsAsXlsxSplitBtn.Items.Add(this.saveAllWorksheetsAsXlsxButton);
            this.saveSelectedWorksheetsAsXlsxSplitBtn.Items.Add(this.saveSelectedWorksheetsAsTxtButton);
            this.saveSelectedWorksheetsAsXlsxSplitBtn.Items.Add(this.saveAllWorksheetsAsTxtButton);
            this.saveSelectedWorksheetsAsXlsxSplitBtn.Label = "Save sel sheets as xlsx";
            this.saveSelectedWorksheetsAsXlsxSplitBtn.Name = "saveSelectedWorksheetsAsXlsxSplitBtn";
            this.saveSelectedWorksheetsAsXlsxSplitBtn.OfficeImageId = "OrgChartSubordinatesExpand";
            this.saveSelectedWorksheetsAsXlsxSplitBtn.ScreenTip = "Save selected worksheets as xlsx";
            this.saveSelectedWorksheetsAsXlsxSplitBtn.SuperTip = "Saves selected worksheets from workbook to separate xlsx files. It will show dial" +
    "og where to save them.";
            this.saveSelectedWorksheetsAsXlsxSplitBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveSelectedWorksheetsAsXlsxSplitBtn_Click);
            // 
            // saveAllWorksheetsAsXlsxButton
            // 
            this.saveAllWorksheetsAsXlsxButton.Label = "Save all sheets as xlsx";
            this.saveAllWorksheetsAsXlsxButton.Name = "saveAllWorksheetsAsXlsxButton";
            this.saveAllWorksheetsAsXlsxButton.OfficeImageId = "OrgChartSubordinatesExpand";
            this.saveAllWorksheetsAsXlsxButton.ScreenTip = "Save all worksheets as xlsx";
            this.saveAllWorksheetsAsXlsxButton.ShowImage = true;
            this.saveAllWorksheetsAsXlsxButton.SuperTip = "Saves all worksheets from workbook to separate txt files. It will show dialog whe" +
    "re to save them";
            this.saveAllWorksheetsAsXlsxButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveAllWorksheetsAsXlsxButton_Click);
            // 
            // saveSelectedWorksheetsAsTxtButton
            // 
            this.saveSelectedWorksheetsAsTxtButton.Label = "Save selected sheets as txt";
            this.saveSelectedWorksheetsAsTxtButton.Name = "saveSelectedWorksheetsAsTxtButton";
            this.saveSelectedWorksheetsAsTxtButton.OfficeImageId = "OrgChartHorizontalGallery";
            this.saveSelectedWorksheetsAsTxtButton.ScreenTip = "Save selected worksheets as txt";
            this.saveSelectedWorksheetsAsTxtButton.ShowImage = true;
            this.saveSelectedWorksheetsAsTxtButton.SuperTip = resources.GetString("saveSelectedWorksheetsAsTxtButton.SuperTip");
            this.saveSelectedWorksheetsAsTxtButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveSelectedWorksheetsAsTxtButton_Click);
            // 
            // saveAllWorksheetsAsTxtButton
            // 
            this.saveAllWorksheetsAsTxtButton.Label = "Save all sheets as txt";
            this.saveAllWorksheetsAsTxtButton.Name = "saveAllWorksheetsAsTxtButton";
            this.saveAllWorksheetsAsTxtButton.OfficeImageId = "OrgChartHorizontalGallery";
            this.saveAllWorksheetsAsTxtButton.ScreenTip = "Save all worksheets as txt";
            this.saveAllWorksheetsAsTxtButton.ShowImage = true;
            this.saveAllWorksheetsAsTxtButton.SuperTip = resources.GetString("saveAllWorksheetsAsTxtButton.SuperTip");
            this.saveAllWorksheetsAsTxtButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveAllWorksheetsAsTxtButton_Click);
            // 
            // duplicateWorksheetsSplitBtn
            // 
            this.duplicateWorksheetsSplitBtn.Items.Add(this.duplicateWorksheetsToNewWorkbookBtn);
            this.duplicateWorksheetsSplitBtn.Items.Add(this.duplicateWorkbookBtn);
            this.duplicateWorksheetsSplitBtn.Label = "Duplicate sel sheets";
            this.duplicateWorksheetsSplitBtn.Name = "duplicateWorksheetsSplitBtn";
            this.duplicateWorksheetsSplitBtn.OfficeImageId = "DuplicateSelectedSlides";
            this.duplicateWorksheetsSplitBtn.ScreenTip = "Duplicate selected worksheets";
            this.duplicateWorksheetsSplitBtn.SuperTip = "Will duplicate selected worksheets and insert to the right";
            this.duplicateWorksheetsSplitBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.duplicateWorksheetsSplitBtn_Click);
            // 
            // duplicateWorksheetsToNewWorkbookBtn
            // 
            this.duplicateWorksheetsToNewWorkbookBtn.Label = "Duplicate sheets to new wb";
            this.duplicateWorksheetsToNewWorkbookBtn.Name = "duplicateWorksheetsToNewWorkbookBtn";
            this.duplicateWorksheetsToNewWorkbookBtn.OfficeImageId = "CopyAllRules";
            this.duplicateWorksheetsToNewWorkbookBtn.ScreenTip = "Duplicate worksheets to new wb";
            this.duplicateWorksheetsToNewWorkbookBtn.ShowImage = true;
            this.duplicateWorksheetsToNewWorkbookBtn.SuperTip = "Will copy selected worksheets to new workbook first creating and saving that new " +
    "workbook";
            this.duplicateWorksheetsToNewWorkbookBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.duplicateWorksheetsToNewWorkbookBtn_Click);
            // 
            // duplicateWorkbookBtn
            // 
            this.duplicateWorkbookBtn.Label = "Duplicate workbook";
            this.duplicateWorkbookBtn.Name = "duplicateWorkbookBtn";
            this.duplicateWorkbookBtn.OfficeImageId = "Copy";
            this.duplicateWorkbookBtn.ScreenTip = "Duplicate workbook";
            this.duplicateWorkbookBtn.ShowImage = true;
            this.duplicateWorkbookBtn.SuperTip = "Will duplicate whole workbook to new one and ask about saving";
            this.duplicateWorkbookBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.duplicateWorkbookBtn_Click);
            // 
            // saveThisWorksheetAsTxt
            // 
            this.saveThisWorksheetAsTxt.Label = "Save this sheet as txt";
            this.saveThisWorksheetAsTxt.Name = "saveThisWorksheetAsTxt";
            this.saveThisWorksheetAsTxt.OfficeImageId = "ExportTextFile";
            this.saveThisWorksheetAsTxt.ScreenTip = "Save this sheet as txt";
            this.saveThisWorksheetAsTxt.ShowImage = true;
            this.saveThisWorksheetAsTxt.SuperTip = "Saves current sheet to tab delimited txt file and ask about saving. If any cell c" +
    "ontains tab or new line it will warn about it and offer to remove it inserting s" +
    "pace if between words.";
            this.saveThisWorksheetAsTxt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveThisWorksheetAsTxt_Click);
            // 
            // divideTableToPartsAndSaveButton
            // 
            this.divideTableToPartsAndSaveButton.Label = "Divide file";
            this.divideTableToPartsAndSaveButton.Name = "divideTableToPartsAndSaveButton";
            this.divideTableToPartsAndSaveButton.OfficeImageId = "ExportLotus";
            this.divideTableToPartsAndSaveButton.ScreenTip = "Divide file";
            this.divideTableToPartsAndSaveButton.ShowImage = true;
            this.divideTableToPartsAndSaveButton.SuperTip = "Divides table into parts and paste to new worksheets";
            this.divideTableToPartsAndSaveButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.divideTableToPartsAndSaveButton_Click);
            // 
            // getFilePathButton
            // 
            this.getFilePathButton.Label = "Get file path";
            this.getFilePathButton.Name = "getFilePathButton";
            this.getFilePathButton.OfficeImageId = "OpenAttach";
            this.getFilePathButton.ScreenTip = "Get file path of current workbook";
            this.getFilePathButton.ShowImage = true;
            this.getFilePathButton.SuperTip = "Displays form with path to current workbook and offers to copy path to it.";
            this.getFilePathButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.getFilePathButton_Click);
            // 
            // exportMacrosButton
            // 
            this.exportMacrosButton.Label = "Export macros";
            this.exportMacrosButton.Name = "exportMacrosButton";
            this.exportMacrosButton.OfficeImageId = "FileMenuPublishHeader";
            this.exportMacrosButton.ScreenTip = "Export macros from choosen workbook";
            this.exportMacrosButton.ShowImage = true;
            this.exportMacrosButton.SuperTip = "It will export macros from choosen workbook to new folder in Downloads folder.";
            this.exportMacrosButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.exportMacrosButton_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // deleteWorksheetButton
            // 
            this.deleteWorksheetButton.Label = "Delete sel sheets";
            this.deleteWorksheetButton.Name = "deleteWorksheetButton";
            this.deleteWorksheetButton.OfficeImageId = "SheetDelete";
            this.deleteWorksheetButton.ScreenTip = "Delete selected worksheets";
            this.deleteWorksheetButton.ShowImage = true;
            this.deleteWorksheetButton.SuperTip = "Deletes selected worksheets from workbook without confirmation";
            this.deleteWorksheetButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.deleteWorksheetButton_Click);
            // 
            // deleteOtherWorksheetsButton
            // 
            this.deleteOtherWorksheetsButton.Label = "Delete other sheets";
            this.deleteOtherWorksheetsButton.Name = "deleteOtherWorksheetsButton";
            this.deleteOtherWorksheetsButton.OfficeImageId = "DeletePagePreviousVersion";
            this.deleteOtherWorksheetsButton.ScreenTip = "Delete other worksheets";
            this.deleteOtherWorksheetsButton.ShowImage = true;
            this.deleteOtherWorksheetsButton.SuperTip = "Deletes other worksheets than selected with a warning";
            this.deleteOtherWorksheetsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.deleteOtherWorksheetsButton_Click);
            // 
            // deleteWorkbookButton
            // 
            this.deleteWorkbookButton.Label = "Delete this wb";
            this.deleteWorkbookButton.Name = "deleteWorkbookButton";
            this.deleteWorkbookButton.OfficeImageId = "DeleteAll";
            this.deleteWorkbookButton.ScreenTip = "Delete this workbook";
            this.deleteWorkbookButton.ShowImage = true;
            this.deleteWorkbookButton.SuperTip = "Deletes this workbook with warning";
            this.deleteWorkbookButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.deleteWorkbookButton_Click);
            // 
            // macroGroup
            // 
            this.macroGroup.Items.Add(this.runMacroButton);
            this.macroGroup.Items.Add(this.runCustomFormButton);
            this.macroGroup.Label = "Macro";
            this.macroGroup.Name = "macroGroup";
            // 
            // runMacroButton
            // 
            this.runMacroButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.runMacroButton.Label = "Run Macro Form";
            this.runMacroButton.Name = "runMacroButton";
            this.runMacroButton.OfficeImageId = "PlayMacro";
            this.runMacroButton.ScreenTip = "Run Macro Form";
            this.runMacroButton.ShowImage = true;
            this.runMacroButton.SuperTip = "More user friendly macro runner and searcher. (To use you need to have in Excel t" +
    "rust settings trust VBA project object model.)";
            this.runMacroButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.runMacroButton_Click);
            // 
            // runCustomFormButton
            // 
            this.runCustomFormButton.Label = "Run Custom Form";
            this.runCustomFormButton.Name = "runCustomFormButton";
            this.runCustomFormButton.OfficeImageId = "MacroRun";
            this.runCustomFormButton.ScreenTip = "Run Custom Form";
            this.runCustomFormButton.ShowImage = true;
            this.runCustomFormButton.SuperTip = resources.GetString("runCustomFormButton.SuperTip");
            this.runCustomFormButton.Visible = false;
            this.runCustomFormButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.runCustomFormButton_Click);
            // 
            // pivotToolsTab
            // 
            this.pivotToolsTab.Groups.Add(this.pivotTemplatesGroup);
            this.pivotToolsTab.Groups.Add(this.pivotFormatGroup);
            this.pivotToolsTab.Groups.Add(this.pivotToolsGroup);
            this.pivotToolsTab.Label = "Pivot\'s tools";
            this.pivotToolsTab.Name = "pivotToolsTab";
            // 
            // pivotTemplatesGroup
            // 
            this.pivotTemplatesGroup.Items.Add(this.createPivotFromTemplateMenu);
            this.pivotTemplatesGroup.Items.Add(this.generatePivotTemlateCodeButton);
            this.pivotTemplatesGroup.Items.Add(this.runPvTemplateBtn);
            this.pivotTemplatesGroup.Label = "Pivot templates";
            this.pivotTemplatesGroup.Name = "pivotTemplatesGroup";
            // 
            // createPivotFromTemplateMenu
            // 
            this.createPivotFromTemplateMenu.Items.Add(this.createPvFromLoadedButton);
            this.createPivotFromTemplateMenu.Items.Add(this.separator3);
            this.createPivotFromTemplateMenu.Items.Add(this.createPvFromInboundButton);
            this.createPivotFromTemplateMenu.Items.Add(this.createPvFromInboundPlusButton);
            this.createPivotFromTemplateMenu.Items.Add(this.separator4);
            this.createPivotFromTemplateMenu.Items.Add(this.createPvFromOutboundButton);
            this.createPivotFromTemplateMenu.Items.Add(this.createPvFromOutboundPlusButton);
            this.createPivotFromTemplateMenu.Items.Add(this.separator5);
            this.createPivotFromTemplateMenu.Items.Add(this.createPvFromMNTXButton);
            this.createPivotFromTemplateMenu.Items.Add(this.separator6);
            this.createPivotFromTemplateMenu.Items.Add(this.createPvFromYottaButton);
            this.createPivotFromTemplateMenu.Items.Add(this.separator7);
            this.createPivotFromTemplateMenu.Items.Add(this.createPvFromSAPCButton);
            this.createPivotFromTemplateMenu.Items.Add(this.createPvFromSAPCLoadedButton);
            this.createPivotFromTemplateMenu.Items.Add(this.separator8);
            this.createPivotFromTemplateMenu.Items.Add(this.createPvFromVBAKButton);
            this.createPivotFromTemplateMenu.Items.Add(this.createPvFromWBRKButton);
            this.createPivotFromTemplateMenu.Items.Add(this.createPvFromVBPAButton);
            this.createPivotFromTemplateMenu.Items.Add(this.createPvFromVBAPButton);
            this.createPivotFromTemplateMenu.Items.Add(this.createPvFromVBPAandKNA1Button);
            this.createPivotFromTemplateMenu.Items.Add(this.separator9);
            this.createPivotFromTemplateMenu.Items.Add(this.createPvFromQlikButton);
            this.createPivotFromTemplateMenu.Items.Add(this.separator10);
            this.createPivotFromTemplateMenu.Items.Add(this.customPivotsTemplatesMenu);
            this.createPivotFromTemplateMenu.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPivotFromTemplateMenu.Label = "Create pivot from";
            this.createPivotFromTemplateMenu.Name = "createPivotFromTemplateMenu";
            this.createPivotFromTemplateMenu.OfficeImageId = "PivotTableNewStyle";
            this.createPivotFromTemplateMenu.ScreenTip = "Create pivot from...";
            this.createPivotFromTemplateMenu.ShowImage = true;
            this.createPivotFromTemplateMenu.SuperTip = "Will use macro from default macro workbook to create pivot table from template us" +
    "ing macro that starts with \"CreatePvFrom\" nad end with button name";
            // 
            // createPvFromLoadedButton
            // 
            this.createPvFromLoadedButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPvFromLoadedButton.Label = "Loaded";
            this.createPvFromLoadedButton.Name = "createPvFromLoadedButton";
            this.createPvFromLoadedButton.OfficeImageId = "AppointmentColor3";
            this.createPvFromLoadedButton.ShowImage = true;
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // createPvFromInboundButton
            // 
            this.createPvFromInboundButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPvFromInboundButton.Label = "Inbound";
            this.createPvFromInboundButton.Name = "createPvFromInboundButton";
            this.createPvFromInboundButton.OfficeImageId = "AppointmentColor10";
            this.createPvFromInboundButton.ShowImage = true;
            // 
            // createPvFromInboundPlusButton
            // 
            this.createPvFromInboundPlusButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPvFromInboundPlusButton.Label = "Inbound+";
            this.createPvFromInboundPlusButton.Name = "createPvFromInboundPlusButton";
            this.createPvFromInboundPlusButton.OfficeImageId = "AppointmentColor10";
            this.createPvFromInboundPlusButton.ShowImage = true;
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // createPvFromOutboundButton
            // 
            this.createPvFromOutboundButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPvFromOutboundButton.Label = "Outbound";
            this.createPvFromOutboundButton.Name = "createPvFromOutboundButton";
            this.createPvFromOutboundButton.OfficeImageId = "AppointmentColor0";
            this.createPvFromOutboundButton.ShowImage = true;
            // 
            // createPvFromOutboundPlusButton
            // 
            this.createPvFromOutboundPlusButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPvFromOutboundPlusButton.Label = "Outbound+";
            this.createPvFromOutboundPlusButton.Name = "createPvFromOutboundPlusButton";
            this.createPvFromOutboundPlusButton.OfficeImageId = "AppointmentColor0";
            this.createPvFromOutboundPlusButton.ShowImage = true;
            // 
            // separator5
            // 
            this.separator5.Name = "separator5";
            // 
            // createPvFromMNTXButton
            // 
            this.createPvFromMNTXButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPvFromMNTXButton.Label = "MNTX";
            this.createPvFromMNTXButton.Name = "createPvFromMNTXButton";
            this.createPvFromMNTXButton.OfficeImageId = "AppointmentColor1";
            this.createPvFromMNTXButton.ShowImage = true;
            // 
            // separator6
            // 
            this.separator6.Name = "separator6";
            // 
            // createPvFromYottaButton
            // 
            this.createPvFromYottaButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPvFromYottaButton.Label = "Yotta";
            this.createPvFromYottaButton.Name = "createPvFromYottaButton";
            this.createPvFromYottaButton.OfficeImageId = "AppointmentBusy";
            this.createPvFromYottaButton.ShowImage = true;
            // 
            // separator7
            // 
            this.separator7.Name = "separator7";
            // 
            // createPvFromSAPCButton
            // 
            this.createPvFromSAPCButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPvFromSAPCButton.Label = "SAP C (webi)";
            this.createPvFromSAPCButton.Name = "createPvFromSAPCButton";
            this.createPvFromSAPCButton.OfficeImageId = "AppointmentColor2";
            this.createPvFromSAPCButton.ShowImage = true;
            // 
            // createPvFromSAPCLoadedButton
            // 
            this.createPvFromSAPCLoadedButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPvFromSAPCLoadedButton.Label = "SAP C (SQL)";
            this.createPvFromSAPCLoadedButton.Name = "createPvFromSAPCLoadedButton";
            this.createPvFromSAPCLoadedButton.OfficeImageId = "AppointmentColor2";
            this.createPvFromSAPCLoadedButton.ShowImage = true;
            // 
            // separator8
            // 
            this.separator8.Name = "separator8";
            // 
            // createPvFromVBAKButton
            // 
            this.createPvFromVBAKButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPvFromVBAKButton.Label = "VBAK";
            this.createPvFromVBAKButton.Name = "createPvFromVBAKButton";
            this.createPvFromVBAKButton.OfficeImageId = "AppointmentColor4";
            this.createPvFromVBAKButton.ShowImage = true;
            // 
            // createPvFromWBRKButton
            // 
            this.createPvFromWBRKButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPvFromWBRKButton.Label = "WBRK";
            this.createPvFromWBRKButton.Name = "createPvFromWBRKButton";
            this.createPvFromWBRKButton.OfficeImageId = "AppointmentColor5";
            this.createPvFromWBRKButton.ShowImage = true;
            // 
            // createPvFromVBPAButton
            // 
            this.createPvFromVBPAButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPvFromVBPAButton.Label = "VBPA";
            this.createPvFromVBPAButton.Name = "createPvFromVBPAButton";
            this.createPvFromVBPAButton.OfficeImageId = "AppointmentColor6";
            this.createPvFromVBPAButton.ShowImage = true;
            // 
            // createPvFromVBAPButton
            // 
            this.createPvFromVBAPButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPvFromVBAPButton.Label = "VBAP";
            this.createPvFromVBAPButton.Name = "createPvFromVBAPButton";
            this.createPvFromVBAPButton.OfficeImageId = "AppointmentColor4";
            this.createPvFromVBAPButton.ShowImage = true;
            // 
            // createPvFromVBPAandKNA1Button
            // 
            this.createPvFromVBPAandKNA1Button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPvFromVBPAandKNA1Button.Label = "VBPA && KNA1";
            this.createPvFromVBPAandKNA1Button.Name = "createPvFromVBPAandKNA1Button";
            this.createPvFromVBPAandKNA1Button.OfficeImageId = "AppointmentColor6";
            this.createPvFromVBPAandKNA1Button.ShowImage = true;
            // 
            // separator9
            // 
            this.separator9.Name = "separator9";
            // 
            // createPvFromQlikButton
            // 
            this.createPvFromQlikButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.createPvFromQlikButton.Label = "Qlik";
            this.createPvFromQlikButton.Name = "createPvFromQlikButton";
            this.createPvFromQlikButton.OfficeImageId = "AppointmentColor9";
            this.createPvFromQlikButton.ShowImage = true;
            // 
            // separator10
            // 
            this.separator10.Name = "separator10";
            // 
            // customPivotsTemplatesMenu
            // 
            this.customPivotsTemplatesMenu.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.customPivotsTemplatesMenu.Items.Add(this.createPvFromCustom1Button);
            this.customPivotsTemplatesMenu.Items.Add(this.createPvFromCustom2Button);
            this.customPivotsTemplatesMenu.Items.Add(this.createPvFromCustom3Button);
            this.customPivotsTemplatesMenu.Items.Add(this.createPvFromCustom4Button);
            this.customPivotsTemplatesMenu.Items.Add(this.createPvFromCustom5Button);
            this.customPivotsTemplatesMenu.Items.Add(this.createPvFromCustom6Button);
            this.customPivotsTemplatesMenu.Items.Add(this.createPvFromCustom7Button);
            this.customPivotsTemplatesMenu.Items.Add(this.createPvFromCustom8Button);
            this.customPivotsTemplatesMenu.Items.Add(this.createPvFromCustom9Button);
            this.customPivotsTemplatesMenu.Label = "Custom";
            this.customPivotsTemplatesMenu.Name = "customPivotsTemplatesMenu";
            this.customPivotsTemplatesMenu.OfficeImageId = "AppointmentColorDialog";
            this.customPivotsTemplatesMenu.ShowImage = true;
            // 
            // createPvFromCustom1Button
            // 
            this.createPvFromCustom1Button.Label = "Custom1";
            this.createPvFromCustom1Button.Name = "createPvFromCustom1Button";
            this.createPvFromCustom1Button.OfficeImageId = "_1";
            this.createPvFromCustom1Button.ShowImage = true;
            // 
            // createPvFromCustom2Button
            // 
            this.createPvFromCustom2Button.Label = "Custom2";
            this.createPvFromCustom2Button.Name = "createPvFromCustom2Button";
            this.createPvFromCustom2Button.ShowImage = true;
            // 
            // createPvFromCustom3Button
            // 
            this.createPvFromCustom3Button.Label = "Custom3";
            this.createPvFromCustom3Button.Name = "createPvFromCustom3Button";
            this.createPvFromCustom3Button.ShowImage = true;
            // 
            // createPvFromCustom4Button
            // 
            this.createPvFromCustom4Button.Label = "Custom4";
            this.createPvFromCustom4Button.Name = "createPvFromCustom4Button";
            this.createPvFromCustom4Button.ShowImage = true;
            // 
            // createPvFromCustom5Button
            // 
            this.createPvFromCustom5Button.Label = "Custom5";
            this.createPvFromCustom5Button.Name = "createPvFromCustom5Button";
            this.createPvFromCustom5Button.ShowImage = true;
            // 
            // createPvFromCustom6Button
            // 
            this.createPvFromCustom6Button.Label = "Custom6";
            this.createPvFromCustom6Button.Name = "createPvFromCustom6Button";
            this.createPvFromCustom6Button.ShowImage = true;
            // 
            // createPvFromCustom7Button
            // 
            this.createPvFromCustom7Button.Label = "Custom7";
            this.createPvFromCustom7Button.Name = "createPvFromCustom7Button";
            this.createPvFromCustom7Button.ShowImage = true;
            // 
            // createPvFromCustom8Button
            // 
            this.createPvFromCustom8Button.Label = "Custom8";
            this.createPvFromCustom8Button.Name = "createPvFromCustom8Button";
            this.createPvFromCustom8Button.ShowImage = true;
            // 
            // createPvFromCustom9Button
            // 
            this.createPvFromCustom9Button.Label = "Custom9";
            this.createPvFromCustom9Button.Name = "createPvFromCustom9Button";
            this.createPvFromCustom9Button.ShowImage = true;
            // 
            // generatePivotTemlateCodeButton
            // 
            this.generatePivotTemlateCodeButton.Label = "Gen. pv template";
            this.generatePivotTemlateCodeButton.Name = "generatePivotTemlateCodeButton";
            this.generatePivotTemlateCodeButton.OfficeImageId = "PivotTableListFormulas";
            this.generatePivotTemlateCodeButton.ScreenTip = "Generate pv template code from pv";
            this.generatePivotTemlateCodeButton.ShowImage = true;
            this.generatePivotTemlateCodeButton.SuperTip = "First format pivot table and then you can use this to create template from it and" +
    " copy template and paste it manualy into default macro workbook and remember to " +
    "save it";
            this.generatePivotTemlateCodeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.generatePivotTemlateCodeButton_Click);
            // 
            // runPvTemplateBtn
            // 
            this.runPvTemplateBtn.Label = "Run pv template";
            this.runPvTemplateBtn.Name = "runPvTemplateBtn";
            this.runPvTemplateBtn.OfficeImageId = "PivotTableNewStyle";
            this.runPvTemplateBtn.ScreenTip = "Run pv template form";
            this.runPvTemplateBtn.ShowImage = true;
            this.runPvTemplateBtn.SuperTip = resources.GetString("runPvTemplateBtn.SuperTip");
            this.runPvTemplateBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.runPvTemplateButton_Click);
            // 
            // pivotFormatGroup
            // 
            this.pivotFormatGroup.Items.Add(this.formatClickedPivotButton);
            this.pivotFormatGroup.Items.Add(this.formatAllPivotButton);
            this.pivotFormatGroup.Items.Add(this.grandTotalsToggleButton);
            this.pivotFormatGroup.Items.Add(this.subtotalsToggleButton);
            this.pivotFormatGroup.Label = "Format pivot";
            this.pivotFormatGroup.Name = "pivotFormatGroup";
            // 
            // formatClickedPivotButton
            // 
            this.formatClickedPivotButton.Label = "Format current pv";
            this.formatClickedPivotButton.Name = "formatClickedPivotButton";
            this.formatClickedPivotButton.OfficeImageId = "GroupQueries";
            this.formatClickedPivotButton.ScreenTip = "Format current pv";
            this.formatClickedPivotButton.ShowImage = true;
            this.formatClickedPivotButton.SuperTip = "Format selected pivot table";
            this.formatClickedPivotButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.formatClickedPivotButton_Click);
            // 
            // formatAllPivotButton
            // 
            this.formatAllPivotButton.Label = "Format all pv";
            this.formatAllPivotButton.Name = "formatAllPivotButton";
            this.formatAllPivotButton.OfficeImageId = "QuerySelectQueryType";
            this.formatAllPivotButton.ScreenTip = "Format all pv";
            this.formatAllPivotButton.ShowImage = true;
            this.formatAllPivotButton.SuperTip = "Format all pivot tables in current sheet";
            this.formatAllPivotButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.formatAllPivotButton_Click);
            // 
            // grandTotalsToggleButton
            // 
            this.grandTotalsToggleButton.Label = "Grand Totals";
            this.grandTotalsToggleButton.Name = "grandTotalsToggleButton";
            this.grandTotalsToggleButton.OfficeImageId = "GroupCalculation";
            this.grandTotalsToggleButton.ScreenTip = "Grand Totals switch for pivot table";
            this.grandTotalsToggleButton.ShowImage = true;
            this.grandTotalsToggleButton.SuperTip = "Toggle grand totals in selected pivot table (click one more time if nothing chang" +
    "ed)";
            this.grandTotalsToggleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.grandTotalsToggleButton_Click);
            // 
            // subtotalsToggleButton
            // 
            this.subtotalsToggleButton.Label = "Subtotals";
            this.subtotalsToggleButton.Name = "subtotalsToggleButton";
            this.subtotalsToggleButton.OfficeImageId = "GroupCalculation";
            this.subtotalsToggleButton.ScreenTip = "Subtotals switch for pv";
            this.subtotalsToggleButton.ShowImage = true;
            this.subtotalsToggleButton.SuperTip = "Toggle subtotals in selected pivot table (click one more time if nothing changed)" +
    "";
            this.subtotalsToggleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.subtotalsToggleButton_Click);
            // 
            // pivotToolsGroup
            // 
            this.pivotToolsGroup.Items.Add(this.changePivotTableSourceButton);
            this.pivotToolsGroup.Items.Add(this.updatePivotTableSourceButton);
            this.pivotToolsGroup.Items.Add(this.refreshPivotsButton);
            this.pivotToolsGroup.Items.Add(this.combinedTableFromPvValuesButton);
            this.pivotToolsGroup.Label = "Pivot tools";
            this.pivotToolsGroup.Name = "pivotToolsGroup";
            // 
            // changePivotTableSourceButton
            // 
            this.changePivotTableSourceButton.Label = "Change pv source";
            this.changePivotTableSourceButton.Name = "changePivotTableSourceButton";
            this.changePivotTableSourceButton.OfficeImageId = "PivotShowDetails";
            this.changePivotTableSourceButton.ScreenTip = "Change pivot table source";
            this.changePivotTableSourceButton.ShowImage = true;
            this.changePivotTableSourceButton.SuperTip = "Will show you range picker and select table or click on one cell in that table to" +
    " change pivot table source of current pivot table to choosen one";
            this.changePivotTableSourceButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.changePivotTableSourceButton_Click);
            // 
            // updatePivotTableSourceButton
            // 
            this.updatePivotTableSourceButton.Label = "Update pv source";
            this.updatePivotTableSourceButton.Name = "updatePivotTableSourceButton";
            this.updatePivotTableSourceButton.OfficeImageId = "PivotShowDetails";
            this.updatePivotTableSourceButton.ScreenTip = "Update pivot table source";
            this.updatePivotTableSourceButton.ShowImage = true;
            this.updatePivotTableSourceButton.SuperTip = "Will update pivot table source to current region of current pivot table";
            this.updatePivotTableSourceButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.updatePivotTableSourceButton_Click);
            // 
            // refreshPivotsButton
            // 
            this.refreshPivotsButton.Label = "Refresh all pv";
            this.refreshPivotsButton.Name = "refreshPivotsButton";
            this.refreshPivotsButton.OfficeImageId = "RefreshWebView";
            this.refreshPivotsButton.ScreenTip = "Refresh all pivot tables";
            this.refreshPivotsButton.ShowImage = true;
            this.refreshPivotsButton.SuperTip = "Refresh all pivot tables in current workbook";
            this.refreshPivotsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.refreshPivotsButton_Click);
            // 
            // combinedTableFromPvValuesButton
            // 
            this.combinedTableFromPvValuesButton.Label = "Table from values";
            this.combinedTableFromPvValuesButton.Name = "combinedTableFromPvValuesButton";
            this.combinedTableFromPvValuesButton.OfficeImageId = "GroupOrganizationChartSelect";
            this.combinedTableFromPvValuesButton.ScreenTip = "Table from values";
            this.combinedTableFromPvValuesButton.ShowImage = true;
            this.combinedTableFromPvValuesButton.SuperTip = "Select values cells in pivot table and it will create combined table from data th" +
    "at would be created by double click";
            this.combinedTableFromPvValuesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.combinedTableFromPvValuesButton_Click);
            // 
            // dataImportTab
            // 
            this.dataImportTab.Groups.Add(this.sqlImportGroup);
            this.dataImportTab.Groups.Add(this.sapImportGroup);
            this.dataImportTab.Groups.Add(this.sdeImportGroup);
            this.dataImportTab.Groups.Add(this.browserGroup);
            this.dataImportTab.Label = "Data import";
            this.dataImportTab.Name = "dataImportTab";
            // 
            // sqlImportGroup
            // 
            this.sqlImportGroup.Items.Add(this.loadToDataTableButton);
            this.sqlImportGroup.Label = "SQL";
            this.sqlImportGroup.Name = "sqlImportGroup";
            this.sqlImportGroup.Visible = false;
            // 
            // loadToDataTableButton
            // 
            this.loadToDataTableButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.loadToDataTableButton.Label = "Load to DataTable";
            this.loadToDataTableButton.Name = "loadToDataTableButton";
            this.loadToDataTableButton.OfficeImageId = "AdpStoredProcedureEditSql";
            this.loadToDataTableButton.ScreenTip = "Load to DataTable";
            this.loadToDataTableButton.ShowImage = true;
            this.loadToDataTableButton.SuperTip = "Will load selected rectangular range to data table (do not select range if you wa" +
    "nt to load file or select range later) and let you SQL query it and paste to exc" +
    "el later";
            this.loadToDataTableButton.Visible = false;
            this.loadToDataTableButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.loadToDataTableButton_Click);
            // 
            // sapImportGroup
            // 
            this.sapImportGroup.Items.Add(this.runS4ExtractButton);
            this.sapImportGroup.Label = "SAP";
            this.sapImportGroup.Name = "sapImportGroup";
            // 
            // runS4ExtractButton
            // 
            this.runS4ExtractButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.runS4ExtractButton.Image = global::ExcelEssentials.Properties.Resources.S4_HANA_cloud;
            this.runS4ExtractButton.Label = "Run SAP Extract";
            this.runS4ExtractButton.Name = "runS4ExtractButton";
            this.runS4ExtractButton.ScreenTip = "Run SAP Extract form";
            this.runS4ExtractButton.ShowImage = true;
            this.runS4ExtractButton.SuperTip = resources.GetString("runS4ExtractButton.SuperTip");
            this.runS4ExtractButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.runS4ExtractButton_Click);
            // 
            // sdeImportGroup
            // 
            this.sdeImportGroup.Items.Add(this.runSdeButton);
            this.sdeImportGroup.Items.Add(this.sdeQueryComboBox);
            this.sdeImportGroup.Items.Add(this.sdeInstancesEditBox);
            this.sdeImportGroup.Label = "SDE";
            this.sdeImportGroup.Name = "sdeImportGroup";
            this.sdeImportGroup.Visible = false;
            // 
            // runSdeButton
            // 
            this.runSdeButton.Image = global::ExcelEssentials.Properties.Resources.rocketLogo_scale_400;
            this.runSdeButton.Label = "Run SDE";
            this.runSdeButton.Name = "runSdeButton";
            this.runSdeButton.ScreenTip = "Run SDE Launcher";
            this.runSdeButton.ShowImage = true;
            this.runSdeButton.SuperTip = "It will run new SDE Launcher v2.x with selected query selected values as range in" +
    "to first filter editor for that query with given instances (path to SDE Launcher" +
    " v2.x need to be specified)";
            this.runSdeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.runSdeButton_Click);
            // 
            // sdeQueryComboBox
            // 
            this.sdeQueryComboBox.Label = "Query";
            this.sdeQueryComboBox.Name = "sdeQueryComboBox";
            this.sdeQueryComboBox.SuperTip = "Choose SDE Lanucher available queries";
            this.sdeQueryComboBox.Text = null;
            // 
            // sdeInstancesEditBox
            // 
            this.sdeInstancesEditBox.Label = "Instances";
            this.sdeInstancesEditBox.MaxLength = 2;
            this.sdeInstancesEditBox.Name = "sdeInstancesEditBox";
            this.sdeInstancesEditBox.SuperTip = "Specify number of instances for SDE Laucher";
            this.sdeInstancesEditBox.Text = "1";
            this.sdeInstancesEditBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sdeInstancesEditBox_TextChanged);
            // 
            // browserGroup
            // 
            this.browserGroup.Items.Add(this.browserButton);
            this.browserGroup.Items.Add(this.browserWebsitesComboBox);
            this.browserGroup.Items.Add(this.importFromBrowserCheckBox);
            this.browserGroup.Label = "Browser";
            this.browserGroup.Name = "browserGroup";
            // 
            // browserButton
            // 
            this.browserButton.Image = global::ExcelEssentials.Properties.Resources.Microsoft_Edge_logo__2019__svg;
            this.browserButton.Label = "Launch Browser";
            this.browserButton.Name = "browserButton";
            this.browserButton.ScreenTip = "Launch Browser";
            this.browserButton.ShowImage = true;
            this.browserButton.SuperTip = resources.GetString("browserButton.SuperTip");
            this.browserButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.browserButton_Click);
            // 
            // browserWebsitesComboBox
            // 
            this.browserWebsitesComboBox.Label = "Website";
            this.browserWebsitesComboBox.Name = "browserWebsitesComboBox";
            this.browserWebsitesComboBox.OfficeImageId = "WebBrowserControl";
            this.browserWebsitesComboBox.ShowImage = true;
            this.browserWebsitesComboBox.SuperTip = "Choose website that will be used when browser opens";
            this.browserWebsitesComboBox.Text = "(blank)";
            // 
            // importFromBrowserCheckBox
            // 
            this.importFromBrowserCheckBox.Checked = true;
            this.importFromBrowserCheckBox.Label = "Auto import";
            this.importFromBrowserCheckBox.Name = "importFromBrowserCheckBox";
            this.importFromBrowserCheckBox.ScreenTip = "Auto import from browser toggle";
            this.importFromBrowserCheckBox.SuperTip = "Auto import downloaded Excel files and txt/csv files";
            // 
            // goToPropertiesButton
            // 
            this.goToPropertiesButton.Label = "Properties files";
            this.goToPropertiesButton.Name = "goToPropertiesButton";
            this.goToPropertiesButton.OfficeImageId = "ProjectManageDeliverables";
            this.goToPropertiesButton.ScreenTip = "Properties files";
            this.goToPropertiesButton.ShowImage = true;
            this.goToPropertiesButton.SuperTip = "Will open location of properties files of ExcelEssencialsPack Add-in";
            this.goToPropertiesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.goToPropertiesButton_Click);
            // 
            // updateMacrosButton
            // 
            this.updateMacrosButton.Label = "Update macros";
            this.updateMacrosButton.Name = "updateMacrosButton";
            this.updateMacrosButton.OfficeImageId = "PublishWorkflow";
            this.updateMacrosButton.ScreenTip = "Update macros";
            this.updateMacrosButton.ShowImage = true;
            this.updateMacrosButton.SuperTip = resources.GetString("updateMacrosButton.SuperTip");
            this.updateMacrosButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.updateMacrosButton_Click);
            // 
            // createMacroUpdateButton
            // 
            this.createMacroUpdateButton.Label = "Create macro update";
            this.createMacroUpdateButton.Name = "createMacroUpdateButton";
            this.createMacroUpdateButton.OfficeImageId = "LogicShowDialogBoxAction";
            this.createMacroUpdateButton.ScreenTip = "Create macro update file in Downloads";
            this.createMacroUpdateButton.ShowImage = true;
            this.createMacroUpdateButton.SuperTip = "Will create specific to custom ribbon Update macros file with clicked macro from " +
    "VBA editor (has to be open and macro clicked) and it will be pasted to Downloads" +
    " folder in macros backup folder";
            this.createMacroUpdateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.createMacroUpdateButton_Click);
            // 
            // checkMacrosButton
            // 
            this.checkMacrosButton.Label = "Check macros";
            this.checkMacrosButton.Name = "checkMacrosButton";
            this.checkMacrosButton.OfficeImageId = "MacroDefault";
            this.checkMacrosButton.ScreenTip = "Check macros";
            this.checkMacrosButton.ShowImage = true;
            this.checkMacrosButton.SuperTip = "It will check if all macros that are assigned to buttons are present in specified" +
    " workbook from Properties Files folder.";
            this.checkMacrosButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkMacrosButton_Click);
            // 
            // excelEssentialsPackInfoBtn
            // 
            this.excelEssentialsPackInfoBtn.Label = "ExcelEssentialsPack info";
            this.excelEssentialsPackInfoBtn.Name = "excelEssentialsPackInfoBtn";
            this.excelEssentialsPackInfoBtn.ShowImage = true;
            this.excelEssentialsPackInfoBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.excelEssentialsPackInfoBtn_Click);
            // 
            // MiscRibbon
            // 
            this.Name = "MiscRibbon";
            // 
            // MiscRibbon.OfficeMenu
            // 
            this.OfficeMenu.Items.Add(this.goToPropertiesButton);
            this.OfficeMenu.Items.Add(this.updateMacrosButton);
            this.OfficeMenu.Items.Add(this.createMacroUpdateButton);
            this.OfficeMenu.Items.Add(this.checkMacrosButton);
            this.OfficeMenu.Items.Add(this.excelEssentialsPackInfoBtn);
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.miscTab);
            this.Tabs.Add(this.pivotToolsTab);
            this.Tabs.Add(this.dataImportTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MiscRibbon_Load);
            this.miscTab.ResumeLayout(false);
            this.miscTab.PerformLayout();
            this.importGroup.ResumeLayout(false);
            this.importGroup.PerformLayout();
            this.modifiersGroup.ResumeLayout(false);
            this.modifiersGroup.PerformLayout();
            this.filterGroup.ResumeLayout(false);
            this.filterGroup.PerformLayout();
            this.validationGroup.ResumeLayout(false);
            this.validationGroup.PerformLayout();
            this.searchGroup.ResumeLayout(false);
            this.searchGroup.PerformLayout();
            this.fileAndExportGroup.ResumeLayout(false);
            this.fileAndExportGroup.PerformLayout();
            this.macroGroup.ResumeLayout(false);
            this.macroGroup.PerformLayout();
            this.pivotToolsTab.ResumeLayout(false);
            this.pivotToolsTab.PerformLayout();
            this.pivotTemplatesGroup.ResumeLayout(false);
            this.pivotTemplatesGroup.PerformLayout();
            this.pivotFormatGroup.ResumeLayout(false);
            this.pivotFormatGroup.PerformLayout();
            this.pivotToolsGroup.ResumeLayout(false);
            this.pivotToolsGroup.PerformLayout();
            this.dataImportTab.ResumeLayout(false);
            this.dataImportTab.PerformLayout();
            this.sqlImportGroup.ResumeLayout(false);
            this.sqlImportGroup.PerformLayout();
            this.sapImportGroup.ResumeLayout(false);
            this.sapImportGroup.PerformLayout();
            this.sdeImportGroup.ResumeLayout(false);
            this.sdeImportGroup.PerformLayout();
            this.browserGroup.ResumeLayout(false);
            this.browserGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup modifiersGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton appendTextButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton changeToTextButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton changeToValueButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton evaluateFormulaButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton repasteAsValuesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton removeEmptyButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton removeNaButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton trimButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton formatStringToDateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup filterGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton filterColumnSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton hideColumnsWithTextButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton takeRowsWithTextButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup validationGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton colorRowsUniqueButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton colorRowsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup searchGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton searchDialogButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton formatTrueFalseButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton filterColumnNotInRangeBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton filterColumnFromRangeInRangeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup importGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup fileAndExportGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton getFilePathButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveSelectedWorksheetsAsTxtButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton deleteWorksheetButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton deleteOtherWorksheetsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton deleteWorkbookButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveThisWorksheetAsTxt;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton divideTableToPartsAndSaveButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton sortingAbsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copyAsPictureButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup pivotTemplatesGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton generatePivotTemlateCodeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup pivotFormatGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton formatClickedPivotButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton formatAllPivotButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton refreshPivotsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu createPivotFromTemplateMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromLoadedButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromInboundButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromInboundPlusButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromOutboundButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromOutboundPlusButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromYottaButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromSAPCButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromSAPCLoadedButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromVBAKButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromVBPAButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromWBRKButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromVBAPButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromVBPAandKNA1Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromQlikButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromMNTXButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu customPivotsTemplatesMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromCustom1Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromCustom2Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromCustom3Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromCustom4Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromCustom5Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromCustom6Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromCustom7Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromCustom8Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createPvFromCustom9Button;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator5;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator6;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator7;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator8;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator9;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator10;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup sqlImportGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup sapImportGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup sdeImportGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton saveSelectedWorksheetsAsXlsxSplitBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton duplicateWorksheetsSplitBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton duplicateWorkbookBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton duplicateWorksheetsToNewWorkbookBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton grandTotalsToggleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton filterColumnFromRangeNotInRangeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton runPvTemplateBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton runMacroButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator11;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup macroGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton colorRowsWithTextSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton removeHiddenRowsSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton removeHiddenColumnsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton removeFormattingButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton runS4ExtractButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton runSdeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox sdeQueryComboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox sdeInstancesEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton browserButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup browserGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox browserWebsitesComboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox importFromBrowserCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton clearRangeOutsideButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton importTxtFileLegacyButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton importSheetOrTxtFileSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton subtotalsToggleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loadToDataTableButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton colorColumnsWithTextButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton colorCellsWithTextButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton filterColumnInRegexButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton filterColumnNotInRegexButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton goToPropertiesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton updateMacrosButton;
        public Microsoft.Office.Tools.Ribbon.RibbonTab miscTab;
        public Microsoft.Office.Tools.Ribbon.RibbonTab pivotToolsTab;
        public Microsoft.Office.Tools.Ribbon.RibbonTab dataImportTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup pivotToolsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton combinedTableFromPvValuesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton exportMacrosButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton hideRowsWithTextSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton runCustomFormButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton prependTextSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator13;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton selectWithoutHeadersButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton fillEmptyWithAboveValueButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copyDelimitedValuesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createMacroUpdateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton checkMacrosButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton updatePivotTableSourceButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveAllWorksheetsAsTxtButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveAllWorksheetsAsXlsxButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton removeErrSplitBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton changePivotTableSourceButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton formatNumberSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton importSheetOrTxtFileAdvButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton removeDuplicatesInColumnsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton removeDuplicatesSplitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton excelEssentialsPackInfoBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton filterColumnFlipFilterBtn;
    }

    partial class ThisRibbonCollection
    {
        internal MiscRibbon MiscRibbon
        {
            get { return this.GetRibbon<MiscRibbon>(); }
        }
    }
}
