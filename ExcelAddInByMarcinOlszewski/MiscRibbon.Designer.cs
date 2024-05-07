namespace ExcelAddInByMarcinOlszewski
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
            this.removeNaButton = this.Factory.CreateRibbonButton();
            this.removeDuplicatesButton = this.Factory.CreateRibbonButton();
            this.removeHiddenRowsSplitButton = this.Factory.CreateRibbonSplitButton();
            this.removeHiddenColumnsButton = this.Factory.CreateRibbonButton();
            this.clearRangeOutsideButton = this.Factory.CreateRibbonButton();
            this.removeFormattingButton = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.prependTextButton = this.Factory.CreateRibbonButton();
            this.trimButton = this.Factory.CreateRibbonButton();
            this.formatNumberButton = this.Factory.CreateRibbonButton();
            this.filterGroup = this.Factory.CreateRibbonGroup();
            this.filterColumnSplitButton = this.Factory.CreateRibbonSplitButton();
            this.filterColumnNotInRangeBtn = this.Factory.CreateRibbonButton();
            this.filterColumnFromRangeInRangeButton = this.Factory.CreateRibbonButton();
            this.filterColumnFromRangeNotInRangeButton = this.Factory.CreateRibbonButton();
            this.filterColumnInRegexButton = this.Factory.CreateRibbonButton();
            this.filterColumnNotInRegexButton = this.Factory.CreateRibbonButton();
            this.hideRowsWithTextSplitButton = this.Factory.CreateRibbonSplitButton();
            this.hideColumnsWithTextButton = this.Factory.CreateRibbonButton();
            this.takeRowsWithTextButton = this.Factory.CreateRibbonButton();
            this.validationGroup = this.Factory.CreateRibbonGroup();
            this.colorRowsUniqueButton = this.Factory.CreateRibbonButton();
            this.colorRowsWithTextSplitButton = this.Factory.CreateRibbonSplitButton();
            this.colorColumnsWithTextButton = this.Factory.CreateRibbonButton();
            this.colorCellsWithTextButton = this.Factory.CreateRibbonButton();
            this.colorRowsButton = this.Factory.CreateRibbonButton();
            this.formatTrueFalseButton = this.Factory.CreateRibbonButton();
            this.searchGroup = this.Factory.CreateRibbonGroup();
            this.searchDialogButton = this.Factory.CreateRibbonButton();
            this.fileAndExportGroup = this.Factory.CreateRibbonGroup();
            this.saveEachSheetAsSplitBtn = this.Factory.CreateRibbonSplitButton();
            this.saveEachWorksheetsAsTxtButton = this.Factory.CreateRibbonButton();
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
            this.setupGroup = this.Factory.CreateRibbonGroup();
            this.goToPropertiesButton = this.Factory.CreateRibbonButton();
            this.updateMacrosButton = this.Factory.CreateRibbonButton();
            this.checkMacrosButton = this.Factory.CreateRibbonButton();
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
            this.refreshPivotsButton = this.Factory.CreateRibbonButton();
            this.grandTotalsToggleButton = this.Factory.CreateRibbonToggleButton();
            this.subtotalsToggleButton = this.Factory.CreateRibbonToggleButton();
            this.pivotToolsGroup = this.Factory.CreateRibbonGroup();
            this.combinedTableFromPvValuesButton = this.Factory.CreateRibbonButton();
            this.dataImportTab = this.Factory.CreateRibbonTab();
            this.sqlImportGroup = this.Factory.CreateRibbonGroup();
            this.sqlEditorBtn = this.Factory.CreateRibbonButton();
            this.sqlEditorDataFolderBtn = this.Factory.CreateRibbonButton();
            this.separator12 = this.Factory.CreateRibbonSeparator();
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
            this.miscTab.SuspendLayout();
            this.importGroup.SuspendLayout();
            this.modifiersGroup.SuspendLayout();
            this.filterGroup.SuspendLayout();
            this.validationGroup.SuspendLayout();
            this.searchGroup.SuspendLayout();
            this.fileAndExportGroup.SuspendLayout();
            this.macroGroup.SuspendLayout();
            this.setupGroup.SuspendLayout();
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
            this.miscTab.Groups.Add(this.setupGroup);
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
            this.importSheetOrTxtFileSplitButton.Items.Add(this.importTxtFileLegacyButton);
            this.importSheetOrTxtFileSplitButton.Label = "Import worksheet or txt file";
            this.importSheetOrTxtFileSplitButton.Name = "importSheetOrTxtFileSplitButton";
            this.importSheetOrTxtFileSplitButton.OfficeImageId = "ImportOpml";
            this.importSheetOrTxtFileSplitButton.SuperTip = "Will create window that will accept txt/csv or Excel file and will import delimit" +
    "ed table or first sheet";
            this.importSheetOrTxtFileSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.importSheetOrTxtFile_Click);
            // 
            // importTxtFileLegacyButton
            // 
            this.importTxtFileLegacyButton.Label = "Legacy Import worksheet or txt file";
            this.importTxtFileLegacyButton.Name = "importTxtFileLegacyButton";
            this.importTxtFileLegacyButton.OfficeImageId = "ImportExcel";
            this.importTxtFileLegacyButton.ShowImage = true;
            this.importTxtFileLegacyButton.SuperTip = "Will create window that will accept txt/csv file and will import delimited table " +
    "as text";
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
            this.modifiersGroup.Items.Add(this.removeNaButton);
            this.modifiersGroup.Items.Add(this.removeDuplicatesButton);
            this.modifiersGroup.Items.Add(this.removeHiddenRowsSplitButton);
            this.modifiersGroup.Items.Add(this.clearRangeOutsideButton);
            this.modifiersGroup.Items.Add(this.removeFormattingButton);
            this.modifiersGroup.Items.Add(this.separator1);
            this.modifiersGroup.Items.Add(this.prependTextButton);
            this.modifiersGroup.Items.Add(this.trimButton);
            this.modifiersGroup.Items.Add(this.formatNumberButton);
            this.modifiersGroup.Label = "Modifiers";
            this.modifiersGroup.Name = "modifiersGroup";
            // 
            // changeToTextButton
            // 
            this.changeToTextButton.Label = "To text";
            this.changeToTextButton.Name = "changeToTextButton";
            this.changeToTextButton.OfficeImageId = "AsianLayoutPhoneticGuide";
            this.changeToTextButton.ShowImage = true;
            this.changeToTextButton.SuperTip = "Select square range and it will change it to text";
            this.changeToTextButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.changeToTextButton_Click);
            // 
            // changeToValueButton
            // 
            this.changeToValueButton.Label = "To value";
            this.changeToValueButton.Name = "changeToValueButton";
            this.changeToValueButton.OfficeImageId = "EquationMatrixGallery";
            this.changeToValueButton.ShowImage = true;
            this.changeToValueButton.SuperTip = "Select square range and it will make it general value type";
            this.changeToValueButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.changeToValueButton_Click);
            // 
            // evaluateFormulaButton
            // 
            this.evaluateFormulaButton.Label = "Ev formula";
            this.evaluateFormulaButton.Name = "evaluateFormulaButton";
            this.evaluateFormulaButton.OfficeImageId = "ShowFormulas";
            this.evaluateFormulaButton.ShowImage = true;
            this.evaluateFormulaButton.SuperTip = "Select square range and it should evaluate and replace formulas to value. For arr" +
    "ays select first cell";
            this.evaluateFormulaButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.evaluateFormulaButton_Click);
            // 
            // repasteAsValuesButton
            // 
            this.repasteAsValuesButton.Label = "Repaste as values";
            this.repasteAsValuesButton.Name = "repasteAsValuesButton";
            this.repasteAsValuesButton.OfficeImageId = "PasteValuesAndNumberFormatting";
            this.repasteAsValuesButton.ShowImage = true;
            this.repasteAsValuesButton.SuperTip = "Select square range and it will repaste it as values in the same place";
            this.repasteAsValuesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.repasteAsValuesButton_Click);
            // 
            // sortingAbsButton
            // 
            this.sortingAbsButton.Label = "Sort by abs(values)";
            this.sortingAbsButton.Name = "sortingAbsButton";
            this.sortingAbsButton.OfficeImageId = "Sort";
            this.sortingAbsButton.ShowImage = true;
            this.sortingAbsButton.SuperTip = "Click cell in column that have numbers and it should insert column with absolute " +
    "values and sort from biggest to smallest and remove column (do not use for pivot" +
    " tables)";
            this.sortingAbsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sortingAbsButton_Click);
            // 
            // copyAsPictureButton
            // 
            this.copyAsPictureButton.Label = "Copy as picture";
            this.copyAsPictureButton.Name = "copyAsPictureButton";
            this.copyAsPictureButton.OfficeImageId = "Camera";
            this.copyAsPictureButton.ShowImage = true;
            this.copyAsPictureButton.SuperTip = "Copy selected square range as picture";
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
            this.removeEmptyButton.ShowImage = true;
            this.removeEmptyButton.SuperTip = "Select square range and it will remove empty cells from it (try to use on smaller" +
    " data or save everything before)";
            this.removeEmptyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeEmptyButton_Click);
            // 
            // removeNaButton
            // 
            this.removeNaButton.Label = "Remove #N/A";
            this.removeNaButton.Name = "removeNaButton";
            this.removeNaButton.OfficeImageId = "ConditionalFormattingClearMenu";
            this.removeNaButton.ShowImage = true;
            this.removeNaButton.SuperTip = "Select square range and it will remove N/A values";
            this.removeNaButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeNaButton_Click);
            // 
            // removeDuplicatesButton
            // 
            this.removeDuplicatesButton.Label = "Remove duplicates";
            this.removeDuplicatesButton.Name = "removeDuplicatesButton";
            this.removeDuplicatesButton.OfficeImageId = "RemoveDuplicates";
            this.removeDuplicatesButton.ShowImage = true;
            this.removeDuplicatesButton.SuperTip = "Select square range and it will remove duplicate rows from it";
            this.removeDuplicatesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeDuplicatesButton_Click);
            // 
            // removeHiddenRowsSplitButton
            // 
            this.removeHiddenRowsSplitButton.Items.Add(this.removeHiddenColumnsButton);
            this.removeHiddenRowsSplitButton.Label = "Remove hidden rows";
            this.removeHiddenRowsSplitButton.Name = "removeHiddenRowsSplitButton";
            this.removeHiddenRowsSplitButton.OfficeImageId = "DeleteRows";
            this.removeHiddenRowsSplitButton.SuperTip = "Should remove hidden rows from selected range or region when only one cell select" +
    "ed";
            this.removeHiddenRowsSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeHiddenRowsSplitButton_Click);
            // 
            // removeHiddenColumnsButton
            // 
            this.removeHiddenColumnsButton.Label = "Remove hidden columns";
            this.removeHiddenColumnsButton.Name = "removeHiddenColumnsButton";
            this.removeHiddenColumnsButton.OfficeImageId = "DeleteColumns";
            this.removeHiddenColumnsButton.ShowImage = true;
            this.removeHiddenColumnsButton.SuperTip = "Should remove hidden columns from selected range or region when only one cell sel" +
    "ected";
            this.removeHiddenColumnsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeHiddenColumnsButton_Click);
            // 
            // clearRangeOutsideButton
            // 
            this.clearRangeOutsideButton.Label = "Clear outside range/region";
            this.clearRangeOutsideButton.Name = "clearRangeOutsideButton";
            this.clearRangeOutsideButton.OfficeImageId = "CellStyleNew";
            this.clearRangeOutsideButton.ShowImage = true;
            this.clearRangeOutsideButton.SuperTip = "Clear everything outside region when one cell selected or outside range when sele" +
    "cted square range";
            this.clearRangeOutsideButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.clearRangeOutsideButton_Click);
            // 
            // removeFormattingButton
            // 
            this.removeFormattingButton.Label = "Remove formatting";
            this.removeFormattingButton.Name = "removeFormattingButton";
            this.removeFormattingButton.OfficeImageId = "HighlightClear";
            this.removeFormattingButton.ShowImage = true;
            this.removeFormattingButton.SuperTip = "Select square range and it will remove formatting from that range";
            this.removeFormattingButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeFormattingButton_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // prependTextButton
            // 
            this.prependTextButton.Label = "Prepend";
            this.prependTextButton.Name = "prependTextButton";
            this.prependTextButton.OfficeImageId = "OutlineDemote";
            this.prependTextButton.ShowImage = true;
            this.prependTextButton.SuperTip = "Select column range and it will prepend all values with given text";
            this.prependTextButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.prependTextButton_Click);
            // 
            // trimButton
            // 
            this.trimButton.Label = "Trim";
            this.trimButton.Name = "trimButton";
            this.trimButton.OfficeImageId = "TextDirectionContext";
            this.trimButton.ShowImage = true;
            this.trimButton.SuperTip = "Select range and it will remove leading and trailing spaces from values";
            this.trimButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.trimButton_Click);
            // 
            // formatNumberButton
            // 
            this.formatNumberButton.Label = "Format number";
            this.formatNumberButton.Name = "formatNumberButton";
            this.formatNumberButton.OfficeImageId = "DollarSign";
            this.formatNumberButton.ShowImage = true;
            this.formatNumberButton.SuperTip = "Select any range and it will apply number format for numbers (does not affect tex" +
    "t)";
            this.formatNumberButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.formatNumberButton_Click);
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
            this.filterColumnSplitButton.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.filterColumnSplitButton.Label = "Filter column";
            this.filterColumnSplitButton.Name = "filterColumnSplitButton";
            this.filterColumnSplitButton.OfficeImageId = "FilterBySelection";
            this.filterColumnSplitButton.SuperTip = resources.GetString("filterColumnSplitButton.SuperTip");
            this.filterColumnSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.filterColumnInRangeSplitButton_Click);
            // 
            // filterColumnNotInRangeBtn
            // 
            this.filterColumnNotInRangeBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.filterColumnNotInRangeBtn.Label = "Filter column not in range";
            this.filterColumnNotInRangeBtn.Name = "filterColumnNotInRangeBtn";
            this.filterColumnNotInRangeBtn.OfficeImageId = "FilterClear";
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
            this.filterColumnNotInRegexButton.ShowImage = true;
            this.filterColumnNotInRegexButton.SuperTip = resources.GetString("filterColumnNotInRegexButton.SuperTip");
            this.filterColumnNotInRegexButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.filterColumnNotInRegexButton_Click);
            // 
            // hideRowsWithTextSplitButton
            // 
            this.hideRowsWithTextSplitButton.Items.Add(this.hideColumnsWithTextButton);
            this.hideRowsWithTextSplitButton.Label = "Hide rows with text";
            this.hideRowsWithTextSplitButton.Name = "hideRowsWithTextSplitButton";
            this.hideRowsWithTextSplitButton.OfficeImageId = "GroupTableMerge";
            this.hideRowsWithTextSplitButton.SuperTip = "Will show dialog that will ask for text and it will hide all the entire rows from" +
    " selected square range ";
            this.hideRowsWithTextSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.hideRowsWithTextSplitButton_Click);
            // 
            // hideColumnsWithTextButton
            // 
            this.hideColumnsWithTextButton.Label = "Hide columns with text";
            this.hideColumnsWithTextButton.Name = "hideColumnsWithTextButton";
            this.hideColumnsWithTextButton.OfficeImageId = "GroupTableMerge";
            this.hideColumnsWithTextButton.ShowImage = true;
            this.hideColumnsWithTextButton.SuperTip = "Will show dialog that will ask for text and it will hide all the entire columns f" +
    "rom selected square range ";
            this.hideColumnsWithTextButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.hideColumnsWithTextButton_Click);
            // 
            // takeRowsWithTextButton
            // 
            this.takeRowsWithTextButton.Label = "Take rows with text";
            this.takeRowsWithTextButton.Name = "takeRowsWithTextButton";
            this.takeRowsWithTextButton.OfficeImageId = "TableInsertRowsBelow";
            this.takeRowsWithTextButton.ShowImage = true;
            this.takeRowsWithTextButton.SuperTip = "Select square range and it will show dialog with text input that will take range " +
    "rows that contain that text with headers to new sheet";
            this.takeRowsWithTextButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.takeRowsWithTextButton_Click);
            // 
            // validationGroup
            // 
            this.validationGroup.Items.Add(this.colorRowsUniqueButton);
            this.validationGroup.Items.Add(this.colorRowsWithTextSplitButton);
            this.validationGroup.Items.Add(this.formatTrueFalseButton);
            this.validationGroup.Label = "Validation";
            this.validationGroup.Name = "validationGroup";
            // 
            // colorRowsUniqueButton
            // 
            this.colorRowsUniqueButton.Label = "Color rows unique";
            this.colorRowsUniqueButton.Name = "colorRowsUniqueButton";
            this.colorRowsUniqueButton.OfficeImageId = "GroupResourceGraphFormat";
            this.colorRowsUniqueButton.ShowImage = true;
            this.colorRowsUniqueButton.SuperTip = "Color rows in selected square range based on first column unique values in unique" +
    " colors (approxx 2000 colors and then thay will repeat and may end up next to ea" +
    "ch other)";
            this.colorRowsUniqueButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.colorRowsUniqueButton_Click);
            // 
            // colorRowsWithTextSplitButton
            // 
            this.colorRowsWithTextSplitButton.Items.Add(this.colorColumnsWithTextButton);
            this.colorRowsWithTextSplitButton.Items.Add(this.colorCellsWithTextButton);
            this.colorRowsWithTextSplitButton.Items.Add(this.colorRowsButton);
            this.colorRowsWithTextSplitButton.Label = "Color rows with text";
            this.colorRowsWithTextSplitButton.Name = "colorRowsWithTextSplitButton";
            this.colorRowsWithTextSplitButton.OfficeImageId = "ColorPickerTable";
            this.colorRowsWithTextSplitButton.SuperTip = "Select square range and it will show text input dialog that will color rows with " +
    "that contains it";
            this.colorRowsWithTextSplitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.colorRowsWithTextSplitButton_Click);
            // 
            // colorColumnsWithTextButton
            // 
            this.colorColumnsWithTextButton.Label = "Color columns with text";
            this.colorColumnsWithTextButton.Name = "colorColumnsWithTextButton";
            this.colorColumnsWithTextButton.OfficeImageId = "ColorPickerTable";
            this.colorColumnsWithTextButton.ShowImage = true;
            this.colorColumnsWithTextButton.SuperTip = "Select square range and it will show text input dialog that will color columns wi" +
    "th that contains it";
            this.colorColumnsWithTextButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.colorColumnsWithTextButton_Click);
            // 
            // colorCellsWithTextButton
            // 
            this.colorCellsWithTextButton.Label = "Color cells with text";
            this.colorCellsWithTextButton.Name = "colorCellsWithTextButton";
            this.colorCellsWithTextButton.OfficeImageId = "GroupResourceGraphFormat";
            this.colorCellsWithTextButton.ShowImage = true;
            this.colorCellsWithTextButton.SuperTip = "Select square range and it will show text input dialog that will color cells with" +
    " that contains it";
            this.colorCellsWithTextButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.colorCellsWithTextButton_Click);
            // 
            // colorRowsButton
            // 
            this.colorRowsButton.Label = "Color rows";
            this.colorRowsButton.Name = "colorRowsButton";
            this.colorRowsButton.OfficeImageId = "GroupNetworkDiagramFormat";
            this.colorRowsButton.ShowImage = true;
            this.colorRowsButton.SuperTip = "Select square range and it will color rows in random colors";
            this.colorRowsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.colorRowsButton_Click);
            // 
            // formatTrueFalseButton
            // 
            this.formatTrueFalseButton.Label = "Format TRUE/FALSE";
            this.formatTrueFalseButton.Name = "formatTrueFalseButton";
            this.formatTrueFalseButton.OfficeImageId = "DataValidation";
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
            this.searchDialogButton.Label = "Search dialog";
            this.searchDialogButton.Name = "searchDialogButton";
            this.searchDialogButton.OfficeImageId = "DrawingExplorer";
            this.searchDialogButton.ShowImage = true;
            this.searchDialogButton.SuperTip = "Will show column search form that helps find columns faster";
            this.searchDialogButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.searchDialogButton_Click);
            // 
            // fileAndExportGroup
            // 
            this.fileAndExportGroup.Items.Add(this.saveEachSheetAsSplitBtn);
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
            // saveEachSheetAsSplitBtn
            // 
            this.saveEachSheetAsSplitBtn.Items.Add(this.saveEachWorksheetsAsTxtButton);
            this.saveEachSheetAsSplitBtn.Label = "Save each sheet as xlsx";
            this.saveEachSheetAsSplitBtn.Name = "saveEachSheetAsSplitBtn";
            this.saveEachSheetAsSplitBtn.OfficeImageId = "OrgChartSubordinatesExpand";
            this.saveEachSheetAsSplitBtn.SuperTip = "Saves each sheet from workbook to separate xlsx file into folder named as workboo" +
    "k and into path of that workbook";
            this.saveEachSheetAsSplitBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveEachSheetAsSplitBtn_Click);
            // 
            // saveEachWorksheetsAsTxtButton
            // 
            this.saveEachWorksheetsAsTxtButton.Label = "Save each sheet as txt";
            this.saveEachWorksheetsAsTxtButton.Name = "saveEachWorksheetsAsTxtButton";
            this.saveEachWorksheetsAsTxtButton.OfficeImageId = "OrgChartHorizontalGallery";
            this.saveEachWorksheetsAsTxtButton.ShowImage = true;
            this.saveEachWorksheetsAsTxtButton.SuperTip = "Saves each sheet from workbook to separate txt file into folder named as workbook" +
    " and into path of that workbook";
            this.saveEachWorksheetsAsTxtButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveEachWorksheetsAsTxtButton_Click);
            // 
            // duplicateWorksheetsSplitBtn
            // 
            this.duplicateWorksheetsSplitBtn.Items.Add(this.duplicateWorksheetsToNewWorkbookBtn);
            this.duplicateWorksheetsSplitBtn.Items.Add(this.duplicateWorkbookBtn);
            this.duplicateWorksheetsSplitBtn.Label = "Duplicate selected sheets";
            this.duplicateWorksheetsSplitBtn.Name = "duplicateWorksheetsSplitBtn";
            this.duplicateWorksheetsSplitBtn.OfficeImageId = "DuplicateSelectedSlides";
            this.duplicateWorksheetsSplitBtn.SuperTip = "Will duplicate to the right selected worksheets";
            this.duplicateWorksheetsSplitBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.duplicateWorksheetsSplitBtn_Click);
            // 
            // duplicateWorksheetsToNewWorkbookBtn
            // 
            this.duplicateWorksheetsToNewWorkbookBtn.Label = "Duplicate sheets to new wb";
            this.duplicateWorksheetsToNewWorkbookBtn.Name = "duplicateWorksheetsToNewWorkbookBtn";
            this.duplicateWorksheetsToNewWorkbookBtn.OfficeImageId = "CopyAllRules";
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
            this.duplicateWorkbookBtn.ShowImage = true;
            this.duplicateWorkbookBtn.SuperTip = "Will duplicate whole workbook to new one and ask about saving";
            this.duplicateWorkbookBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.duplicateWorkbookBtn_Click);
            // 
            // saveThisWorksheetAsTxt
            // 
            this.saveThisWorksheetAsTxt.Label = "Save this sheet as txt";
            this.saveThisWorksheetAsTxt.Name = "saveThisWorksheetAsTxt";
            this.saveThisWorksheetAsTxt.OfficeImageId = "ExportTextFile";
            this.saveThisWorksheetAsTxt.ShowImage = true;
            this.saveThisWorksheetAsTxt.SuperTip = "Saves current sheet to txt file and ask about saving";
            this.saveThisWorksheetAsTxt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.saveThisWorksheetAsTxt_Click);
            // 
            // divideTableToPartsAndSaveButton
            // 
            this.divideTableToPartsAndSaveButton.Label = "Divide file";
            this.divideTableToPartsAndSaveButton.Name = "divideTableToPartsAndSaveButton";
            this.divideTableToPartsAndSaveButton.OfficeImageId = "ExportLotus";
            this.divideTableToPartsAndSaveButton.ShowImage = true;
            this.divideTableToPartsAndSaveButton.SuperTip = "Divides table into parts and paste to new worksheets";
            this.divideTableToPartsAndSaveButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.divideTableToPartsAndSaveButton_Click);
            // 
            // getFilePathButton
            // 
            this.getFilePathButton.Label = "Get file path";
            this.getFilePathButton.Name = "getFilePathButton";
            this.getFilePathButton.OfficeImageId = "OpenAttach";
            this.getFilePathButton.ShowImage = true;
            this.getFilePathButton.SuperTip = "Paste file path of workbook into selected cell";
            this.getFilePathButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.getFilePathButton_Click);
            // 
            // exportMacrosButton
            // 
            this.exportMacrosButton.Label = "Export macros";
            this.exportMacrosButton.Name = "exportMacrosButton";
            this.exportMacrosButton.OfficeImageId = "FileMenuPublishHeader";
            this.exportMacrosButton.ScreenTip = "It will export macros from choosen Workbook to filder in Downloads folder.";
            this.exportMacrosButton.ShowImage = true;
            this.exportMacrosButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.exportMacrosButton_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // deleteWorksheetButton
            // 
            this.deleteWorksheetButton.Label = "Delete selected sheets";
            this.deleteWorksheetButton.Name = "deleteWorksheetButton";
            this.deleteWorksheetButton.OfficeImageId = "DeleteTable";
            this.deleteWorksheetButton.ShowImage = true;
            this.deleteWorksheetButton.SuperTip = "Deletes selected sheets from workbook without confirmation";
            this.deleteWorksheetButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.deleteWorksheetButton_Click);
            // 
            // deleteOtherWorksheetsButton
            // 
            this.deleteOtherWorksheetsButton.Label = "Delete other sheets";
            this.deleteOtherWorksheetsButton.Name = "deleteOtherWorksheetsButton";
            this.deleteOtherWorksheetsButton.OfficeImageId = "DeletePagePreviousVersion";
            this.deleteOtherWorksheetsButton.ShowImage = true;
            this.deleteOtherWorksheetsButton.SuperTip = "Deletes other sheets than selected with a warning";
            this.deleteOtherWorksheetsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.deleteOtherWorksheetsButton_Click);
            // 
            // deleteWorkbookButton
            // 
            this.deleteWorkbookButton.Label = "Delete this workbook";
            this.deleteWorkbookButton.Name = "deleteWorkbookButton";
            this.deleteWorkbookButton.OfficeImageId = "DeleteAll";
            this.deleteWorkbookButton.ShowImage = true;
            this.deleteWorkbookButton.SuperTip = "Deletes this workbook with warning";
            this.deleteWorkbookButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.deleteWorkbookButton_Click);
            // 
            // macroGroup
            // 
            this.macroGroup.Items.Add(this.runMacroButton);
            this.macroGroup.Label = "Macro";
            this.macroGroup.Name = "macroGroup";
            // 
            // runMacroButton
            // 
            this.runMacroButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.runMacroButton.Label = "Run Macro Form";
            this.runMacroButton.Name = "runMacroButton";
            this.runMacroButton.OfficeImageId = "PlayMacro";
            this.runMacroButton.ShowImage = true;
            this.runMacroButton.SuperTip = "More user friendly macro runner and searcher. (To use you need to have in Excel t" +
    "rust settings trust VBA project object model.)";
            this.runMacroButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.runMacroButton_Click);
            // 
            // setupGroup
            // 
            this.setupGroup.Items.Add(this.goToPropertiesButton);
            this.setupGroup.Items.Add(this.updateMacrosButton);
            this.setupGroup.Items.Add(this.checkMacrosButton);
            this.setupGroup.Label = "Setup";
            this.setupGroup.Name = "setupGroup";
            // 
            // goToPropertiesButton
            // 
            this.goToPropertiesButton.Label = "Properties files";
            this.goToPropertiesButton.Name = "goToPropertiesButton";
            this.goToPropertiesButton.OfficeImageId = "ProjectManageDeliverables";
            this.goToPropertiesButton.ShowImage = true;
            this.goToPropertiesButton.SuperTip = "Will open location of properties files";
            this.goToPropertiesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.goToPropertiesButton_Click);
            // 
            // updateMacrosButton
            // 
            this.updateMacrosButton.Label = "Update macros";
            this.updateMacrosButton.Name = "updateMacrosButton";
            this.updateMacrosButton.OfficeImageId = "PublishWorkflow";
            this.updateMacrosButton.ShowImage = true;
            this.updateMacrosButton.SuperTip = resources.GetString("updateMacrosButton.SuperTip");
            this.updateMacrosButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.updateMacrosButton_Click);
            // 
            // checkMacrosButton
            // 
            this.checkMacrosButton.Label = "Check macros";
            this.checkMacrosButton.Name = "checkMacrosButton";
            this.checkMacrosButton.OfficeImageId = "MacroDefault";
            this.checkMacrosButton.ShowImage = true;
            this.checkMacrosButton.SuperTip = "It will check if all macros that are assigned to buttons are present";
            this.checkMacrosButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkMacrosButton_Click);
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
            this.generatePivotTemlateCodeButton.ShowImage = true;
            this.generatePivotTemlateCodeButton.SuperTip = "First format pivot table and then you can create template from it and copy templa" +
    "te and paste it manualy into default macro workbook";
            this.generatePivotTemlateCodeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.generatePivotTemlateCodeButton_Click);
            // 
            // runPvTemplateBtn
            // 
            this.runPvTemplateBtn.Label = "Run pv template";
            this.runPvTemplateBtn.Name = "runPvTemplateBtn";
            this.runPvTemplateBtn.OfficeImageId = "PivotTableNewStyle";
            this.runPvTemplateBtn.ShowImage = true;
            this.runPvTemplateBtn.SuperTip = resources.GetString("runPvTemplateBtn.SuperTip");
            this.runPvTemplateBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.runPvTemplateButton_Click);
            // 
            // pivotFormatGroup
            // 
            this.pivotFormatGroup.Items.Add(this.formatClickedPivotButton);
            this.pivotFormatGroup.Items.Add(this.formatAllPivotButton);
            this.pivotFormatGroup.Items.Add(this.refreshPivotsButton);
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
            this.formatClickedPivotButton.ShowImage = true;
            this.formatClickedPivotButton.SuperTip = "Format selected pivot table";
            this.formatClickedPivotButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.formatClickedPivotButton_Click);
            // 
            // formatAllPivotButton
            // 
            this.formatAllPivotButton.Label = "Format all pv";
            this.formatAllPivotButton.Name = "formatAllPivotButton";
            this.formatAllPivotButton.OfficeImageId = "QuerySelectQueryType";
            this.formatAllPivotButton.ShowImage = true;
            this.formatAllPivotButton.SuperTip = "Format all pivot tables in current sheet";
            this.formatAllPivotButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.formatAllPivotButton_Click);
            // 
            // refreshPivotsButton
            // 
            this.refreshPivotsButton.Label = "Refresh all pv";
            this.refreshPivotsButton.Name = "refreshPivotsButton";
            this.refreshPivotsButton.OfficeImageId = "RefreshWebView";
            this.refreshPivotsButton.ShowImage = true;
            this.refreshPivotsButton.SuperTip = "Refresh all pivot tables in current sheet";
            this.refreshPivotsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.refreshPivotsButton_Click);
            // 
            // grandTotalsToggleButton
            // 
            this.grandTotalsToggleButton.Label = "Grand Totals";
            this.grandTotalsToggleButton.Name = "grandTotalsToggleButton";
            this.grandTotalsToggleButton.OfficeImageId = "GroupCalculation";
            this.grandTotalsToggleButton.ShowImage = true;
            this.grandTotalsToggleButton.SuperTip = "Toggle grand totals in selected pivot table";
            this.grandTotalsToggleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.grandTotalsToggleButton_Click);
            // 
            // subtotalsToggleButton
            // 
            this.subtotalsToggleButton.Label = "Subtotals";
            this.subtotalsToggleButton.Name = "subtotalsToggleButton";
            this.subtotalsToggleButton.OfficeImageId = "GroupCalculation";
            this.subtotalsToggleButton.ShowImage = true;
            this.subtotalsToggleButton.SuperTip = "Toggle subtotals in selected pivot table";
            this.subtotalsToggleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.subtotalsToggleButton_Click);
            // 
            // pivotToolsGroup
            // 
            this.pivotToolsGroup.Items.Add(this.combinedTableFromPvValuesButton);
            this.pivotToolsGroup.Label = "Pivot tools";
            this.pivotToolsGroup.Name = "pivotToolsGroup";
            // 
            // combinedTableFromPvValuesButton
            // 
            this.combinedTableFromPvValuesButton.Label = "Table from values";
            this.combinedTableFromPvValuesButton.Name = "combinedTableFromPvValuesButton";
            this.combinedTableFromPvValuesButton.OfficeImageId = "GroupOrganizationChartSelect";
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
            this.sqlImportGroup.Items.Add(this.sqlEditorBtn);
            this.sqlImportGroup.Items.Add(this.sqlEditorDataFolderBtn);
            this.sqlImportGroup.Items.Add(this.separator12);
            this.sqlImportGroup.Items.Add(this.loadToDataTableButton);
            this.sqlImportGroup.Label = "SQL";
            this.sqlImportGroup.Name = "sqlImportGroup";
            // 
            // sqlEditorBtn
            // 
            this.sqlEditorBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.sqlEditorBtn.Image = global::ExcelAddInByMarcinOlszewski.Properties.Resources.sql;
            this.sqlEditorBtn.Label = "SQL Editor";
            this.sqlEditorBtn.Name = "sqlEditorBtn";
            this.sqlEditorBtn.ShowImage = true;
            this.sqlEditorBtn.SuperTip = "Display SQL editor form that let you use queries for Oracle and Microsoft SQL Ser" +
    "ver and paste results into directly into Excel (do not close editor until desire" +
    "d data is not pulled)";
            this.sqlEditorBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sqlEditorBtn_Click);
            // 
            // sqlEditorDataFolderBtn
            // 
            this.sqlEditorDataFolderBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.sqlEditorDataFolderBtn.Label = "SQL Editor Data";
            this.sqlEditorDataFolderBtn.Name = "sqlEditorDataFolderBtn";
            this.sqlEditorDataFolderBtn.OfficeImageId = "Folder";
            this.sqlEditorDataFolderBtn.ShowImage = true;
            this.sqlEditorDataFolderBtn.SuperTip = "Go to folder that SQL Editor saved databases are and SQL queries";
            this.sqlEditorDataFolderBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sqlEditorDataFolderBtn_Click);
            // 
            // separator12
            // 
            this.separator12.Name = "separator12";
            // 
            // loadToDataTableButton
            // 
            this.loadToDataTableButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.loadToDataTableButton.Label = "Load to DataTable";
            this.loadToDataTableButton.Name = "loadToDataTableButton";
            this.loadToDataTableButton.OfficeImageId = "AdpStoredProcedureEditSql";
            this.loadToDataTableButton.ShowImage = true;
            this.loadToDataTableButton.SuperTip = "Will load selected square range to data table (do not select range if you want to" +
    " load file or select range later) and let you SQL query it and paste to excel la" +
    "ter";
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
            this.runS4ExtractButton.Image = global::ExcelAddInByMarcinOlszewski.Properties.Resources.S4_HANA_cloud;
            this.runS4ExtractButton.Label = "Run SAP Extract";
            this.runS4ExtractButton.Name = "runS4ExtractButton";
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
            // 
            // runSdeButton
            // 
            this.runSdeButton.Image = global::ExcelAddInByMarcinOlszewski.Properties.Resources.rocketLogo_scale_400;
            this.runSdeButton.Label = "Run SDE";
            this.runSdeButton.Name = "runSdeButton";
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
            this.sdeQueryComboBox.SuperTip = "Choose SDE Lanucher avaliable queries";
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
            this.browserButton.Image = global::ExcelAddInByMarcinOlszewski.Properties.Resources.Microsoft_Edge_logo__2019__svg;
            this.browserButton.Label = "Launch Browser";
            this.browserButton.Name = "browserButton";
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
            this.importFromBrowserCheckBox.SuperTip = "Auto import downloaded Excel files and txt/csv files";
            // 
            // MiscRibbon
            // 
            this.Name = "MiscRibbon";
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
            this.setupGroup.ResumeLayout(false);
            this.setupGroup.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton prependTextButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton changeToTextButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton changeToValueButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton evaluateFormulaButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton repasteAsValuesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton removeEmptyButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton removeNaButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton trimButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton formatNumberButton;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton saveEachWorksheetsAsTxtButton;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton saveEachSheetAsSplitBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton duplicateWorksheetsSplitBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton duplicateWorkbookBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton duplicateWorksheetsToNewWorkbookBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton sqlEditorBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton sqlEditorDataFolderBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton grandTotalsToggleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton filterColumnFromRangeNotInRangeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton runPvTemplateBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton removeDuplicatesButton;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator12;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loadToDataTableButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton colorColumnsWithTextButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton colorCellsWithTextButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton filterColumnInRegexButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton filterColumnNotInRegexButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup setupGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton goToPropertiesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton checkMacrosButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton updateMacrosButton;
        public Microsoft.Office.Tools.Ribbon.RibbonTab miscTab;
        public Microsoft.Office.Tools.Ribbon.RibbonTab pivotToolsTab;
        public Microsoft.Office.Tools.Ribbon.RibbonTab dataImportTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup pivotToolsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton combinedTableFromPvValuesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton exportMacrosButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton hideRowsWithTextSplitButton;
    }

    partial class ThisRibbonCollection
    {
        internal MiscRibbon MiscRibbon
        {
            get { return this.GetRibbon<MiscRibbon>(); }
        }
    }
}
