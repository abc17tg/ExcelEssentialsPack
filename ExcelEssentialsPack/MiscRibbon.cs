using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using ExcelEssentials.Forms;
using ExcelEssentials.Scripts;
using Microsoft.Office.Tools.Ribbon;
using Color = System.Drawing.Color;
using Excel = Microsoft.Office.Interop.Excel;
using WTC = ImportTableToExcel.WorksheetFromTxtCreator;

namespace ExcelEssentials
{
    public partial class MiscRibbon
    {
        private Excel.Workbook m_macroWorkbook;
        private Excel.Workbook m_functionsWorkbook;

        private async void MiscRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            //var assembly = Assembly.GetExecutingAssembly();
            //GetMacrosWorkbooks();
            FileManager.CheckForCustomMacrosWbNames();
            m_macroWorkbook = await EnsureWorkbookIsOpenAsync(FileManager.MacrosWbName);
            m_functionsWorkbook = await EnsureWorkbookIsOpenAsync(FileManager.FunctionsWbName);

            // add event to every pivot button
            List<RibbonButton> createPivotButtons = createPivotFromTemplateMenu.Items.Where(p => p.GetType().Name == "RibbonButtonImpl").Cast<RibbonButton>().ToList();
            createPivotButtons.AddRange((createPivotFromTemplateMenu.Items.Where(p => p.GetType().Name == "RibbonMenuImpl").ToList().First() as RibbonMenu).Items.Where(p => p.GetType().Name == "RibbonButtonImpl").Cast<RibbonButton>().ToList());
            foreach (var pvBtn in createPivotButtons)
            {
                pvBtn.SuperTip = $"PivotTablesTemplates.{pvBtn.Name.Replace("Button", "")}";
                pvBtn.Click += (s, _) => UtilsExcel.RunMacro($"PivotTablesTemplates.{(s as RibbonButton).Name.Replace("Button", "")}");
            }

            List<string> queries = Sde.SdeQueries();
            if (queries != null)
                foreach (var query in queries)
                {
                    RibbonDropDownItem ribbonDropDownItem = Factory.CreateRibbonDropDownItem();
                    ribbonDropDownItem.Label = query;
                    sdeQueryComboBox.Items.Add(ribbonDropDownItem);
                }

            List<string> bookmarks = BrowserViewForm.GetBookmarks().Keys.ToList();
            if (bookmarks != null)
                foreach (var bookmark in bookmarks)
                {
                    RibbonDropDownItem ribbonDropDownItem = Factory.CreateRibbonDropDownItem();
                    ribbonDropDownItem.Label = bookmark;
                    browserWebsitesComboBox.Items.Add(ribbonDropDownItem);
                }
            //SQLitePCL.Batteries.Init();
        }

        public enum MissingWorkbookAction
        {
            Ignore,
            TryOpen,
            Prompt
        }

        public async Task<Excel.Workbook> EnsureWorkbookIsOpenAsync(string workbookName)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Workbook wb = GetOpenWorkbook(app, workbookName);
            if (wb != null)
                return wb;

            await Task.Delay(5000);

            wb = CheckAndOpenWorkbook(app, workbookName, MissingWorkbookAction.TryOpen);
            if (wb != null)
                return wb;

            wb = CheckAndOpenWorkbook(app, workbookName, MissingWorkbookAction.Prompt);
            return wb;
        }

        private Excel.Workbook GetOpenWorkbook(Excel.Application app, string workbookName)
        {
            try
            {
                var wb = Globals.ThisAddIn.Application.Workbooks[FileManager.MacrosWbName];
                if (wb != null)
                    return wb;
                else
                    return null;
            }
            catch (COMException)
            {
                return null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private Excel.Workbook CheckAndOpenWorkbook(Excel.Application app, string workbookName, MissingWorkbookAction action)
        {
            Excel.Workbook wb = GetOpenWorkbook(app, workbookName);
            if (wb != null)
                return wb;

            string startupPath = app.StartupPath;
            string fullPath = Path.Combine(startupPath, workbookName);
            bool fileExists = File.Exists(fullPath);

            switch (action)
            {
                case MissingWorkbookAction.Ignore:
                    return null;

                case MissingWorkbookAction.TryOpen:
                    if (fileExists)
                    {
                        try { return app.Workbooks.Open(fullPath); }
                        catch (Exception) { }
                    }
                    else
                    {
                        return null;
                    }
                    break;

                case MissingWorkbookAction.Prompt:
                    wb = CheckAndOpenWorkbook(app, workbookName, MissingWorkbookAction.TryOpen);
                    if (wb != null)
                        return wb;
                    else
                    {
                        DialogResult result = MessageBox.Show($"{workbookName} was not found in XLSTART foder:\n\"{startupPath}\"\nLocate manually?",
                                                              "Workbook Not Found", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            using (OpenFileDialog ofd = new OpenFileDialog())
                            {
                                ofd.Filter = "Excel macros workbooks|*.xlsb;*.xlam;*.xlsm";
                                ofd.Title = $"Locate macro workbook {workbookName}";
                                ofd.InitialDirectory = startupPath;
                                ofd.Multiselect = false;
                                if (ofd.ShowDialog() == DialogResult.OK)
                                {
                                    try { return app.Workbooks.Open(ofd.FileName); }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show($"Error opening {workbookName}:\n{ex.Message}",
                                                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }
                            }
                        }
                    }
                    return null;
            }

            return null;
        }

        private bool GetMacrosWorkbooks()
        {
            FileManager.CheckForCustomMacrosWbNames();

            bool result = false;
            try
            {
                m_macroWorkbook = Globals.ThisAddIn.Application.Workbooks[FileManager.MacrosWbName];
                if (m_macroWorkbook != null)
                    result = true;
                else
                    result = false;
            }
            catch
            {
                MessageBox.Show($"Problem seeing workbook \"{FileManager.MacrosWbName}\", check if it exists and is open.");
            }

            try
            {
                m_functionsWorkbook = Globals.ThisAddIn.Application.Workbooks[FileManager.FunctionsWbName];
                if (m_functionsWorkbook != null && m_macroWorkbook != null)
                    result = true;
                else
                    result = false;
            }
            catch
            {
                MessageBox.Show($"Problem seeing workbook \"{FileManager.FunctionsWbName}\", check if it exists and is open.");
            }
            return result;
        }

        private void changeToTextButton_Click(object sender, RibbonControlEventArgs e)
        {
            //Utils.RunMacro("ConvertSelectedRangeToText");

            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection;

            if (!rng.Valid())
                return;

            using (new ExcelExecutionBlock(app))
            {
                foreach (var ar in rng.Areas.Cast<Excel.Range>())
                {
                    if (!ar.Valid())
                        continue;
                    ar.GetUsableRange().ChangeToText();
                }
            }
        }

        private void changeToValueButton_Click(object sender, RibbonControlEventArgs e)
        {
            //Utils.RunMacro("ConvertSelectedRangeToValues");
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection;

            if (!rng.Valid())
                return;

            using (new ExcelExecutionBlock(app))
            {
                foreach (var ar in rng.Areas.Cast<Excel.Range>())
                {
                    if (!ar.Valid())
                        continue;
                    ar.GetUsableRange().ChangeToValue();
                }
            }
        }

        private void evaluateFormulaButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("Converting.EvaluateAndReplaceFormula");
        }

        private void repasteAsValuesButton_Click(object sender, RibbonControlEventArgs e)
        {
            //UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection;

            if (!rng.Valid())
                return;

            using (new ExcelExecutionBlock(app))
            {
                foreach (var ar in rng.Areas.Cast<Excel.Range>())
                {
                    if (!ar.Valid())
                        continue;

                    Excel.Range firstCell = ar.Cells[1, 1] as Excel.Range;
                    string formula = null;
                    if (firstCell.Valid() && firstCell.HasFormula)
                        formula = firstCell.Formula;

                    ar.GetUsableRange().RepasteAsValues();

                    if (formula != null)
                        Clipboard.SetText(formula);
                }
            }
            //UtilsExcel.RunMacro("Converting.RepasteSelectedRangeAsValues");
        }

        private void removeEmptyButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("RemoveCells.RemoveEmptyCells");
        }
        private void removeErrSplitBtn_Click(object sender, RibbonControlEventArgs e)
        {
            //UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonSplitButton).Id, m_macroWorkbook));
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection.GetUsableRange();
            if (!rng.Valid())
                return;

            using (new ExcelExecutionBlock(app))
            {
                foreach (var cell in rng.Cells.Cast<Excel.Range>())
                {
                    try
                    {
                        if (app.WorksheetFunction.IsError(cell))
                            cell.Clear();
                    }
                    catch (COMException) { }
                    catch (Exception) { }
                }
            }
        }

        private void removeNaButton_Click(object sender, RibbonControlEventArgs e)
        {
            //UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("RemoveCells.DeleteNAFromSelection");
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection.GetUsableRange();
            if (!rng.Valid())
                return;

            using (new ExcelExecutionBlock(app))
            {
                foreach (var cell in rng.Cells.Cast<Excel.Range>())
                {
                    try
                    {
                        if (app.WorksheetFunction.IsNA(cell) || cell.Value2.ToString() == "#N/A")
                            cell.Clear();
                    }
                    catch (COMException) { }
                    catch (Exception) { }
                }
            }
        }

        private void prependTextSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            //UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonSplitButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("Utils.PrependText");

            Excel.Application app = Globals.ThisAddIn.Application;

            string input = app.InputBox("Enter the text to prepend:", "Input text");
            if (string.IsNullOrEmpty(input))
                return;

            Excel.Range rng = app.ActiveWindow.RangeSelection.GetUsableRange();
            if (!rng.Valid())
                return;

            using (new ExcelExecutionBlock(app))
            {
                foreach (var cell in rng.Cells.Cast<Excel.Range>())
                {
                    if (!app.WorksheetFunction.IsError(cell) && !string.IsNullOrEmpty(cell.Value2?.ToString()))
                    {
                        if (cell.NumberFormat != "@" && !char.IsDigit(input[0]) && input[0] != '0')
                            cell.NumberFormat = "@";
                        cell.Value2 = string.Concat(input, cell.Value2.ToString());
                    }
                }
            }
        }

        private void appendTextButton_Click(object sender, RibbonControlEventArgs e)
        {
            //UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            Excel.Application app = Globals.ThisAddIn.Application;
            bool isInputNumeric = false;
            string input = app.InputBox("Enter the text to append:", "Input text");
            if (string.IsNullOrEmpty(input))
                return;

            Excel.Range rng = app.ActiveWindow.RangeSelection.GetUsableRange();
            if (!rng.Valid())
                return;

            isInputNumeric = long.TryParse(input, out _);

            using (new ExcelExecutionBlock(app))
            {
                foreach (var cell in rng.Cells.Cast<Excel.Range>())
                {
                    if (!app.WorksheetFunction.IsError(cell) && !string.IsNullOrEmpty(cell.Value2?.ToString()))
                    {
                        if (cell.NumberFormat != "@" && (!isInputNumeric || !(double.TryParse(cell.Value2.ToString(), out double _) || long.TryParse(cell.Value2.ToString(), out long _))))
                            cell.NumberFormat = "@";
                        cell.Value2 = string.Concat(cell.Value2.ToString(), input);
                    }
                }
            }
        }

        private void trimButton_Click(object sender, RibbonControlEventArgs e)
        {
            //UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("Utils.RemoveLeadingTrailingSpaces");
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection.GetUsableRange();

            if (!rng.Valid())
                return;

            using (new ExcelExecutionBlock(app))
            {
                foreach (var cell in rng.Cells.Cast<Excel.Range>())
                {
                    if (app.WorksheetFunction.IsError(cell) || string.IsNullOrEmpty(cell.Value2?.ToString()))
                        continue;

                    string trimmed = (cell.Value2.ToString()).Trim();
                    if (cell.Value2.ToString() != trimmed)
                    {
                        if (long.TryParse(trimmed, out _) && cell.NumberFormat == "General")
                            cell.NumberFormat = "@";
                        cell.Value2 = (cell.Value2.ToString()).Trim();
                    }
                }
            }
        }

        private void formatNumberSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.ApplyCustomNumberFormat(app.ActiveWindow.RangeSelection);
        }

        private void formatStringToDateButton_Click(object sender, RibbonControlEventArgs e)
        {
            FormatToDateForm formatToDateForm = new FormatToDateForm(Globals.ThisAddIn.Application);
            formatToDateForm.Show();
        }

        private void hideRowsWithTextSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            string text = Microsoft.VisualBasic.Interaction.InputBox("Type text that will hide rows with it in cell", "Text input", string.Empty);
            UtilsExcel.HideRangeWithText(app.ActiveWindow.RangeSelection.GetUsableRange(), text, UtilsExcel.DirectionType.Rows);
        }

        private void hideColumnsWithTextButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            string text = Microsoft.VisualBasic.Interaction.InputBox("Type text that will hide columns with it in cell", "Text input", string.Empty);
            UtilsExcel.HideRangeWithText(app.ActiveWindow.RangeSelection.GetUsableRange(), text, UtilsExcel.DirectionType.Colums);
        }

        private void takeRowsWithTextButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("GetRowsThatContainTextValue.GetRowsThatContainTextValueInput");
        }

        private void searchDialogButton_Click(object sender, RibbonControlEventArgs e)
        {
            //UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("HeadersListForm.OpenFinder");
            SearchColumnsForm form = new SearchColumnsForm(Globals.ThisAddIn.Application);
            form.Show();
            form.FormClosed += (s, _) =>
            {
            };
        }

        private void colorRowsUniqueSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            //Utils.RunMacro("ColorSelectedRowsUnique");
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Range rng = app.ActiveWindow.RangeSelection;

                if (!rng.Valid())
                    return;

                UtilsExcel.ColorRowsUnique(rng.GetUsableRange());
            }
            catch (Exception) { }
        }

        private void colorCellsUniqueButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Range rngSel = app.ActiveWindow.RangeSelection.GetUsableRange();
                if (rngSel == null)
                    return;

                Dictionary<string, Color> valueColorD = new Dictionary<string, Color>();
                List<string> values;
                values = rngSel.Cells.Cast<Excel.Range>().Select(p => ((object)p.Value2)?.ToString() ?? "").Distinct().ToList();

                var getUniqueColorsTaskResult = new Task<Dictionary<string, Color>>(() =>
                {
                    List<Color> colorsList = Utils.GenerateColorPalette(values.Count);
                    colorsList.Shuffle();
                    for (int i = 0; i < values.Count; i++)
                        valueColorD.Add(values[i], colorsList[i]);
                    try
                    { valueColorD[""] = Color.WhiteSmoke; }
                    catch (Exception) { }
                    return valueColorD;
                });

                getUniqueColorsTaskResult.GetAwaiter().OnCompleted(() =>
                {
                    if (app == null)
                        return;

                    if (Globals.ThisAddIn.Dispatcher == null)
                        return;

                    Globals.ThisAddIn.Dispatcher.Invoke(new Action(() =>
                    {
                        try
                        {
                            app.StatusBar = string.Empty;
                            using (new ExcelExecutionBlock(app))
                            {
                                rngSel.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;

                                foreach (Excel.Range c in rngSel.Cells.Cast<Excel.Range>())
                                    c.Interior.Color = valueColorD[((object)c.Value)?.ToString() ?? ""];

                                Excel.Range cell;
                                string val = null, oldVal = null;

                                foreach (Excel.Range col in rngSel.Columns.Cast<Excel.Range>())
                                {
                                    for (int i = 1; i <= col.Cells.Count; i++)
                                    {
                                        cell = col.Cells[i] as Excel.Range;
                                        oldVal = val;
                                        val = ((object)cell.Value)?.ToString() ?? "";
                                        cell.Interior.Color = valueColorD[val];

                                        if (i == 1)
                                            continue;

                                        if (!val.Equals(oldVal, StringComparison.Ordinal))
                                        {
                                            Excel.Border border = cell.Borders[Excel.XlBordersIndex.xlEdgeTop];
                                            border.Color = ColorTranslator.FromOle((int)((double)cell.Interior.Color)).DarkenColor(0.5f).ToArgb();
                                            border.Weight = Excel.XlBorderWeight.xlThin;
                                            border.LineStyle = Excel.XlLineStyle.xlContinuous;
                                        }
                                    }
                                }
                            }
                        }
                        catch (COMException)
                        {
                            app.StatusBar = "Err";
                        }
                        catch (Exception)
                        {
                            app.StatusBar = "Err";
                        }
                    }));
                });

                app.StatusBar = "Getting unique colors";
                getUniqueColorsTaskResult.Start();
            }
            catch (Exception) { }
        }

        private void colorRowsButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("ColorRows.ColorSelectedRows");
        }

        private void formatTrueFalseButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("Utils.ConditionalFormattingTRUEandFALSE");
        }

        private void filterColumnInRangeSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.FilterByRange(app.ActiveWindow.RangeSelection);
        }

        private void filterColumnNotInRangeButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.FilterByRange(app.ActiveWindow.RangeSelection, true);
        }

        private void filterColumnFromRangeInRangeButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.FilterByRange(app.ActiveWindow.RangeSelection, false, true);
        }

        private void filterColumnFromRangeNotInRangeButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.FilterByRange(app.ActiveWindow.RangeSelection, true, true);
        }

        private void filterColumnInRegexButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.FilterByRegex(app.ActiveWindow.RangeSelection, false);
        }

        private void filterColumnNotInRegexButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.FilterByRegex(app.ActiveWindow.RangeSelection, true);
        }

        private void sortingAbsButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection;
            UtilsExcel.SortColumnByAbsoluteValues(rng);
        }

        private void saveSelectedWorksheetsAsXlsxSplitBtn_Click(object sender, RibbonControlEventArgs e)
        {
            //UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonSplitButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("DivideFile.SaveSheetsAsExcelFiles");
            Excel.Application app = Globals.ThisAddIn.Application;
            try
            {
                Excel.Sheets wss = app.ActiveWindow.SelectedSheets;
                UtilsExcel.SaveWorksheetsAsExcelFiles(wss);
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(ex.Message, "Operation Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void saveAllWorksheetsAsXlsxButton_Click(object sender, RibbonControlEventArgs e)
        {
            //UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            Excel.Application app = Globals.ThisAddIn.Application;
            try
            {
                Excel.Sheets wss = app.ActiveWorkbook.Sheets;
                UtilsExcel.SaveWorksheetsAsExcelFiles(wss);
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(ex.Message, "Operation Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void saveSelectedWorksheetsAsTxtButton_Click(object sender, RibbonControlEventArgs e)
        {
            //UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            Excel.Application app = Globals.ThisAddIn.Application;
            try
            {
                Excel.Sheets wss = app.ActiveWindow.SelectedSheets;
                UtilsExcel.SaveWorksheetsAsTxtFiles(wss);
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(ex.Message, "Operation Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void saveAllWorksheetsAsTxtButton_Click(object sender, RibbonControlEventArgs e)
        {
            //UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            Excel.Application app = Globals.ThisAddIn.Application;
            try
            {
                Excel.Sheets wss = app.ActiveWorkbook.Sheets;
                UtilsExcel.SaveWorksheetsAsTxtFiles(wss);
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(ex.Message, "Operation Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void duplicateWorkbookBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.MoveWorksheets2NewWb(app, app.ActiveWorkbook.Worksheets.Cast<Excel.Worksheet>().ToList());
        }

        private void duplicateWorksheetsToNewWorkbookBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.MoveWorksheets2NewWb(app, app.ActiveWindow.SelectedSheets.Cast<Excel.Worksheet>().ToList());
        }

        private void duplicateWorksheetsSplitBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;

            List<Excel.Worksheet> ws2Copy = app.ActiveWindow.SelectedSheets.Cast<Excel.Worksheet>().ToList();
            ws2Copy.Reverse();

            Excel.Worksheet startWs = ws2Copy.First();

            foreach (var ws in ws2Copy)
                ws.Copy(After: startWs);
        }

        private void saveThisWorksheetAsTxt_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Workbook wb = app.ActiveWorkbook;
            Excel.Worksheet ws = app.ActiveSheet;

            SaveFileDialog saveDlg = new SaveFileDialog();

            if (!string.IsNullOrEmpty(wb.Path))
                saveDlg.InitialDirectory = wb.Path;
            else
                saveDlg.InitialDirectory = FileManager.DownloadsPath;

            saveDlg.FileName = ws.Name;
            saveDlg.OverwritePrompt = false;
            saveDlg.DefaultExt = ".txt";
            saveDlg.AddExtension = true;
            saveDlg.Filter = "Text Files | *.txt";

            if (saveDlg.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Excel.Workbook newWb;
                    Excel.Worksheet newWs;

                    // Copy the worksheet into the new workbook
                    ws.Copy();
                    newWb = app.ActiveWorkbook;
                    // Get the reference to the copied worksheet, which is now the active sheet
                    newWs = app.ActiveSheet;

                    if (newWb == null || newWs == null)
                        throw new InvalidOperationException("Failed to retrieve the copied worksheet.");

                    // Check for new lines or tabs in the copied worksheet
                    UtilsExcel.CheckAndAskToRemoveNewLineOrTabInCellsInWorksheet(newWs);

                    // Save the new workbook as a text file
                    app.DisplayAlerts = false;
                    newWb.SaveAs(saveDlg.FileName, Excel.XlFileFormat.xlTextWindows);
                    newWb.Close(false);
                    app.DisplayAlerts = true;
                    Marshal.ReleaseComObject(newWb);

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to save sheet '{ws.Name}': {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void divideTableToPartsAndSaveButton_Click(object sender, RibbonControlEventArgs e)
        {
            //UtilsExcel.RunMacro("DivideTableToParts");
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection;
            if (rng.Cells.Count < 2)
                rng = rng.CurrentRegion;

            if (!int.TryParse(Microsoft.VisualBasic.Interaction.InputBox("To parts", "Divide into parts", "2"), out int parts))
            {
                MessageBox.Show("Wrong input number!");
                return;
            }

            UtilsExcel.DivideTableToParts(rng, parts);
        }

        private void getFilePathButton_Click(object sender, RibbonControlEventArgs e)
        {
            //UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("FilesSubs.GetCurrentFilePath");
            string path = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;
            var result = MessageBox.Show($"Workbook path: {path ?? string.Empty}{Environment.NewLine}Copy?", "Workbook path", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (result == DialogResult.Yes)
                Clipboard.SetText(path);
        }

        private void deleteWorksheetButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;

            if (app.ActiveWindow.SelectedSheets.Count >= app.ActiveWorkbook.Worksheets.Count)
                return;

            using (new ExcelExecutionBlock(app))
            {
                app.DisplayAlerts = false;
                app.ActiveWindow.SelectedSheets.Delete();
                app.DisplayAlerts = true;
            }
        }

        private void deleteOtherWorksheetsButton_Click(object sender, RibbonControlEventArgs e)
        {
            var result = MessageBox.Show("This is an irreversible operation, confirm that you want to continue.", "Confirm", MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result != DialogResult.OK)
                return;

            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Sheets wss = app.ActiveWindow.SelectedSheets;
            Excel.Sheets allWss = app.ActiveWorkbook.Worksheets;

            if (wss.Count == allWss.Count)
                return;

            using (new ExcelExecutionBlock(app))
            {
                app.DisplayAlerts = false;
                foreach (Excel.Worksheet sheet in allWss)
                {
                    bool delete = true;
                    foreach (Excel.Worksheet sheet2 in wss)
                        if (sheet.Name == sheet2.Name)
                        {
                            delete = false;
                            break;
                        }
                    if (delete)
                        sheet.Delete();
                }
                app.DisplayAlerts = true;
            }
        }

        private void deleteWorkbookButton_Click(object sender, RibbonControlEventArgs e)
        {
            var result = MessageBox.Show("This is an irreversible operation, confirm that you want to continue.", "Confirm", MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (result != DialogResult.OK)
                return;
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Workbook wb = app.ActiveWorkbook;
            string filePath = wb.FullName;
            wb.Close();
            if (!string.IsNullOrEmpty(wb.Path) && File.Exists(filePath))
                File.Delete(filePath);
        }

        private void importSheetOrTxtFile_Click(object sender, RibbonControlEventArgs e)
        {
            string filePath = string.Empty;
            FileDropForm form = new FileDropForm(Utils.TextExt.Concat(Utils.ExcelExt).ToList());
            form.Show();
            form.FormClosed += (s, _) =>
            {
                if (form.DialogResult != DialogResult.OK)
                    return;

                filePath = form.FilePath;
                if (Utils.TextExt.Contains(Path.GetExtension(filePath), StringComparer.OrdinalIgnoreCase))
                {
                    char delimiter = Utils.DetermineTableDelimiter(filePath);
                    if (delimiter == default(char))
                    {
                        string choosenDelimiter = Microsoft.VisualBasic.Interaction.InputBox("Can not determine delimiter, write one in ' characters:", "Write delimiter in ''", "", 0, 0);

                        if (choosenDelimiter.Length != 1)
                        {
                            MessageBox.Show("Delimiter too long or missing!", "Delimiter too long or missing!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else
                            delimiter = choosenDelimiter[0];
                    }
                    Excel.Application app = Globals.ThisAddIn.Application;
                    Excel.Worksheet aWs = app.ActiveSheet;
                    Excel.Worksheet ws = (aWs.Parent as Excel.Workbook).Worksheets.Add(aWs);

                    if (File.ReadLines(filePath).LongCount() > ws.Rows.Count)
                    {
                        int columnCount = 0; long rowCount = File.ReadLines(filePath).LongCount();
                        using (StreamReader reader = new StreamReader(filePath))
                        {
                            string firstLine = reader.ReadLine();
                            columnCount = !string.IsNullOrEmpty(firstLine) ? firstLine.Split(delimiter).Length : 0;
                        }
                        UtilsExcel.RunMacro("LoadTextFileIntoDataModel", new object[] { $"\"{filePath}\"", delimiter.ToString(), columnCount.ToString() });
                        return;
                    }
                    else
                        WTC.ImportTextFileToExcel(ws, filePath, delimiter);

                    ws.Rename(Path.GetFileNameWithoutExtension(filePath));
                }
                else if (Utils.ExcelExt.Contains(Path.GetExtension(filePath), StringComparer.OrdinalIgnoreCase))
                {
                    Excel.Workbook wb = Microsoft.VisualBasic.Interaction.GetObject(filePath) as Excel.Workbook;
                    Excel.Application app = Globals.ThisAddIn.Application;
                    Excel.Worksheet aWs = app.ActiveSheet;
                    wb.Worksheets.Item[1].Copy(aWs);
                    if ((wb.Worksheets.Item[1].Name as string).StartsWith("Sheet"))
                        (app.ActiveSheet as Excel.Worksheet).Rename(Path.GetFileNameWithoutExtension(wb.FullName));
                    //Utils.RunMacro("RenameSheet", new object[] { Path.GetFileNameWithoutExtension(wb.FullName) });
                    wb.Close();
                    return;
                }
                else
                    return;
            };
        }

        private void importSheetOrTxtFileAdv_Click(object sender, RibbonControlEventArgs e)
        {
            string filePath = string.Empty;
            FileDropForm form = new FileDropForm(Utils.TextExt.Concat(Utils.ExcelExt).ToList());
            form.Show();
            form.FormClosed += (s, _) =>
            {
                if (form.DialogResult != DialogResult.OK)
                    return;

                filePath = form.FilePath;
                if (Utils.TextExt.Contains(Path.GetExtension(filePath), StringComparer.OrdinalIgnoreCase))
                {
                    char delimiter = Utils.DetermineTableDelimiter(filePath);
                    if (delimiter == default(char))
                    {
                        string choosenDelimiter = Microsoft.VisualBasic.Interaction.InputBox("Can not determine delimiter, write one in ' characters:", "Write delimiter in ''", "", 0, 0);

                        if (choosenDelimiter.Length != 1)
                        {
                            MessageBox.Show("Delimiter too long or missing!", "Delimiter too long or missing!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else
                            delimiter = choosenDelimiter[0];
                    }
                    Excel.Application app = Globals.ThisAddIn.Application;
                    Excel.Worksheet aWs = app.ActiveSheet;
                    Excel.Worksheet ws = (aWs.Parent as Excel.Workbook).Worksheets.Add(aWs);

                    if (File.ReadLines(filePath).LongCount() > ws.Rows.Count)
                    {
                        int columnCount = 0; long rowCount = File.ReadLines(filePath).LongCount();
                        using (StreamReader reader = new StreamReader(filePath))
                        {
                            string firstLine = reader.ReadLine();
                            columnCount = !string.IsNullOrEmpty(firstLine) ? firstLine.Split(delimiter).Length : 0;
                        }
                        UtilsExcel.RunMacro("LoadTextFileIntoDataModel", new object[] { $"\"{filePath}\"", delimiter.ToString(), columnCount.ToString() });
                        return;
                    }
                    else
                        WTC.ImportTextFileToExcelAdv(ws, filePath, delimiter);

                    ws.Rename(Path.GetFileNameWithoutExtension(filePath));
                }
                else if (Utils.ExcelExt.Contains(Path.GetExtension(filePath), StringComparer.OrdinalIgnoreCase))
                {
                    Excel.Workbook wb = Microsoft.VisualBasic.Interaction.GetObject(filePath) as Excel.Workbook;
                    Excel.Application app = Globals.ThisAddIn.Application;
                    Excel.Worksheet aWs = app.ActiveSheet;
                    wb.Worksheets.Item[1].Copy(aWs);
                    if ((wb.Worksheets.Item[1].Name as string).StartsWith("Sheet"))
                        (app.ActiveSheet as Excel.Worksheet).Rename(Path.GetFileNameWithoutExtension(wb.FullName));
                    //Utils.RunMacro("RenameSheet", new object[] { Path.GetFileNameWithoutExtension(wb.FullName) });
                    wb.Close();
                    return;
                }
                else
                    return;
            };
        }

        private void importTxtFileLegacyButton_Click(object sender, RibbonControlEventArgs e)
        {
            string filePath = string.Empty;
            FileDropForm form = new FileDropForm(Utils.TextExt.Concat(Utils.ExcelExt).ToList());
            form.Show();
            form.FormClosed += (s, _) =>
            {
                if (form.DialogResult != DialogResult.OK)
                    return;

                filePath = form.FilePath;
                if (Utils.TextExt.Contains(Path.GetExtension(filePath), StringComparer.OrdinalIgnoreCase))
                {
                    char delimiter = Utils.DetermineTableDelimiter(filePath);
                    if (delimiter == default(char))
                    {
                        string choosenDelimiter = Microsoft.VisualBasic.Interaction.InputBox("Can not determine delimiter, write one in ' characters:", "Write delimiter in ''", "", 0, 0);

                        if (choosenDelimiter.Length != 1)
                        {
                            MessageBox.Show("Delimiter too long or missing!", "Delimiter too long or missing!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else
                            delimiter = choosenDelimiter[0];
                    }
                    Excel.Application app = Globals.ThisAddIn.Application;
                    Excel.Worksheet aWs = app.ActiveSheet;
                    Excel.Worksheet ws = (aWs.Parent as Excel.Workbook).Worksheets.Add(aWs);

                    if (File.ReadLines(filePath).LongCount() > ws.Rows.Count)
                    {
                        int columnCount = 0; long rowCount = File.ReadLines(filePath).LongCount();
                        using (StreamReader reader = new StreamReader(filePath))
                        {
                            string firstLine = reader.ReadLine();
                            columnCount = !string.IsNullOrEmpty(firstLine) ? firstLine.Split(delimiter).Length : 0;
                        }
                        UtilsExcel.RunMacro("LoadTextFileIntoDataModel", new object[] { $"\"{filePath}\"", delimiter.ToString(), columnCount.ToString() });
                        return;
                    }
                    else
                        WTC.ImportTextFileToExcelLegacy(ws, filePath, delimiter);

                    ws.Rename(Path.GetFileNameWithoutExtension(filePath));
                }
                else if (Utils.ExcelExt.Contains(Path.GetExtension(filePath), StringComparer.OrdinalIgnoreCase))
                {
                    Excel.Workbook wb = Microsoft.VisualBasic.Interaction.GetObject(filePath) as Excel.Workbook;
                    Excel.Application app = Globals.ThisAddIn.Application;
                    Excel.Worksheet aWs = app.ActiveSheet;
                    wb.Worksheets.Item[1].Copy(aWs);
                    if ((wb.Worksheets.Item[1].Name as string).StartsWith("Sheet"))
                        (app.ActiveSheet as Excel.Worksheet).Rename(Path.GetFileNameWithoutExtension(wb.FullName));
                    //Utils.RunMacro("RenameSheet", new object[] { Path.GetFileNameWithoutExtension(wb.FullName) });
                    wb.Close();
                    return;
                }
                else
                    return;
            };
        }

        private void copyAsPictureButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveWindow.RangeSelection.GetUsableRange().CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);
        }

        private void removeDuplicatesButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection.GetUsableRange();

            if (!rng.Valid())
                return;

            using (new ExcelExecutionBlock(app))
            {
                foreach (var ar in rng.Areas.Cast<Excel.Range>())
                {
                    if (!ar.Valid())
                        continue;
                    object[] columIndexArray = Enumerable.Range(1, ar.Columns.Count).Cast<object>().ToArray();
                    ar.RemoveDuplicates((object)columIndexArray, Excel.XlYesNoGuess.xlNo);
                }
            }
        }

        private void removeDuplicatesInColumnsButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection;

            if (!rng.Valid())
                return;

            using (new ExcelExecutionBlock(app))
            {
                foreach (var ar in rng.Areas.Cast<Excel.Range>())
                {
                    if (!ar.Valid())
                        continue;

                    foreach (var col in ar.Columns.Cast<Excel.Range>())
                    {
                        Excel.Range column = col.GetUsableRange();
                        if (!column.Valid())
                            continue;
                        column.RemoveDuplicates(1, Header: Excel.XlYesNoGuess.xlNo);
                    }
                }
            }
        }

        private void runMacroButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            RunMacroForm form = new RunMacroForm(m_macroWorkbook);
            form.Show();
        }

        private void colorRowsWithTextSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            ColorCellsWithTextForm form = new ColorCellsWithTextForm(Globals.ThisAddIn.Application, UtilsExcel.RangeType.Rows);
            form.Show();
        }

        private void colorCellsWithTextButton_Click(object sender, RibbonControlEventArgs e)
        {
            ColorCellsWithTextForm form = new ColorCellsWithTextForm(Globals.ThisAddIn.Application, UtilsExcel.RangeType.Cells);
            form.Show();
        }

        private void colorColumnsWithTextButton_Click(object sender, RibbonControlEventArgs e)
        {
            ColorCellsWithTextForm form = new ColorCellsWithTextForm(Globals.ThisAddIn.Application, UtilsExcel.RangeType.Colums);
            form.Show();
        }

        private void removeHiddenColumnsButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection.CurrentRegion;
            UtilsExcel.DeleteNonVisibleColumns(rng);
            rng.Columns.Hidden = false;
        }

        private void removeHiddenRowsSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection.CurrentRegion;
            UtilsExcel.DeleteNonVisibleRows(rng);
            rng.Rows.Hidden = false;
        }

        private void removeFormattingButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection;

            if (!rng.Valid())
                return;

            using (new ExcelExecutionBlock(app))
            {
                foreach (var ar in rng.Areas.Cast<Excel.Range>())
                {
                    if (!ar.Valid())
                        continue;

                    UtilsExcel.RemoveFormattingFromRange(ar);
                }
            }
        }

        private void clearRangeOutsideButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection;
            UtilsExcel.DeleteOutsideRngOrRegion(rng);
        }

        private void goToPropertiesButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!FileManager.IsExplorerPathOpen(FileManager.PropertiesFilesPath))
                Process.Start("explorer.exe", FileManager.PropertiesFilesPath);
        }

        private void checkMacrosButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string filePath = Path.Combine(FileManager.PropertiesFilesPath, "ButtonSubroutineMapping.xml");
                XElement xe = XElement.Load(filePath);

                // Using LINQ to extract ButtonID and Subroutine values
                var mappings = from mapping in xe.Elements("Mapping")
                               select new
                               {
                                   ButtonID = mapping.Element("ButtonID").Value.Trim(),
                                   Subroutine = mapping.Element("Subroutine").Value.Trim()
                               };

                string message = string.Join("\n", mappings.Select(m => $"{m.ButtonID}\t{m.Subroutine}\t{(Macro.Exists(m.Subroutine.Split('.')[1], m.Subroutine.Split('.')[0], m_macroWorkbook) ? "✔" : "❌")}"));
                MessageBoxForm messageBox = new MessageBoxForm(message, "Macros mapping", true);
                messageBox.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void updateMacrosButton_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Update files|*.bas;*.macro";
            var result = openFileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                WorkbookPickerForm workbookPicker = new WorkbookPickerForm(Globals.ThisAddIn.Application);
                var result2 = workbookPicker.ShowDialog();
                if (result2 == DialogResult.OK)
                {
                    string logs = string.Empty;
                    foreach (string fileName in openFileDialog.FileNames)
                    {
                        string extension = Path.GetExtension(fileName);
                        switch (extension.ToLower())
                        {
                            case ".bas":
                                bool replace = MessageBox.Show("Keep macros that do not exist in updated version of VBA module?", $"Module {Path.GetFileNameWithoutExtension(fileName)}", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No;
                                logs += $"\n{fileName}\t{(UtilsExcel.UpdateModule(fileName, workbookPicker.Workbook, replace) ? "✔" : "❌")}";
                                break;
                            case ".macro":
                                logs += $"\n{fileName}\t{(UtilsExcel.UpdateMacro(fileName, workbookPicker.Workbook) ? "✔" : "❌")}";
                                break;
                            default:
                                MessageBox.Show("Unsupported file type.");
                                break;
                        }
                    }
                    MessageBox.Show(logs, "Result of update");
                }
            }
        }

        private void exportMacrosButton_Click(object sender, RibbonControlEventArgs e)
        {
            WorkbookPickerForm workbookPicker = new WorkbookPickerForm(Globals.ThisAddIn.Application);
            var result = workbookPicker.ShowDialog();
            if (result == DialogResult.OK)
            {
                string path = workbookPicker.Workbook.ExportMacros();
                var msgResult = MessageBox.Show($"Operation exporting macros ended with {(!string.IsNullOrEmpty(path) ? "success.\n\nOpen directory?" : "fail")}", "Export macros result", !string.IsNullOrEmpty(path) ? MessageBoxButtons.YesNo : MessageBoxButtons.OK);
                if (msgResult == DialogResult.Yes)
                {
                    if (!FileManager.IsExplorerPathOpen(path))
                        Process.Start("explorer.exe", path);
                }
            }
        }

        private void runCustomFormButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (m_macroWorkbook == null)
                GetMacrosWorkbooks();
            RunMacroWithSearchPhraseForm form = new RunMacroWithSearchPhraseForm(m_macroWorkbook, "RunCustom");
            form.Show();
        }

        private void fillEmptyWithAboveValueButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection.GetUsableRange();

            if (!rng.Valid())
                return;

            UtilsExcel.FillEmptyCellWithAboveValue(rng);
        }

        private void copyDelimitedValuesButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection.GetUsableRange();

            if (!rng.Valid())
                return;

            FormatDelimitedForm form = new FormatDelimitedForm(rng);
            form.Show();
        }

        private void selectWithoutHeadersButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection;
            UtilsExcel.SelectCurrentRegionWithoutHeaders(rng);
        }

        private void createMacroUpdateButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            string directory = Directory.CreateDirectory(Path.Combine(FileManager.DownloadsPath, "Macros Updates")).FullName;
            UtilsExcel.CreateMacroUpdateFileFromActiveVbaCode(app, directory);
        }
    }
}
