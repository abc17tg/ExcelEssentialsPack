using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Linq;
using ExcelAddInByMarcinOlszewski.Forms;
using ExcelAddInByMarcinOlszewski.Scripts;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using WTC = ImportTableToExcel.WorksheetFromTxtCreator;

namespace ExcelAddInByMarcinOlszewski
{
    public partial class MiscRibbon
    {
        private Excel.Workbook m_macroWorkbook;
        private Excel.Workbook m_functionsWorkbook;

        private void MiscRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            var assembly = Assembly.GetExecutingAssembly();

            GetMacrosWorkbooks();

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
            using (new ExcelExecutionBlock(app))
            {
                Excel.Range rng = app.ActiveWindow.RangeSelection;
                rng.NumberFormat = "General";
                foreach (Excel.Range col in rng.Columns)
                {
                    try
                    {
                        col.TextToColumns(
                        DataType: Excel.XlTextParsingType.xlDelimited,
                        TextQualifier: Excel.XlTextQualifier.xlTextQualifierNone,
                        ConsecutiveDelimiter: false,
                        Tab: false,
                        Semicolon: false,
                        Comma: false,
                        Space: false,
                        Other: false,
                        FieldInfo: new object[] { new object[] { 1, Excel.XlColumnDataType.xlTextFormat } }
                    );
                    }
                    catch { }
                }
            }
        }

        private void changeToValueButton_Click(object sender, RibbonControlEventArgs e)
        {
            //Utils.RunMacro("ConvertSelectedRangeToValues");
            Excel.Application app = Globals.ThisAddIn.Application;
            using (new ExcelExecutionBlock(app))
            {
                Excel.Range rng = app.ActiveWindow.RangeSelection;
                rng.NumberFormat = "General";
                foreach (Excel.Range col in rng.Columns)
                {
                    try
                    {
                        col.TextToColumns(
                        DataType: Excel.XlTextParsingType.xlDelimited,
                        TextQualifier: Excel.XlTextQualifier.xlTextQualifierNone,
                        ConsecutiveDelimiter: false,
                        Tab: false,
                        Semicolon: false,
                        Comma: false,
                        Space: false,
                        Other: false,
                        FieldInfo: new object[] { new object[] { 1, Excel.XlColumnDataType.xlGeneralFormat } }
                    );
                    }
                    catch { }
                }
            }
        }

        private void evaluateFormulaButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("Converting.EvaluateAndReplaceFormula");
        }

        private void repasteAsValuesButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("Converting.RepasteSelectedRangeAsValues");
        }

        private void removeEmptyButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("RemoveCells.RemoveEmptyCells");
        }

        private void removeNaButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("RemoveCells.DeleteNAFromSelection");
        }

        private void prependTextButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("Utils.PrependText");
        }

        private void trimButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("Utils.RemoveLeadingTrailingSpaces");
        }

        private void formatNumberButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.ApplyCustomNumberFormat(app.ActiveWindow.RangeSelection);
            /*using (new ExcelExecutionBlock(app))
            {
                (app.Selection as Excel.Range).NumberFormatLocal = @"[Color49]# ##0.00;[Color9]-# ##0.00;[Color16]0;@";
            }*/

        }

        private void hideRowsWithTextButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("HideRows.HideRowsWithValue");
        }

        private void takeRowsWithTextButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("GetRowsThatContainTextValue.GetRowsThatContainTextValueInput");
        }

        private void searchDialogButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("HeadersListForm.OpenFinder");
            /*SearchColumnsForm form = new SearchColumnsForm(Globals.ThisAddIn.Application);
            form.Show();
            form.FormClosed += (s, _) =>
            {
            };*/
        }

        private void colorRowsUniqueButton_Click(object sender, RibbonControlEventArgs e)
        {
            //Utils.RunMacro("ColorSelectedRowsUnique");
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Range rng = app.ActiveWindow.RangeSelection.Columns[1];
                Dictionary<string, Color> valueColorD = new Dictionary<string, Color>();
                List<string> values;
                values = rng.Cells.Cast<Excel.Range>().Select(p => ((object)p.Value)?.ToString() ?? "").Distinct().ToList();
                List<Color> colorsList = Utils.GenerateColorPalette(values.Count);
                colorsList.Shuffle();
                for (int i = 0; i < values.Count; i++)
                    valueColorD.Add(values[i], colorsList[i]);

                using (new ExcelExecutionBlock(app))
                {
                    app.ActiveWindow.RangeSelection.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;

                    foreach (Excel.Range r in app.ActiveWindow.RangeSelection.Rows.Cast<Excel.Range>())
                        r.Interior.Color = valueColorD[((object)r.Columns[1].Value)?.ToString() ?? ""];

                    Excel.Range row;
                    string val = null, oldVal = null;
                    for (int i = 1; i <= app.ActiveWindow.RangeSelection.Rows.Count; i++)
                    {
                        row = app.ActiveWindow.RangeSelection.Rows[i] as Excel.Range;
                        oldVal = val;
                        val = ((object)row.Columns[1].Value)?.ToString() ?? "";
                        row.Interior.Color = valueColorD[val];

                        if (i == 1)
                            continue;

                        if (!val.Equals(oldVal, StringComparison.Ordinal))
                        {
                            Excel.Border border = row.Borders[Excel.XlBordersIndex.xlEdgeTop];
                            border.Color = ColorTranslator.FromOle((int)((double)row.Interior.Color)).DarkenColor(0.5f).ToArgb();
                            border.Weight = Excel.XlBorderWeight.xlThin;
                            border.LineStyle = Excel.XlLineStyle.xlContinuous;
                        }
                    }
                }
            }
            catch { }
        }

        private void colorRowsButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("ColorRows.ColorSelectedRows");
        }

        private void formatTrueFalseButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("Utils.ConditionalFormattingTRUEandFALSE");
        }

        private void filterColumnInRangeSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.FilterByRange(app.Selection);
        }

        private void filterColumnNotInRangeButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.FilterByRange(app.Selection, true);
        }

        private void filterColumnFromRangeInRangeButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.FilterByRange(app.Selection, false, true);
        }

        private void filterColumnFromRangeNotInRangeButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.FilterByRange(app.Selection, true, true);
        }

        private void sortingAbsButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection;
            UtilsExcel.SortColumnByAbsoluteValues(rng);
        }

        private void filterColumnInRegexButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.FilterByRegex(app.Selection, false);
        }

        private void filterColumnNotInRegexButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            UtilsExcel.FilterByRegex(app.Selection, true);
        }

        private void saveEachWorksheetsAsTxtButton_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("DivideFile.SaveSheetsAsTextFiles");
        }

        private void saveEachSheetAsSplitBtn_Click(object sender, RibbonControlEventArgs e)
        {
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonSplitButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("DivideFile.SaveSheetsAsExcelFiles");
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
                saveDlg.InitialDirectory = Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();

            saveDlg.FileName = ws.Name;
            saveDlg.OverwritePrompt = false;
            saveDlg.DefaultExt = ".txt";
            saveDlg.AddExtension = true;
            saveDlg.Filter = "Text Files | *.txt";

            if (saveDlg.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    wb = app.Workbooks.Add();
                    Excel.Worksheet newWs = wb.Worksheets.Add();
                    ws.Copy(newWs);
                    newWs.Delete();
                    wb.SaveAs(saveDlg.FileName, Excel.XlFileFormat.xlTextWindows);
                    app.DisplayAlerts = false;
                    wb.Close();
                    app.DisplayAlerts = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
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
            UtilsExcel.RunMacro(Macro.GetMacroNameForButton((sender as RibbonButton).Id, m_macroWorkbook));
            //UtilsExcel.RunMacro("FilesSubs.GetCurrentFilePath");
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

        private void copyAsPictureButton_Click(object sender, RibbonControlEventArgs e)
        {
            (Globals.ThisAddIn.Application.Selection as Excel.Range).CopyPicture(Excel.XlPictureAppearance.xlScreen,
                Excel.XlCopyPictureFormat.xlBitmap);
        }

        private void removeDuplicatesButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection;

            if (!rng.Valid())
                return;

            object[] columIndexArray = Enumerable.Range(1, rng.Columns.Count).Cast<object>().ToArray();
            using (new ExcelExecutionBlock(app))
            {
                rng.RemoveDuplicates((object)columIndexArray, Excel.XlYesNoGuess.xlNo);
            }
        }

        private void runMacroButton_Click(object sender, RibbonControlEventArgs e)
        {
            RunMacroForm form = new RunMacroForm(m_macroWorkbook);
            form.Show();
        }


        private void colorRowsWithTextSplitButton_Click(object sender, RibbonControlEventArgs e)
        {
            ColorCellsWithTextForm form = new ColorCellsWithTextForm(Globals.ThisAddIn.Application, ColorCellsWithTextForm.Type.Rows);
            form.Show();
        }
        private void colorCellsWithTextButton_Click(object sender, RibbonControlEventArgs e)
        {
            ColorCellsWithTextForm form = new ColorCellsWithTextForm(Globals.ThisAddIn.Application, ColorCellsWithTextForm.Type.Cells);
            form.Show();
        }

        private void colorColumnsWithTextButton_Click(object sender, RibbonControlEventArgs e)
        {
            ColorCellsWithTextForm form = new ColorCellsWithTextForm(Globals.ThisAddIn.Application, ColorCellsWithTextForm.Type.Colums);
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
            UtilsExcel.RemoveFormattingFromRange(rng);
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

        private void grandTotalsToggleButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection;
            if (rng.IsPivotCell() && rng.PivotCell.PivotTable != null)
            {
                var pv = rng.PivotCell.PivotTable;
                bool toggle = !pv.ColumnGrand && !pv.RowGrand;
                pv.ColumnGrand = toggle;
                pv.RowGrand = toggle;
                grandTotalsToggleButton.Checked = toggle;
            }
        }

        private void subtotalsToggleButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.ActiveWindow.RangeSelection;
            if (rng.IsPivotCell() && rng.PivotCell.PivotTable != null)
            {
                var pf = rng.PivotCell.PivotField;
                pf.Subtotals[1] = !pf.Subtotals[1];
                subtotalsToggleButton.Checked = pf.Subtotals[1];
            }
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
                WorkbookPickerForm workbookPicker = new WorkbookPickerForm(m_macroWorkbook);
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
                                logs += $"\n{fileName}\t{(UtilsExcel.UpdateModule(fileName, workbookPicker.Workbook) ? "✔" : "❌")}";
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
            WorkbookPickerForm workbookPicker = new WorkbookPickerForm(m_macroWorkbook);
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
    }
}
