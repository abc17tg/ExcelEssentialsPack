using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelAddInByMarcinOlszewski;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelVB = Microsoft.Vbe.Interop;
using System.Runtime.InteropServices;
using ExcelAddInByMarcinOlszewski.Forms;
using System.Data;
using ExcelAddInByMarcinOlszewski.Scripts;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Numerics;
using System.Drawing;
using static SQLite.TableMapping;

public static class UtilsExcel
{
    public static void Updating(this Excel.Application application, bool enableUpdates)
    {
        application.ScreenUpdating = enableUpdates;
        application.Calculation = enableUpdates ? Excel.XlCalculation.xlCalculationAutomatic : Excel.XlCalculation.xlCalculationManual;
        application.EnableEvents = enableUpdates;
    }

    public static object RunMacro(string macroName, string macroWbName = "")
    {
        if (macroName == string.Empty)
        {
            MessageBox.Show("Macro not implemented yet.");
            return null;
        }

        try
        {
            return RunMacro(Globals.ThisAddIn.Application, new Object[] { $"{(string.IsNullOrWhiteSpace(macroWbName) ? FileManager.MacrosWbName : macroWbName)}!{macroName}" });
        }
        catch (Exception e)
        {
            MessageBox.Show($"Problem running macro \"{(string.IsNullOrWhiteSpace(macroWbName) ? FileManager.MacrosWbName : macroWbName)}!{macroName}\", check if you have workbook \"{(string.IsNullOrWhiteSpace(macroWbName) ? FileManager.MacrosWbName : macroWbName)}\" open with macro \"{macroName}\" in any module\n\n{e.Message}");
        }
        return null;
    }

    public static object RunMacro(string macroName, object[] args, string macroWbName = "")
    {
        if (macroName == string.Empty)
        {
            MessageBox.Show("Macro not implemented yet.");
            return null;
        }

        try
        {
            object o = $"{(string.IsNullOrWhiteSpace(macroWbName) ? FileManager.MacrosWbName : macroWbName)}!{macroName}";
            object[] objects = args;
            objects = objects.Prepend(o).ToArray();
            return RunMacro(Globals.ThisAddIn.Application, objects);
        }
        catch (Exception e)
        {
            MessageBox.Show($"Problem running macro \"{(string.IsNullOrWhiteSpace(macroWbName) ? FileManager.MacrosWbName : macroWbName)}!{macroName}\", check if you have workbook \"{(string.IsNullOrWhiteSpace(macroWbName) ? FileManager.MacrosWbName : macroWbName)}\" open with macro \"{macroName}\" in any module\n\n{e.Message}");
        }
        return null;
    }

    private static object RunMacro(object oApp, object[] oRunArgs)
    {
        return oApp.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod, null, oApp, oRunArgs);
    }

    public static string FormatRangeToSqlPattern(Excel.Range rng)
    {
        if (!rng.Valid())
            return string.Empty;

        if (rng.Cells.Count == 1 && rng.Row == 1)
            return rng.Text.ToString();

        List<string> values = new List<string>();
        foreach (Excel.Range c in rng.Cells.Cast<Excel.Range>())
            if (c != null && !string.IsNullOrEmpty(c.Text) && !c.EntireRow.Hidden && !c.EntireColumn.Hidden)
                values.Add(c.Text.ToString());

        if (values.Count < 1)
            return string.Empty;

        if (rng.Cells.Count > 1 && rng.Rows.Count == 1)
            return string.Join(", ", values.Distinct().Select(p => p.Contains(" ") ? $"[{p}]" : p).ToList());

        return $"('{string.Join("', '", values.Distinct())}')";
    }

    public static string GenerateSqlFilterFromExcelSelection(Excel.Range rng)
    {
        if (!rng.Valid() || rng.Rows.Count < 2)
            return string.Empty;

        var filterParts = new List<string>();

        int columnCount = rng.Columns.Count;
        int rowCount = rng.Rows.Count;

        for (int row = 2; row <= rowCount; row++) // Start from 2 to skip header row
        {
            var rowFilterParts = new List<string>();

            for (int column = 1; column <= columnCount; column++)
            {
                string fieldName = Convert.ToString(rng.Cells[1, column].Value); // Get field name from first row
                if (fieldName.Contains(" "))
                    fieldName = $"[{fieldName}]";
                string fieldValue = Convert.ToString(rng.Cells[row, column].Value); // Get field value from current row

                rowFilterParts.Add($"{fieldName} IN ('{fieldValue}')");
            }

            filterParts.Add($"\t\t{string.Join("\n\t\tAND\n\t\t", rowFilterParts)}");
        }

        return $"\n(\n\t(\n{string.Join("\n\t)\n\tOR\n\t(\n", filterParts)}\n\t)\n)\n";
    }

    public static void Rename<T>(this T ws, string name, string append = "") where T : Excel.Worksheet
    {
        List<string> wsNames = (ws.Parent as Excel.Workbook).Worksheets.Cast<Excel.Worksheet>().AsEnumerable().Select(p => p.Name).ToList();

        string newName = name.Substring(0, Math.Min(name.Length, 31 - append.Length)) + append;
        int i = 1;

        while (wsNames.Contains(newName, StringComparer.OrdinalIgnoreCase))
        {
            string suffix = $" ({i})";
            int baseStringLength = 31 - append.Length - suffix.Length;
            newName = name.Substring(0, Math.Min(baseStringLength, name.Length)) + append + suffix;
            i++;
        }
        ws.Name = newName;
    }

    public static void FilterByRange(Excel.Range selection, bool notInMode = false, bool chooseRangeByForm = false)
    {
        bool isSelectionRange, isPivotTable;
        isSelectionRange = selection is Excel.Range;
        isPivotTable = selection.IsPivotCell() && selection.PivotCell.PivotTable != null;
        if (!(isSelectionRange || isPivotTable) || (isPivotTable && chooseRangeByForm))
            return;

        Excel.PivotTable pivotTable = null;
        Excel.PivotField pivotField = null;
        Excel.Range rng = null;

        try
        {
            InputRangeForm getRangeToUseAsFilterForm = new InputRangeForm(selection.Application, "Select range as filter", "Select range as filter");

            if (chooseRangeByForm)
            {
                InputRangeForm getRangeToUseTableToFilterForm = new InputRangeForm(selection.Application, "Select range as area to filter", "Select range as area to filter");
                getRangeToUseTableToFilterForm.Show();
                getRangeToUseTableToFilterForm.FormClosed += (s, _) =>
                {
                    if (getRangeToUseTableToFilterForm.DialogResult != DialogResult.OK)
                        return;

                    rng = selection.CurrentRegion;

                    rng = getRangeToUseTableToFilterForm.Range;
                    getRangeToUseAsFilterForm.Show();
                };
            }

            if (!chooseRangeByForm)
                getRangeToUseAsFilterForm.Show();

            getRangeToUseAsFilterForm.FormClosed += (s, _) =>
            {
                if (getRangeToUseAsFilterForm.DialogResult != DialogResult.OK)
                    return;

                selection.Application.Updating(false);
                if (isPivotTable)
                {
                    pivotTable = selection.PivotCell.PivotTable;
                    pivotField = selection.PivotCell.PivotField;
                }
                else if (!chooseRangeByForm)
                    rng = selection.CurrentRegion;

                var valuesToFilter = getRangeToUseAsFilterForm.Range.Cells.Cast<Excel.Range>().Select(p => ((object)p.Value).ToString()).Distinct().ToList();
                if (notInMode)
                {
                    List<string> values = new List<string>();
                    Excel.Range column;
                    if (isPivotTable)
                    {
                        foreach (Excel.PivotItem item in pivotField.PivotItems())
                            values.Add(item.Value.ToString());
                    }
                    else
                    {
                        column = rng.Columns[selection.Column - rng.Column + 1];
                        values = column.Cells.Cast<Excel.Range>().Select(p => ((object)p.Value).ToString()).Skip(1).ToList();
                    }

                    valuesToFilter = values.Distinct().Except(valuesToFilter).ToList();
                }
                if (isPivotTable)
                    FilterPivotItems(pivotField, valuesToFilter);
                else
                    rng.AutoFilter(selection.Column - rng.Column + 1, valuesToFilter.ToArray(), Excel.XlAutoFilterOperator.xlFilterValues);

                selection.Application.Updating(true);
            };
        }
        catch (Exception ex)
        {
            selection.Application.Updating(true);
            MessageBox.Show($"There was a problem with filtering{(isPivotTable ? " pivot" : "")} table {ex.Message}");
        }
    }

    public static void FilterByRegex(Excel.Range selection, bool notInMode = false, string regexString = "")
    {
        bool isSelectionRange, isPivotTable;
        isSelectionRange = selection is Excel.Range;
        isPivotTable = selection.IsPivotCell() && selection.PivotCell.PivotTable != null;
        if (!(isSelectionRange || isPivotTable))
            return;

        Excel.PivotTable pivotTable = null;
        Excel.PivotField pivotField = null;
        Excel.Range rng = null;

        try
        {
            if (string.IsNullOrWhiteSpace(regexString))
                regexString = Microsoft.VisualBasic.Interaction.InputBox("Paste Regular Expresions (Regex) pattern", "Regex pattern", string.Empty);

            if (string.IsNullOrWhiteSpace(regexString))
                return;

            selection.Application.Updating(false);
            if (isPivotTable)
            {
                pivotTable = selection.PivotCell.PivotTable;
                pivotField = selection.PivotCell.PivotField;
            }

            rng = selection.CurrentRegion;

            List<string> values = new List<string>();
            Excel.Range column;
            if (isPivotTable)
            {
                foreach (Excel.PivotItem item in pivotField.PivotItems())
                    values.Add(item.Value.ToString());
            }
            else
            {
                column = rng.Columns[selection.Column - rng.Column + 1];
                values = column.Cells.Cast<Excel.Range>().Select(p => ((object)p.Value).ToString()).Skip(1).ToList();
            }

            Regex regex = new Regex(regexString);
            var valuesToFilter = values.FindAll(p => regex.IsMatch(p));

            if (notInMode)
            {
                valuesToFilter = values.Distinct().Except(valuesToFilter).ToList();
            }

            if (isPivotTable)
                FilterPivotItems(pivotField, valuesToFilter);
            else
                rng.AutoFilter(selection.Column - rng.Column + 1, valuesToFilter.ToArray(), Excel.XlAutoFilterOperator.xlFilterValues);

            selection.Application.Updating(true);
        }
        catch (Exception ex)
        {
            selection.Application.Updating(true);
            MessageBox.Show($"There was a problem with filtering{(isPivotTable ? " pivot" : "")} table {ex.Message}");
        }
    }

    public static bool IsPivotCell<T>(this T cell) where T : Excel.Range
    {
        try
        {
            var pivotCell = cell.PivotCell;
            return true;
        }
        catch (COMException)
        {
            return false;
        }
    }


    public static void FilterPivotItems(Excel.PivotField pf, List<string> pivotItemNames)
    {
        Excel.PivotItems pis = pf.ChildItems;

        // Orientation != XlPivotFieldOrientation.xlHidden and we need to filter by at least one value (as Excel implies)
        if (pf.Orientation != 0 && pivotItemNames.Count > 0)
        {
            int oldAutoSortOrder = 0;

            if (pf.AutoSortOrder != (int)Excel.Constants.xlManual)
            {
                oldAutoSortOrder = pf.AutoSortOrder;
                pf.AutoSort((int)Excel.Constants.xlManual, pf.Name);
            }

            int pivotItemsCount = pf.PivotItems().Count;
            List<int> pivotItemsToHide = new List<int>();

            for (int i = 1; i <= pivotItemsCount; i++)
            {
                Excel.PivotItem pi = pf.PivotItems(i);

                // check if current pivot item needs to be hidden (if it exists in pivotItemNames)
                var match = pivotItemNames.FirstOrDefault(stringToCheck => stringToCheck.Equals(pi.Value));

                if (match == null)
                {
                    // hide these pivot items later because we can hit exception "Unable to set the Visible property of the PivotItem class"
                    // (this happens because all pivot items get hidden and we need to have at least one visible)
                    pivotItemsToHide.Add(i);
                }
                else
                {
                    TryFilterPivotItems(pi, true, true);
                }
            }

            for (int i = 0; i < pivotItemsToHide.Count; i++)
            {
                Excel.PivotItem pi = pf.PivotItems(pivotItemsToHide[i]);
                TryFilterPivotItems(pi, false, true);
            }

            if (oldAutoSortOrder != 0)
            {
                pf.AutoSort(oldAutoSortOrder, pf.Name);
            }

            Excel.PivotTable pt = pf.Parent as Excel.PivotTable;
            if (pt != null)
            {
                pt.Update();
            }
        }
    }

    public static void TryFilterPivotItems(Excel.PivotItem currentPI, bool filterValue, bool deferLayoutUpdate = false)
    {
        try
        {
            Excel.PivotField pf = currentPI.Parent;
            Excel.PivotTable pt = pf.Parent as Excel.PivotTable;

            if (currentPI.Visible != filterValue)
            {
                if (deferLayoutUpdate == true && pt != null)
                {
                    // just keep these three lines stick together, no if, no nothing (otherwise ManualUpdate will reset back to false)
                    pt.ManualUpdate = true;
                    currentPI.Visible = filterValue;

                    // this may be redundant since setting Visible property of pivot item, resets ManualUpdate to false
                    pt.ManualUpdate = false;
                }
                else
                {
                    currentPI.Visible = filterValue;
                }
            }
        }
        catch (Exception)
        {

        }
    }

    public static void TryFilterPivotItems(Excel.PivotField pf, string itemValue, bool filterValue, bool deferLayoutUpdate = false)
    {
        try
        {
            Excel.PivotItem currentPI = pf.PivotItems(itemValue);
            TryFilterPivotItems(currentPI, filterValue, deferLayoutUpdate);
        }
        catch (Exception)
        {

        }
    }

    public static void PasteDataTableToRange(DataTable dt, Excel.Range rng, bool headers = true)
    {
        if (dt == null || !rng.Valid())
            return;

        Excel.Range startCell = rng.Cells[1, 1];
        Excel.Worksheet ws = rng.Worksheet;

        if (dt.Rows.Count < 1 || dt.Columns.Count < 1)
            return;

        object[,] dataArr = new object[dt.Rows.Count + (headers ? 1 : 0), dt.Columns.Count];

        if (headers)
            for (int c = 0; c < dt.Columns.Count; c++)
                dataArr[0, c] = dt.Columns[c].ColumnName;

        // Convert DataTable to 2D array
        int startRow = headers ? 1 : 0;
        object lockObject = new object();
        Parallel.For(startRow, dt.Rows.Count + startRow, r =>
        {
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                if (dt.Rows[r - startRow][c] is DateTime)
                    lock (lockObject)
                        dataArr[r, c] = dt.Rows[r - startRow][c].ToString();
                else
                    lock (lockObject)
                        dataArr[r, c] = dt.Rows[r - startRow][c];
            }
        });

        // Define the range
        Excel.Range endCell = ws.Cells[startCell.Row + dt.Rows.Count + (headers ? 0 : -1), startCell.Column + dt.Columns.Count - 1];
        Excel.Range writeRange = ws.Range[startCell, endCell];

        double t = 0;
        while (!ws.Application.Ready)
        {
            if (t == 1800)
                break;
            System.Threading.Thread.Sleep(100); // Sleep for 100 milliseconds
            t += 0.1;
        }

        // Set number formats for each column based on first row of data
        using (new ExcelExecutionBlock(ws.Application))
        {
            for (int c = dt.Columns.Count - 1; c >= 0; c--)
            {
                if (dataArr[(headers ? 1 : 0), c] is string || dataArr[(headers ? 1 : 0), c] is DateTime)
                {
                    Excel.Range columnRange = ws.Range[ws.Cells[startCell.Row + (headers ? 1 : 0), startCell.Column + c], ws.Cells[startCell.Row + dt.Rows.Count + (headers ? 0 : -1), startCell.Column + c]];
                    columnRange.NumberFormat = "@";
                }
            }

            // Write data to Excel in one go
            writeRange.Value = dataArr;
        }
    }

    public static void DeleteNonVisibleRows(Excel.Range rng)
    {
        for (int i = rng.Rows.Count; i >= 1; i--)
            if (rng.Rows[i].Hidden)
                rng.Rows[i].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
    }

    public static void DeleteNonVisibleColumns(Excel.Range rng)
    {
        for (int i = rng.Columns.Count; i >= 1; i--)
            if (rng.Columns[i].Hidden)
                rng.Columns[i].Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
    }

    public static void DivideTableToParts(Excel.Range tableRng, int parts)
    {
        if (tableRng == null || parts < 1)
            return;

        Excel.Worksheet ws = tableRng.Worksheet;
        Excel.Workbook wb = ws.Parent;
        Excel.Application app = wb.Application;

        int totalRows = tableRng.Rows.Count - 1; // Exclude header
        int rowsPerPart = (int)Math.Ceiling(totalRows / (double)parts);
        parts = Math.Min(parts, totalRows);

        Excel.Range headerRow = tableRng.Rows[1];

        using (new ExcelExecutionBlock(app))
        {
            for (int i = 0; i < parts; i++)
            {
                Excel.Worksheet newSheet = (Excel.Worksheet)wb.Sheets.Add(After: ws);
                newSheet.Rename(ws.Name, $"_{i + 1}of{parts}");

                // Copy header row to the new sheet.
                headerRow.Copy(Destination: newSheet.Range["A1"]);

                // Calculate the range to copy for this part.
                int startRow = 2 + (i * rowsPerPart); // Skip header, adjust for base-1 index, and consider previous parts.
                int endRow = Math.Min(startRow + rowsPerPart - 1, totalRows + 1); // Ensure not to exceed total rows.

                if (startRow <= endRow)
                {
                    Excel.Range rowsToCopy = ws.Range[tableRng.Cells[startRow, 1], tableRng.Cells[endRow, tableRng.Columns.Count]];
                    rowsToCopy.Copy(Destination: newSheet.Range["A2"]);
                }
            }
        }
    }

    public static bool MoveWorksheets2NewWb(Excel.Application app, List<Excel.Worksheet> wsList)
    {
        SaveFileDialog saveDlg = new SaveFileDialog();

        if (!string.IsNullOrEmpty(app.ActiveWorkbook.Path))
            saveDlg.InitialDirectory = app.ActiveWorkbook.Path;
        else
            saveDlg.InitialDirectory = Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();

        saveDlg.FileName = $"{Path.GetFileNameWithoutExtension(app.ActiveWorkbook.FullName)} (Copy)";
        saveDlg.OverwritePrompt = false;
        saveDlg.DefaultExt = ".xlsb";
        saveDlg.AddExtension = true;
        saveDlg.Filter = "Excel Files|*.xls;*.xlsx;*.xlsb;*.xlsm;*.xltm";
        DialogResult result = saveDlg.ShowDialog();

        Excel.Workbook newWb;
        Excel.Worksheet startWs;

        using (new ExcelExecutionBlock(app))
        {
            newWb = app.Workbooks.Add();
            startWs = newWb.Worksheets[1];
            startWs.Name = "khdbacddkbl";

            foreach (var ws in wsList)
                ws.Copy(startWs);

            app.DisplayAlerts = false;
            startWs.Delete();
            app.DisplayAlerts = true;
        }

        if (result == DialogResult.OK)
        {
            string filePath = saveDlg.FileName;
            string fileExtension = Path.GetExtension(filePath);
            switch (fileExtension.ToLower())
            {
                case ".xls":
                    newWb.SaveAs(filePath, Excel.XlFileFormat.xlWorkbookNormal);
                    break;
                case ".xlsx":
                    newWb.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbook);
                    break;
                case ".xlsb":
                    newWb.SaveAs(filePath, Excel.XlFileFormat.xlExcel12);
                    break;
                case ".xlsm":
                    newWb.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled);
                    break;
                case ".xltm":
                    newWb.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLTemplateMacroEnabled);
                    break;
                default:
                    Console.WriteLine("Invalid file extension.");
                    break;
            }
            return true;
        }
        else
            return false;
    }

    public static void RemoveFormattingFromRange(Excel.Range rng)
    {
        using (new ExcelExecutionBlock(rng.Application))
        {
            rng.ClearFormats();
        }
    }

    public static void ApplyCustomNumberFormat(Excel.Range rng)
    {
        string decimalSeparator;
        string thousandsSeparator;
        string sWRS;
        string format;
        List<Excel.Range> dateColumns;

        string pattern = "(?=.*d)(?=.*m)(?=.*y)";
        sWRS = string.Format("{0:#,##0.00}", 1000);
        thousandsSeparator = sWRS.Substring(1, 1).Replace((char)160, ' ');
        decimalSeparator = sWRS.Substring(5, 1);

        sWRS = string.Format("{0:#,##0.00}", -1);
        using (new ExcelExecutionBlock(rng.Application))
        {
            dateColumns = rng.Columns.Cast<Excel.Range>().Where(p => (p.Cells.SpecialCellsOrDefault(Excel.XlCellType.xlCellTypeConstants) ?? p).Cast<Excel.Range>()
            .Take(5)
            .Any(c => Regex.IsMatch(((c.NumberFormatLocal is string ? c.NumberFormatLocal : string.Empty) as string), pattern))).ToList();

            if (!sWRS.StartsWith("("))
            {
                format = "[Color49]#" + thousandsSeparator + "##0" + decimalSeparator + "00;[Color9]-#" + thousandsSeparator + "##0" + decimalSeparator + "00;[Color16]0;@";
                if ((rng.NumberFormatLocal is string ? rng.NumberFormatLocal : string.Empty) == format)
                    format = "[Color49]#" + thousandsSeparator + "##0;[Color9]-#" + thousandsSeparator + "##0;[Color16]0;@";
                rng.NumberFormatLocal = format;
                if (dateColumns != null && dateColumns.Count > 0)
                    dateColumns.ForEach(p => p.NumberFormatLocal = "[Color10]yyyy.mm.dd;@");
            }
            else
            {
                format = "[Color49]#" + thousandsSeparator + "##0" + decimalSeparator + "00;[Color9](#" + thousandsSeparator + "##0" + decimalSeparator + "00);[Color16]0;@";
                if ((rng.NumberFormatLocal is string ? rng.NumberFormatLocal : string.Empty) == format)
                    format = "[Color49]#" + thousandsSeparator + "##0;[Color9](#" + thousandsSeparator + "##0);[Color16]0;@";
                rng.NumberFormatLocal = format;
                if (dateColumns != null && dateColumns.Count > 0)
                    dateColumns.ForEach(p => p.NumberFormatLocal = "[Color10]yyyy.mm.dd;@");
            }
        }
    }

    public static Excel.Range SpecialCellsOrDefault(this Excel.Range rng, Excel.XlCellType cellType)
    {
        try
        {
            return rng.SpecialCells(cellType);
        }
        catch (COMException)
        {
            return null;
        }
    }

    public static void DeleteOutsideRngOrRegion(Excel.Range selection)
    {
        if (!selection.Valid())
            return;

        Excel.Worksheet activeSheet = selection.Worksheet;
        Excel.Range region;
        if (selection.Cells.Count == 1)
        {
            if (selection.IsPivotCell())
                region = selection.PivotCell.PivotTable.TableRange2;
            else
                region = selection.CurrentRegion;
        }
        else if (selection.Cells.Count > 1)
            region = selection;
        else
            return;

        using (new ExcelExecutionBlock(selection.Application))
        {
            // Delete columns to the left of the region.
            if (region.Column > 1)
            {
                Excel.Range leftColumns = activeSheet.Range[activeSheet.Cells[1, 1], activeSheet.Cells[1, region.Column - 1]].EntireColumn;
                leftColumns.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
            }

            // Delete columns to the right of the region.
            int lastColumn = region.Column + region.Columns.Count - 1;
            if (lastColumn < activeSheet.Columns.Count)
            {
                Excel.Range rightColumns = activeSheet.Range[activeSheet.Cells[1, lastColumn + 1], activeSheet.Cells[1, activeSheet.Columns.Count]].EntireColumn;
                rightColumns.ClearFormats();
                rightColumns.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
            }

            // Delete rows above the region.
            if (region.Row > 1)
            {
                Excel.Range aboveRows = activeSheet.Range[activeSheet.Cells[1, 1], activeSheet.Cells[region.Row - 1, 1]].EntireRow;
                aboveRows.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            }

            // Delete rows below the region.
            int lastRow = region.Row + region.Rows.Count - 1;
            if (lastRow < activeSheet.Rows.Count)
            {
                Excel.Range belowRows = activeSheet.Range[activeSheet.Cells[lastRow + 1, 1], activeSheet.Cells[activeSheet.Rows.Count, 1]].EntireRow;
                belowRows.ClearFormats();
                belowRows.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            }
        }
    }

    public static void SortColumnByAbsoluteValues(Excel.Range selection)
    {
        if (selection == null)
            return;

        Excel.Worksheet activeSheet = selection.Worksheet;
        Excel.Range currentRegion = selection.CurrentRegion;

        if (currentRegion.Rows.Count < 3)
            return;

        int selectedColumnIndex = selection.Column - currentRegion.Column + 1;
        Excel.Range selectedColumn = currentRegion.Columns[selectedColumnIndex];
        Excel.Range selectedColumnData = selectedColumn.Offset[1, 0].Resize[selectedColumn.Rows.Count - 1];

        // Check if the whole column below the header consists of numbers
        if (selectedColumnData.Cells.Cast<Excel.Range>().Any(cell => !double.TryParse(string.IsNullOrWhiteSpace(cell.Value2.ToString()) ? "0" : cell.Value2.ToString(), out double result)))
            return;

        //using (new ExcelExecutionBlock(selection.Application))
        //{
        // Insert temporary column or cells if autofilter is off
        if (activeSheet.AutoFilterMode)
            selectedColumn.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
        else
            selectedColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
        activeSheet.Calculate();

        if (selectedColumnIndex == 1)
        {
            currentRegion = currentRegion.Offset[0, -1].Resize[currentRegion.Rows.Count, currentRegion.Columns.Count + 1];
        }

        Excel.Range tempColumn = currentRegion.Columns[selectedColumnIndex];//.EntireColumn;
        if (activeSheet.AutoFilterMode)
            tempColumn = tempColumn.EntireColumn;
        Excel.Range tempColumnData = currentRegion.Columns[selectedColumnIndex].Offset[1, 0].Resize[currentRegion.Rows.Count - 1];

        // Set formula to absolute value of selected column
        tempColumnData.Formula = $"=ABS({selectedColumnData.Address})";// $"=ABS(RC[{selectedColumnIndex + 1}])";
        tempColumnData.Calculate();

        // Sort by temporary column
        currentRegion.Sort(Key1: currentRegion.Columns[selectedColumnIndex], Order1: Excel.XlSortOrder.xlDescending, Header: Excel.XlYesNoGuess.xlYes);

        tempColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
        //}
    }

    public static DataTable GetDataTable<T>(this T rng, bool dataHasHeaders = true) where T : Excel.Range
    {
        DataTable dt = new DataTable();
        int rowCount = rng.Rows.Count;
        int colCount = rng.Columns.Count;

        object[,] cellValues = (object[,])rng.Value2;

        // Add columns to DataTable
        for (int i = 1; i <= colCount; i++)
        {
            // Determine the type of the first non-empty cell in the column
            Type columnType = typeof(string); // Default to string
            for (int j = dataHasHeaders ? 2 : 1; j <= rowCount; j++)
            {
                if (cellValues[j, i] != null)
                {
                    columnType = cellValues[j, i].GetType();
                    break;
                }
            }
            dt.Columns.Add(dataHasHeaders ? (string.IsNullOrWhiteSpace(cellValues[1, i].ToString()) ? $"Column{i}" : cellValues[1, i].ToString()) : $"Column{i}", columnType);
        }

        // Add rows to DataTable
        for (int i = dataHasHeaders ? 2 : 1; i <= rowCount; i++)
        {
            DataRow newRow = dt.NewRow();
            for (int j = 1; j <= colCount; j++)
            {
                //newRow[j - 1] = rng.Cells[i, j].Value2;
                newRow[j - 1] = cellValues[i, j] ?? DBNull.Value;
            }
            dt.Rows.Add(newRow);
        }

        return dt;
    }

    public static DataTable GetDataTable2<T>(this T rng, bool dataHasHeaders = true) where T : Excel.Range
    {
        DataTable dt = new DataTable();
        int rowCount = rng.Rows.Count;
        int colCount = rng.Columns.Count;

        object[,] cellValues = (object[,])rng.Value2;

        // Add columns to DataTable
        for (int i = 1; i <= colCount; i++)
        {
            // Determine the type of the first non-empty cell in the column
            Type columnType = typeof(string); // Default to string
            for (int j = dataHasHeaders ? 2 : 1; j <= rowCount; j++)
            {
                if (cellValues[j, i] != null)
                {
                    columnType = cellValues[j, i].GetType();
                    break;
                }
            }
            dt.Columns.Add(dataHasHeaders ? (string.IsNullOrWhiteSpace(cellValues[1, i].ToString()) ? $"Column{i}" : cellValues[1, i].ToString()) : $"Column{i}", columnType);
        }

        object lockObject = new object();
        Parallel.For(dataHasHeaders ? 2 : 1, cellValues.GetLength(0) + 1, i =>
        {
            DataRow newRow;
            lock (lockObject)
                newRow = dt.NewRow();

            for (int j = 1; j <= colCount; j++)
                newRow[j - 1] = cellValues[i, j] ?? DBNull.Value;

            lock (lockObject)
                dt.Rows.Add(newRow);
        });

        return dt;
    }

    public static DataTable GetStringDataTable<T>(this T rng, bool dataHasHeaders = true) where T : Excel.Range
    {
        DataTable dt = new DataTable();
        int rowCount = rng.Rows.Count;
        int colCount = rng.Columns.Count;
        List<Vector2> errorsList = new List<Vector2>();
        object[,] cellValues = (object[,])rng.Value2;

        // Add columns to DataTable
        for (int i = 1; i <= colCount; i++)
        {
            dt.Columns.Add(dataHasHeaders ? (string.IsNullOrWhiteSpace(cellValues[1, i].ToString()) ? $"Column{i}" : cellValues[1, i].ToString()) : $"Column{i}", typeof(string));
        }

        int offset = dataHasHeaders ? 1 : 0;

        // Add rows to DataTable
        for (int i = 1; i <= rowCount - offset; i++)
        {
            dt.Rows.Add(dt.NewRow());
        }

        Parallel.For(1, colCount + 1, j =>
        {
            for (int i = 1 + offset; i <= rowCount; i++)
            {
                try
                {
                    dt.Rows[i - offset - 1].SetField(j - 1, cellValues[i, j]?.ToString() ?? string.Empty);
                }
                catch (Exception)
                {
                    errorsList.Add(new Vector2(i, j));
                    //dt.Rows[i - offset - 1].SetField(j - 1, string.Empty);
                }
            }
        });

        foreach (var v in errorsList)
        {
            dt.Rows[(int)v.X - offset - 1].SetField((int)v.Y - 1, cellValues[(int)v.X, (int)v.Y]?.ToString() ?? string.Empty);
        }

        return dt;
    }

    public static bool UpdateMacro(string fileName, Excel.Workbook wb)
    {
        bool deleteMacro = false;
        string[] macroName2Update;
        try
        {
            var fileLines = File.ReadLines(fileName);
            if ((fileLines.FirstOrDefault(p => p.Trim().Length > 0) ?? string.Empty).StartsWith("~"))
            {
                deleteMacro = true;
                macroName2Update = (fileLines.FirstOrDefault(p => p.Trim().StartsWith("~")) ?? string.Empty).Trim('#', ' ').Split('.');
            }
            else
            {
                macroName2Update = (fileLines.FirstOrDefault(p => p.Trim().StartsWith("#")) ?? string.Empty).Trim('#', ' ').Split('.');
            }


            if (macroName2Update.Length != 2)
                return false;

            if (!deleteMacro && !wb.VBProject.VBComponents.Cast<ExcelVB.VBComponent>().Select(p => p.Name).Contains(macroName2Update[0]))
            {
                var vbc = wb.VBProject.VBComponents.Add(ExcelVB.vbext_ComponentType.vbext_ct_StdModule);
                vbc.Name = macroName2Update[0];
                vbc.CodeModule.InsertLines(1, "Option Explicit\n\n" + string.Join("\n", fileLines.SkipWhile(p => p.Trim() == string.Empty).Skip(1)));
                return true;
            }

            var component = wb.VBProject.VBComponents.Item(macroName2Update[0]);
            for (int i = 1; i < wb.VBProject.VBComponents.Item(macroName2Update[0]).CodeModule.CountOfLines; i++)
            {
                string macroName = component.CodeModule.ProcOfLine[i, out ExcelVB.vbext_ProcKind procKind];
                if (!string.IsNullOrWhiteSpace(macroName) && procKind == ExcelVB.vbext_ProcKind.vbext_pk_Proc && macroName == macroName2Update[1])
                {
                    int startLine = component.CodeModule.ProcStartLine[macroName, procKind];
                    component.CodeModule.DeleteLines(startLine, component.CodeModule.ProcCountLines[macroName, procKind]);
                    if (!deleteMacro)
                        component.CodeModule.InsertLines(startLine, "\n" + string.Join("\n", fileLines.SkipWhile(p => p.Trim() == string.Empty).Skip(1)));
                    return true;
                }
            }

            if (deleteMacro)
                return false;

            string lastLine = component.CodeModule.Lines[component.CodeModule.CountOfLines, 1];
            component.CodeModule.ReplaceLine(component.CodeModule.CountOfLines, $"{lastLine.Trim()}\n\n{string.Join("\n", fileLines.SkipWhile(p => p.Trim() == string.Empty).Skip(1))}\n");

            return true;
        }
        catch (Exception)
        {
            return false;
        }
    }

    public static bool UpdateModule(string fileName, Excel.Workbook wb, bool replace = false)
    {
        bool deleteModule = false;
        try
        {
            string name = Path.GetFileNameWithoutExtension(fileName);
            if (name.StartsWith("~"))
            {
                deleteModule = true;
                name = name.TrimStart('~');
            }

            if (!wb.VBProject.VBComponents.Cast<ExcelVB.VBComponent>().Select(p => p.Name).Contains(name))
            {
                if (!deleteModule)
                    wb.VBProject.VBComponents.Import(fileName);
                return true;
            }

            ExcelVB.VBComponent oldComponent = wb.VBProject.VBComponents.Item(name);

            if (!deleteModule && replace)
            {
                wb.VBProject.VBComponents.Remove(oldComponent);
                wb.VBProject.VBComponents.Import(fileName);
                return true;
            }
            else if (deleteModule)
            {
                wb.VBProject.VBComponents.Remove(oldComponent);
                return true;
            }

            string nameOfOldModule = name + "_old";
            string oldBaseName = nameOfOldModule;
            int j = 1;
            while (wb.VBProject.VBComponents.Cast<ExcelVB.VBComponent>().Select(p => p.Name).Contains(nameOfOldModule))
                nameOfOldModule = oldBaseName + j++;


            oldComponent.Name = nameOfOldModule;
            ExcelVB.VBComponent newComponent = wb.VBProject.VBComponents.Import(fileName);

            Dictionary<string, string> oldModuleMacros = new Dictionary<string, string>();
            List<string> newModuleMacros = new List<string>();

            for (int i = 1; i < newComponent.CodeModule.CountOfLines; i++)
            {
                string macroName = newComponent.CodeModule.ProcOfLine[i, out ExcelVB.vbext_ProcKind procKind];
                if (!string.IsNullOrWhiteSpace(macroName))
                {
                    newModuleMacros.Add(macroName);
                    i += newComponent.CodeModule.ProcCountLines[macroName, procKind] - 1;
                }
            }

            for (int i = 1; i < oldComponent.CodeModule.CountOfLines; i++)
            {
                string macroName = oldComponent.CodeModule.ProcOfLine[i, out ExcelVB.vbext_ProcKind procKind];
                if (!string.IsNullOrWhiteSpace(macroName))
                {
                    oldModuleMacros.Add(macroName, oldComponent.CodeModule.Lines[oldComponent.CodeModule.ProcStartLine[macroName, procKind], oldComponent.CodeModule.ProcCountLines[macroName, procKind]]);
                    i += oldComponent.CodeModule.ProcCountLines[macroName, procKind] - 1;
                }
            }

            if (newModuleMacros.OrderBy(p => p) == oldModuleMacros.Select(p => p.Key.ToString()).ToList().OrderBy(p => p))
            {
                wb.VBProject.VBComponents.Remove(oldComponent);
                return true;
            }

            string extraMacros = string.Join("\n\n", oldModuleMacros.Keys.Where(p => !newModuleMacros.Contains(p)).Select(p => oldModuleMacros[p].Trim()).ToList());
            //newComponent.CodeModule.InsertLines(newComponent.CodeModule.CountOfLines, "\n" + extraMacros);
            string lastLine = newComponent.CodeModule.Lines[newComponent.CodeModule.CountOfLines, 1];
            newComponent.CodeModule.ReplaceLine(newComponent.CodeModule.CountOfLines, $"{lastLine.Trim()}\n\n{extraMacros}\n");

            wb.VBProject.VBComponents.Remove(oldComponent);
            return true;
        }
        catch (Exception)
        {
            return false;
        }
    }

    public static void HideRangeWithText(Excel.Range rng, string text, DirectionType type)
    {
        if (!rng.Valid() || string.IsNullOrWhiteSpace(text))
            return;
        using (new ExcelExecutionBlock(rng.Application))
        {
            switch (type)
            {
                case DirectionType.Rows:
                    foreach (var row in rng.Rows.Cast<Excel.Range>())
                    {
                        if (row.Cells.Cast<Excel.Range>().Any(p => p.Value2 != null && (p.Value2.ToString() == text || p.Text.ToString() == text)))
                            row.Hidden = true;
                    }
                    break;
                case DirectionType.Colums:
                    foreach (var col in rng.Columns.Cast<Excel.Range>())
                    {
                        if (col.Cells.Cast<Excel.Range>().Any(p => p.Value2 != null && (p.Value2.ToString() == text || p.Text.ToString() == text)))
                            col.Hidden = true;
                    }
                    break;
                case DirectionType.Cells:
                default:
                    return;
            }
        }
    }

    public enum DirectionType
    {
        Cells,
        Colums,
        Rows
    }

    public static bool Exists<T>(this T worksheet) where T : Excel.Worksheet
    {
        try
        {
            if (worksheet == null)
                return false;

            var testName = worksheet.Name;
            var testRowsCount = worksheet.Rows.Count;
            // If no exception is thrown, then the worksheet is still valid.
            return true;
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // If a COMException is caught, it's likely because the worksheet no longer exists.
            return false;
        }
        catch (Exception)
        {
            return false;
        }
    }

    public static bool Valid<T>(this T rng) where T : Excel.Range
    {
        try
        {
            if (rng == null || rng.Cells.Count < 1)
                return false;

            var testAddress = rng.Address;
            // If no exception is thrown, then the range is still valid.
            return true;
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // If a COMException is caught, it's likely because the range no longer exists.
            return false;
        }
        catch (Exception)
        {
            return false;
        }
    }

    public static SortedDictionary<string, long> GetCounts(Excel.Range rng, string searchWord = "")
    {
        var counts = new SortedDictionary<string, long>();
        object[,] cellValues = (object[,])rng.Value2;
        int column = rng.Column;

        object lockObject = new object();
        Parallel.For(1, cellValues.GetLength(1) + 1, i =>
        {
            string columnName;
            long count = 0;

            columnName = cellValues[1, i].ToString() ?? $"Column{i - column + 1}";
            for (int j = 2; j < cellValues.GetLength(0) + 1; j++)
            {
                string value = cellValues[j, i]?.ToString();
                if (!string.IsNullOrEmpty(value) && (string.IsNullOrEmpty(searchWord) || value.Contains(searchWord, StringComparison.OrdinalIgnoreCase)))
                    ++count;
            }

            // Synchronize access to the counts dictionary
            lock (lockObject)
                counts[columnName] = count;
        });

        return counts;
    }

    public static Excel.Range GetUsableRange<T>(this T rng) where T : Excel.Range
    {
        if (!rng.Valid())
            return null;

        Excel.Range intersection = rng.Application.Intersect(rng, rng.Worksheet.UsedRange);

        if (!intersection.Valid())
            return null;

        return intersection;
    }

    public static void ChangeToText<T>(this T rng) where T : Excel.Range
    {
        if (!rng.Valid())
            return;

        using (new ExcelExecutionBlock(rng.Application))
        {
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
                catch (Exception) { }
            }
        }
    }

    public static void ChangeToValue<T>(this T rng) where T : Excel.Range
    {
        if (!rng.Valid())
            return;

        using (new ExcelExecutionBlock(rng.Application))
        {
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
                catch (Exception) { }
            }
        }
    }

    public static void FillEmptyCellWithAboveValue(Excel.Range rng)
    {
        if (!rng.Valid())
            return;

        int firstRowNumber = rng.Row;

        using (new ExcelExecutionBlock(rng.Application))
        {
            foreach (Excel.Range col in rng.Columns)
            {
                // Loop through each cell in the column
                foreach (Excel.Range cell in col.Cells)
                {
                    // Check if the cell is not the first one in the column and if it's empty or contains only spaces
                    if (cell.Row != firstRowNumber && string.IsNullOrWhiteSpace(cell.Value2?.ToString()))
                    {
                        Excel.Range cellAbove = cell.Offset[-1, 0];
                        if (cellAbove.HasFormula)
                        {
                            // If the cell above contains a formula, use the FillDown method to copy it
                            cellAbove.Copy(cell);
                        }
                        else
                        {
                            cell.NumberFormatLocal = cellAbove.NumberFormatLocal;
                            cell.Value2 = cellAbove.Value2;
                        }
                    }
                }
            }
        }
    }

    public static void SelectCurrentRegionWithoutHeaders(Excel.Range rng)
    {
        if (rng.Valid())
        {
            rng = rng.CurrentRegion;

            if (rng.Rows.Count > 1)
            {
                rng = rng.Offset[1, 0].Resize[rng.Rows.Count - 1, rng.Columns.Count];
                rng.Select();
            }
            else
            {
                MessageBox.Show("The current region does not have more than one row.");
            }
        }
        else
        {
            MessageBox.Show("Please select a range first.");
        }
    }

    public static string ExportMacros(this Excel.Workbook wb, string path = "")
    {
        try
        {
            if (path == string.Empty || !Directory.Exists(path))
            {
                using (var folderBrowserDialog = new FolderBrowserDialog())
                {
                    DialogResult result = folderBrowserDialog.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog.SelectedPath))
                    {
                        string selectedPath = folderBrowserDialog.SelectedPath;
                        string newDirectoryName = Path.GetFileNameWithoutExtension(wb.FullName);
                        string newPath = Path.Combine(selectedPath, newDirectoryName);
                        int i = 1;
                        while (Directory.Exists(newPath))
                        {
                            newPath = Path.Combine(selectedPath, newDirectoryName + $" ({i})");
                        }
                        path = Directory.CreateDirectory(newPath).FullName;
                    }
                }
            }

            foreach (var comp in wb.VBProject.VBComponents.Cast<ExcelVB.VBComponent>().Where(p => p.Type == ExcelVB.vbext_ComponentType.vbext_ct_StdModule || p.Type == ExcelVB.vbext_ComponentType.vbext_ct_ClassModule || p.Type == ExcelVB.vbext_ComponentType.vbext_ct_MSForm))
            {
                comp.Export(Path.Combine(path, comp.Name + ".bas"));
            }

            foreach (var comp in wb.VBProject.VBComponents.Cast<ExcelVB.VBComponent>().Where(p => p.Type == ExcelVB.vbext_ComponentType.vbext_ct_MSForm))
            {
                comp.Export(Path.Combine(path, comp.Name + ".frm"));
            }

            return path;
        }
        catch (COMException)
        {
            return string.Empty;
        }
        catch (Exception)
        {
            return string.Empty;
        }
    }

    public static void CreateMacroUpdateFileFromActiveVbaCode(Excel.Application app, string outputPath)
    {
        ExcelVB.VBE vbe = app.VBE;
        ExcelVB.CodePane activeCodePane = vbe.ActiveCodePane;
        if (activeCodePane != null)
        {
            activeCodePane.GetSelection(out int selectionStartLine, out _, out int selectionEndLine, out _);
            ExcelVB.CodeModule cm = activeCodePane.CodeModule;
            ExcelVB.VBComponent component = cm.Parent;

            var macros = Macro.GetMacrosFromLines(component, selectionStartLine, selectionEndLine);

            if (macros == null || macros.Count < 1)
                MessageBox.Show("No active subroutine or function selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            foreach (var macro in macros)
            {
                string fileName = $"{macro.FullName}.macro";
                string fullPath = Path.Combine(outputPath, fileName);
                File.Create(fullPath).Close();
                File.WriteAllText(fullPath, $"#{macro.FullName}{Environment.NewLine}{macro.Code}");
            }
        }
        else
        {
            MessageBox.Show("VBA Editor is not open or no code pane is active.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
    public static void ColorRange(Excel.Range rng, RangeType type, Color color, string searchWord = "", bool invertFontColor = false)
    {
        if (!rng.Valid())
            return;
        using (new ExcelExecutionBlock(rng.Application))
        {
            switch (type)
            {
                case RangeType.Rows:
                    foreach (var row in rng.Rows.Cast<Excel.Range>())
                    {
                        if (row.Cells.Cast<Excel.Range>().Any(p => p.Value2 != null && (p.Value2.ToString().Contains(searchWord) || p.Text.ToString().Contains(searchWord))))
                        {
                            if (color != Color.Transparent)
                                row.Interior.Color = color;
                            else
                                row.Interior.ColorIndex = 0;

                            if (invertFontColor)
                                row.Font.Color = ColorTranslator.FromOle((int)(row.Cells[1, 1] as Excel.Range).Font.Color).Invert();
                        }
                    }
                    break;
                case RangeType.Colums:
                    foreach (var col in rng.Columns.Cast<Excel.Range>())
                    {
                        if (col.Cells.Cast<Excel.Range>().Any(p => p.Value2 != null && (p.Value2.ToString().Contains(searchWord) || p.Text.ToString().Contains(searchWord))))
                        {
                            if (color != Color.Transparent)
                                col.Interior.Color = color;
                            else
                                col.Interior.ColorIndex = 0;

                            if (invertFontColor)
                                col.Font.Color = ColorTranslator.FromOle((int)(col.Cells[1, 1] as Excel.Range).Font.Color).Invert();
                        }
                    }
                    break;
                case RangeType.Cells:
                    foreach (var cell in rng.Cells.Cast<Excel.Range>())
                    {
                        if (cell.Value2 != null && (cell.Value2.ToString().Contains(searchWord) || cell.Text.ToString().Contains(searchWord)))
                        {
                            if (color != Color.Transparent)
                                cell.Interior.Color = color;
                            else
                                cell.Interior.ColorIndex = 0;

                            if (invertFontColor)
                                cell.Font.Color = ColorTranslator.FromOle((int)(cell.Cells[1, 1] as Excel.Range).Font.Color).Invert();
                        }
                    }
                    break;
                default:
                    return;
            }
        }
    }

    public enum RangeType
    {
        Cells,
        Colums,
        Rows
    }
}

