using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Policy;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelEssentials.Scripts;
using Excel = Microsoft.Office.Interop.Excel;
using WTC = ImportTableToExcel.WorksheetFromTxtCreator;

namespace ExcelEssentials.Forms
{
    public partial class BrowserViewForm : Form
    {
        public string DownloadedFilePath;
        public string Url;
        public bool AutoImport;
        public Dictionary<string, string> Bookmarks;

        public const Int32 WM_SYSCOMMAND = 0x112;
        public const Int32 MF_BYPOSITION = 0x400;
        public const Int32 ToggleTopMostMenuItem = 1000;
        public const Int32 CenterFormMenuItem = 1001;

        [DllImport("user32.dll")]
        private static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
        [DllImport("user32.dll")]
        private static extern bool InsertMenu(IntPtr hMenu, Int32 wPosition, Int32 wFlags, Int32 wIDNewItem, string lpNewItem);


        public BrowserViewForm(bool autoImport, string url = @"https://google.com")
        {
            InitializeComponent();
            webView2.CreationProperties = new Microsoft.Web.WebView2.WinForms.CoreWebView2CreationProperties();
            webView2.CreationProperties.UserDataFolder = @"C:\Temp\ExcelEssentialsPack\";
            Url = url;
            urlTextBox.Text = url;
            AutoImport = autoImport;
            Bookmarks = GetBookmarks();
            bookmarksComboBox.Items.AddRange(Bookmarks.Keys.ToArray());
            Reload();
        }
        private void BrowserViewForm_Load(object sender, EventArgs e)
        {
            CenterToParent();
            IntPtr MenuHandle = GetSystemMenu(Handle, false);
            InsertMenu(MenuHandle, 5, MF_BYPOSITION, ToggleTopMostMenuItem, "Pin/Unpin this window");
            InsertMenu(MenuHandle, 6, MF_BYPOSITION, CenterFormMenuItem, "Center window");
        }

        protected override void WndProc(ref Message msg)
        {
            if (msg.Msg == WM_SYSCOMMAND)
            {
                switch (msg.WParam.ToInt32())
                {
                    case ToggleTopMostMenuItem:
                        ToggleTopMost();
                        return;
                    case CenterFormMenuItem:
                        Utils.MoveFormToCenter(this);
                        return;
                    default:
                        break;
                }
            }
            base.WndProc(ref msg);
        }

        private void ToggleTopMost()
        {
            TopMost = !TopMost;
        }

        private void CoreWebView2_DownloadStarting(object sender, Microsoft.Web.WebView2.Core.CoreWebView2DownloadStartingEventArgs e)
        {
            Microsoft.Web.WebView2.Core.CoreWebView2DownloadOperation download = e.DownloadOperation;
            DownloadedFilePath = e.ResultFilePath;
            download.StateChanged += Download_StateChanged;

        }

        private async void Download_StateChanged(object sender, object e)
        {
            DialogResult result = DialogResult.None;

            var download = sender as Microsoft.Web.WebView2.Core.CoreWebView2DownloadOperation;
            if (download?.State != Microsoft.Web.WebView2.Core.CoreWebView2DownloadState.Completed || !AutoImport)
                return;

            // Marshal everything after this point to the UI thread
            if (this.InvokeRequired)  // 'this' = your Form instance
            {
                this.BeginInvoke(new Action(() => Download_StateChanged(sender, e)));
                return;
            }

            if (download.State == Microsoft.Web.WebView2.Core.CoreWebView2DownloadState.Completed && AutoImport)
            {
                try
                {
                    if (Utils.TextExt.Contains(Path.GetExtension(DownloadedFilePath), StringComparer.OrdinalIgnoreCase))
                    {
                        char delimiter = Utils.DetermineTableDelimiter(DownloadedFilePath);
                        if (delimiter == default(char))
                        {
                            string chosenDelimiter = Microsoft.VisualBasic.Interaction.InputBox("Can not determine delimiter, write it below (write \\t for tab character):", "Write delimiter in ''", "", 0, 0);

                            if (string.IsNullOrWhiteSpace(chosenDelimiter))
                                return;  // User canceled

                            chosenDelimiter = chosenDelimiter.Replace("\\t", "\t");
                            if (chosenDelimiter.Length > 1)
                                chosenDelimiter.Trim();

                            if (chosenDelimiter.Length != 1)
                            {
                                MessageBox.Show("Delimiter too long or missing!", "Delimiter too long or missing!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                if (!FileManager.IsExplorerPathOpen(DownloadedFilePath))
                                    Process.Start("explorer.exe", Path.GetDirectoryName(DownloadedFilePath));
                                result = MessageBox.Show("Close browser?", "Close browser", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                if (result == DialogResult.Yes)
                                    Close();
                                return;
                            }
                            else
                                delimiter = chosenDelimiter[0];
                        }
                        Excel.Application app = Globals.ThisAddIn.Application;
                        Excel.Worksheet aWs = app.ActiveSheet;
                        Excel.Worksheet ws = (aWs.Parent as Excel.Workbook).Worksheets.Add(aWs);
                        ws.Activate();
                        Globals.ThisAddIn.SetDefaultZoom();

                        if (File.ReadLines(DownloadedFilePath).LongCount() > ws.Rows.Count)
                        {
                            int columnCount = 0; long rowCount = File.ReadLines(DownloadedFilePath).LongCount();
                            using (StreamReader reader = new StreamReader(DownloadedFilePath))
                            {
                                string firstLine = reader.ReadLine();
                                columnCount = !string.IsNullOrEmpty(firstLine) ? firstLine.Split(delimiter).Length : 0;
                            }
                            UtilsExcel.RunMacro("LoadTextFileIntoDataModel", new object[] { $"\"{DownloadedFilePath}\"", delimiter.ToString(), columnCount.ToString() });
                            result = MessageBox.Show("Delete file after import?", "Delete file", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (result == DialogResult.Yes)
                            {
                                try
                                {
                                    File.Delete(DownloadedFilePath);
                                }
                                catch (IOException ex) { }
                            }
                            return;
                        }
                        else
                            await WTC.ImportTextFileToExcel(ws, DownloadedFilePath, delimiter);

                        ws.Rename(Path.GetFileNameWithoutExtension(DownloadedFilePath));
                        result = MessageBox.Show("Delete file after import?", "Delete file", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            try
                            {
                                File.Delete(DownloadedFilePath);
                            }
                            catch (IOException ex) { }
                        }
                        return;
                    }
                    else if (Utils.ExcelExt.Contains(Path.GetExtension(DownloadedFilePath), StringComparer.OrdinalIgnoreCase))
                    {
                        if (!UtilsExcel.IsValidExcelFile(DownloadedFilePath))
                        {
                            if (Utils.IsHtmlFile(DownloadedFilePath))
                            {
                                string full = Path.GetFullPath(DownloadedFilePath);
                                string newPath = Path.ChangeExtension(full, ".html");
                                File.Move(full, newPath);
                                var uri = new Uri(newPath).AbsoluteUri;          // -> "file:///C:/path/to/file.html"
                                webView2.CoreWebView2.Navigate(uri);
                                return;
                            }
                            else if (Utils.IsLikelyDelimitedTextFile(DownloadedFilePath, out string suggestedExtension))
                            {
                                string sourcePath = DownloadedFilePath;
                                string destPath = Path.ChangeExtension(sourcePath, suggestedExtension);

                                try
                                {
                                    // Overwrite
                                    if (File.Exists(destPath))
                                        File.Delete(destPath);

                                    File.Move(sourcePath, destPath);
                                    DownloadedFilePath = destPath;
                                }
                                catch (IOException ex)
                                {
                                    if (!FileManager.IsExplorerPathOpen(DownloadedFilePath))
                                        Process.Start("explorer.exe", Path.GetDirectoryName(DownloadedFilePath));
                                    result = MessageBox.Show("Close browser?", "Close browser", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                    if (result == DialogResult.Yes)
                                        Close();
                                    return;
                                }

                                char delimiter = Utils.DetermineTableDelimiter(DownloadedFilePath);
                                if (delimiter == default(char))
                                {
                                    string chosenDelimiter = Microsoft.VisualBasic.Interaction.InputBox("Can not determine delimiter, write it below (write \\t for tab character):", "Write delimiter in ''", "", 0, 0);

                                    if (string.IsNullOrWhiteSpace(chosenDelimiter))
                                        return;  // User canceled

                                    chosenDelimiter = chosenDelimiter.Replace("\\t", "\t");
                                    if (chosenDelimiter.Length > 1)
                                        chosenDelimiter.Trim();

                                    if (chosenDelimiter.Length != 1)
                                    {
                                        MessageBox.Show("Delimiter too long or missing!", "Delimiter too long or missing!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        if (!FileManager.IsExplorerPathOpen(DownloadedFilePath))
                                            Process.Start("explorer.exe", Path.GetDirectoryName(DownloadedFilePath));
                                        result = MessageBox.Show("Close browser?", "Close browser", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                        if (result == DialogResult.Yes)
                                            Close();
                                        return;
                                    }
                                    else
                                        delimiter = chosenDelimiter[0];
                                }

                                Excel.Worksheet activeWs = Globals.ThisAddIn.Application.ActiveSheet;
                                Excel.Worksheet ws = (activeWs.Parent as Excel.Workbook).Worksheets.Add(activeWs);
                                ws.Activate();
                                Globals.ThisAddIn.SetDefaultZoom();

                                if (File.ReadLines(DownloadedFilePath).LongCount() > ws.Rows.Count)
                                {
                                    int columnCount = 0; long rowCount = File.ReadLines(DownloadedFilePath).LongCount();
                                    using (StreamReader reader = new StreamReader(DownloadedFilePath))
                                    {
                                        string firstLine = reader.ReadLine();
                                        columnCount = !string.IsNullOrEmpty(firstLine) ? firstLine.Split(delimiter).Length : 0;
                                    }
                                    UtilsExcel.RunMacro("LoadTextFileIntoDataModel", new object[] { $"\"{DownloadedFilePath}\"", delimiter.ToString(), columnCount.ToString() });
                                    result = MessageBox.Show("Delete file after import?", "Delete file", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                    if (result == DialogResult.Yes)
                                    {
                                        try
                                        {
                                            File.Delete(DownloadedFilePath);
                                        }
                                        catch (IOException ex) { }
                                    }
                                    return;
                                }
                                else
                                    await WTC.ImportTextFileToExcel(ws, DownloadedFilePath, delimiter);

                                ws.Rename(Path.GetFileNameWithoutExtension(DownloadedFilePath));
                                result = MessageBox.Show("Delete file after import?", "Delete file", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                if (result == DialogResult.Yes)
                                {
                                    try
                                    {
                                        File.Delete(DownloadedFilePath);
                                    }
                                    catch (IOException ex) { }
                                }
                                return;
                            }
                            else
                            {
                                if (!FileManager.IsExplorerPathOpen(DownloadedFilePath))
                                    Process.Start("explorer.exe", Path.GetDirectoryName(DownloadedFilePath));
                                result = MessageBox.Show("Close browser?", "Close browser", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                if (result == DialogResult.Yes)
                                    Close();
                                return;
                            }
                        }

                        Excel.Application app = Globals.ThisAddIn.Application;
                        Excel.Worksheet aWs = app.ActiveSheet;
                        Excel.Workbook wb = Microsoft.VisualBasic.Interaction.GetObject(DownloadedFilePath) as Excel.Workbook;

                        int sheetCount = wb.Worksheets.Count;
                        bool addAll = false;

                        // Check if there are multiple worksheets and prompt the user
                        if (sheetCount > 1)
                        {
                            result = MessageBox.Show(
                                $"The workbook {wb.Name} contains multiple worksheets. Do you want to add all worksheets? (Yes = Add All, No = Add First Only)", "Multiple Worksheets",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                            addAll = (result == DialogResult.Yes);
                        }

                        if (addAll)
                        {
                            // Add all worksheets
                            Excel.Worksheet lastSheet = aWs;
                            foreach (Excel.Worksheet ws in wb.Worksheets)
                            {
                                if (ws.IsEmpty())
                                    continue;

                                if (aWs == lastSheet)
                                    ws.Copy(Before: aWs);
                                else
                                    ws.Copy(After: lastSheet);

                                Excel.Worksheet newSheet = app.ActiveSheet as Excel.Worksheet;
                                if (Regex.Match(ws.Name, @"^Sheet[1-9][0-9]*$").Success)
                                {
                                    newSheet.Rename(Path.GetFileNameWithoutExtension(wb.FullName), ws.Name);
                                }
                                lastSheet = newSheet;
                            }
                        }
                        else
                        {
                            // Add only the first not empty worksheet
                            Excel.Worksheet ws = wb.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(p => !p.IsEmpty());

                            if (ws != null)
                            {
                                ws.Copy(Before: aWs);
                                Excel.Worksheet newSheet = app.ActiveSheet as Excel.Worksheet;
                                if (Regex.Match(ws.Name, @"^Sheet[1-9][0-9]*$").Success)
                                {
                                    newSheet.Rename(Path.GetFileNameWithoutExtension(wb.FullName));
                                }
                            }
                        }

                        wb.Close();
                        result = MessageBox.Show("Delete file after import?", "Delete file", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);

                        if (result == DialogResult.Yes)
                        {
                            try
                            {
                                File.Delete(DownloadedFilePath);
                            }
                            catch (IOException ex) { }
                        }

                        return;
                    }
                    if (!FileManager.IsExplorerPathOpen(DownloadedFilePath))
                        Process.Start("explorer.exe", Path.GetDirectoryName(DownloadedFilePath));
                    //result = MessageBox.Show("Close browser?", "Close browser", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    //if (result == DialogResult.Yes)
                    //    this.Close();
                    return;
                }
                catch
                {
                    if (!FileManager.IsExplorerPathOpen(DownloadedFilePath))
                        Process.Start("explorer.exe", Path.GetDirectoryName(DownloadedFilePath));
                    result = MessageBox.Show("Close browser?", "Close browser", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                        Close();
                }
            }
            else if (download.State == Microsoft.Web.WebView2.Core.CoreWebView2DownloadState.Completed && !FileManager.IsExplorerPathOpen(DownloadedFilePath))
                Process.Start("explorer.exe", Path.GetDirectoryName(DownloadedFilePath));
        }

        private async void Reload()
        {
            try
            {
                webView2.CoreWebView2.DownloadStarting -= CoreWebView2_DownloadStarting;
            }
            catch (Exception) { }

            await webView2.EnsureCoreWebView2Async(null);
            webView2.CoreWebView2.DownloadStarting += CoreWebView2_DownloadStarting;
            if (IsValidUrl(Url))
            {
                webView2.CoreWebView2.Navigate(Url);
                webView2.CoreWebView2.NewWindowRequested += CoreWebView2_NewWindowRequested;
            }
            else
            {
                MessageBox.Show($"URL: \"{Url}\" is not valid or complete", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CoreWebView2_NewWindowRequested(object sender, Microsoft.Web.WebView2.Core.CoreWebView2NewWindowRequestedEventArgs e)
        {
            /*e.OriginalSourceFrameInfo.
            e.NewWindow.DownloadStarting += CoreWebView2_DownloadStarting;*/
            e.Handled = true;
            webView2.CoreWebView2.Navigate(e.Uri);
        }

        private void webView21_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Modifiers == Keys.Control && e.KeyCode == Keys.R)
                Reload();
        }

        private void searchButton_Click(object sender, EventArgs e)
        {
            string url = string.IsNullOrWhiteSpace(urlTextBox.Text) ? "www.google.com" : urlTextBox.Text;

            // Check if URL starts with "www" and add "https://" prefix if needed
            if (url.StartsWith("www", StringComparison.OrdinalIgnoreCase))
            {
                url = @"https://" + url;
            }

            Url = url;
            Reload();

        }

        private bool IsValidUrl(string url)
        {
            return Uri.IsWellFormedUriString(url, UriKind.Absolute);
        }

        private void urlTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                searchButton.PerformClick();
        }

        public static Dictionary<string, string> GetBookmarks()
        {
            List<string> lines = File.ReadLines(Path.Combine(FileManager.PropertiesFilesPath, "WebsitesListMap.txt")).Select(p => p.Trim()).SkipWhile(p => p == string.Empty).ToList();

            if (lines.Count < 1)
                return null;

            Dictionary<string, string> bookmarksD = new Dictionary<string, string>();

            foreach (string line in lines)
            {
                string[] parts = line.Split('~');
                bookmarksD.Add(parts[0], parts.Length > 1 ? parts[1] : string.Empty);
            }

            return bookmarksD.SkipWhile(p => p.Value == string.Empty).ToDictionary(p => p.Key, p => p.Value);
        }

        private void bookmarksComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Bookmarks.TryGetValue(bookmarksComboBox.Text, out string bookmarkUrl))
                urlTextBox.Text = bookmarkUrl;
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            if (webView2.CoreWebView2 == null)
            {
                MessageBox.Show("Browser engine is not ready.", "Error");
                return;
            }

            string url = webView2.CoreWebView2.Source;
            if (string.IsNullOrEmpty(url))
                return;

            if (Bookmarks.Values.Contains(url))
            {
                MessageBox.Show($"Bookmarks already contain that link\n{url}", "Not saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!IsValidUrl(url))
            {
                MessageBox.Show($"Link is not valid\n{url}", "Not saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string name = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(new Uri(url).Host.Split('.').FirstOrDefault() ?? new Uri(url).Host);
            string tempName = name;
            int x = 1;
            while (Bookmarks.ContainsKey(tempName))
                tempName = name + x++;
            name = tempName;
            InputBoxForm inputBox = new InputBoxForm("Choose unique bookmark name", "Bookmark name", name);
            inputBox.Show(this);
            inputBox.FormClosing += (s, ce) =>
            {
                if (inputBox.DialogResult != DialogResult.OK)
                    return;
                tempName = inputBox.TextBoxText.Replace("~", "-").Trim();
                if (Bookmarks.ContainsKey(tempName))
                {
                    MessageBox.Show($"Bookmarks already contain that name\n{url}", "Name not unique", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ce.Cancel = true;
                }
            };

            inputBox.FormClosed += (s, ce) =>
            {
                if (inputBox.Result != null)
                {
                    name = inputBox.Result.Replace("~", "-").Trim();
                    var urlList = new List<string>
                    {
                        $"{name}~{url}"
                    };
                    try
                    {
                        File.AppendAllLines(Path.Combine(FileManager.PropertiesFilesPath, "WebsitesListMap.txt"), urlList);
                        GetBookmarks();
                        bookmarksComboBox.Items.Clear();
                        bookmarksComboBox.Items.AddRange(Bookmarks.Keys.ToArray());
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show($"Could not save to file: {ex.Message}");
                    }
                }
            };
        }
    }
}
