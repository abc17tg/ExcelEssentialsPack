using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelEssentials.Scripts;
using System.Windows.Forms;
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
            this.CenterToParent();
            IntPtr MenuHandle = GetSystemMenu(this.Handle, false);
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
            this.TopMost = !this.TopMost;
        }

        private void CoreWebView2_DownloadStarting(object sender, Microsoft.Web.WebView2.Core.CoreWebView2DownloadStartingEventArgs e)
        {
            Microsoft.Web.WebView2.Core.CoreWebView2DownloadOperation download = e.DownloadOperation;
            DownloadedFilePath = e.ResultFilePath;
            download.StateChanged += Download_StateChanged;

        }

        private void Download_StateChanged(object sender, object e)
        {
            DialogResult result = DialogResult.None;
            if ((sender as Microsoft.Web.WebView2.Core.CoreWebView2DownloadOperation).State == Microsoft.Web.WebView2.Core.CoreWebView2DownloadState.Completed && AutoImport)
            {
                try
                {
                    if (Utils.TextExt.Contains(Path.GetExtension(DownloadedFilePath), StringComparer.OrdinalIgnoreCase))
                    {
                        char delimiter = Utils.DetermineTableDelimiter(DownloadedFilePath);
                        if (delimiter == default(char))
                        {
                            string choosenDelimiter = Microsoft.VisualBasic.Interaction.InputBox("Can not determine delimiter, write one in ' characters:", "Write delimiter in ''", "", 0, 0);

                            if (choosenDelimiter.Length != 1)
                            {
                                MessageBox.Show("Delimiter too long or missing!", "Delimiter too long or missing!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                Process.Start("explorer.exe", Path.GetDirectoryName(DownloadedFilePath));
                                result = MessageBox.Show("Close browser?", "Close browser", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                if (result == DialogResult.Yes)
                                    this.Close();
                                return;
                            }
                            else
                                delimiter = choosenDelimiter[0];
                        }
                        Excel.Application app = Globals.ThisAddIn.Application;
                        Excel.Worksheet aWs = app.ActiveSheet;
                        Excel.Worksheet ws = (aWs.Parent as Excel.Workbook).Worksheets.Add(aWs);

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
                                File.Delete(this.DownloadedFilePath);
                            //result = MessageBox.Show("Close browser?", "Close browser", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            //if (result == DialogResult.Yes)
                            //    this.Close();
                            return;
                        }
                        else
                            WTC.ImportTextFileToExcel(ws, DownloadedFilePath, delimiter);

                        ws.Rename(Path.GetFileNameWithoutExtension(DownloadedFilePath));
                        result = MessageBox.Show("Delete file after import?", "Delete file", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                            File.Delete(this.DownloadedFilePath);
                        //result = MessageBox.Show("Close browser?", "Close browser", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        //if (result == DialogResult.Yes)
                        //    this.Close();
                        return;
                    }
                    else if (Utils.ExcelExt.Contains(Path.GetExtension(this.DownloadedFilePath), StringComparer.OrdinalIgnoreCase))
                    {
                        Excel.Workbook wb = Microsoft.VisualBasic.Interaction.GetObject(this.DownloadedFilePath) as Excel.Workbook;
                        Excel.Application app = Globals.ThisAddIn.Application;
                        Excel.Worksheet aWs = app.ActiveSheet;
                        wb.Worksheets.Item[1].Copy(aWs);
                        if ((wb.Worksheets.Item[1].Name as string).StartsWith("Sheet"))
                            (app.ActiveSheet as Excel.Worksheet).Rename(Path.GetFileNameWithoutExtension(wb.FullName));
                        //Utils.RunMacro("RenameSheet", new object[] { Path.GetFileNameWithoutExtension(wb.FullName) });
                        wb.Close();
                        result = MessageBox.Show("Delete file after import?", "Delete file", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                            File.Delete(this.DownloadedFilePath);
                        //result = MessageBox.Show("Close browser?", "Close browser", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        //if (result == DialogResult.Yes)
                        //    this.Close();
                        return;
                    }
                    Process.Start("explorer.exe", Path.GetDirectoryName(DownloadedFilePath));
                    //result = MessageBox.Show("Close browser?", "Close browser", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    //if (result == DialogResult.Yes)
                    //    this.Close();
                    return;
                }
                catch
                {
                    Process.Start("explorer.exe", Path.GetDirectoryName(DownloadedFilePath));
                    result = MessageBox.Show("Close browser?", "Close browser", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                        this.Close();
                }
            }
            else if ((sender as Microsoft.Web.WebView2.Core.CoreWebView2DownloadOperation).State == Microsoft.Web.WebView2.Core.CoreWebView2DownloadState.Completed)
                Process.Start("explorer.exe", Path.GetDirectoryName(DownloadedFilePath));
        }

        private async void Reload()
        {
            try
            {
                webView2.CoreWebView2.DownloadStarting -= CoreWebView2_DownloadStarting;
            }
            catch { }

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
            string url = string.IsNullOrWhiteSpace(urlTextBox.Text) ? "www.gooogle.com" : urlTextBox.Text;

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
    }
}
