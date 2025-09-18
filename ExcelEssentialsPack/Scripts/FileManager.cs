using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Shell32;

namespace ExcelEssentials.Scripts
{

    internal class FileManager
    {
        public static string BasePath => AppDomain.CurrentDomain.BaseDirectory;
        public static string PropertiesFilesPath => Path.Combine(BasePath, "Properties Files");
        public static string ResourcesPath => Path.Combine(BasePath, "Resources");
        public static string DownloadsPath => Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();

#if DEBUG
        public static string MacrosWbName = "MyPERSONAL.xlsb";
        public static string FunctionsWbName = "MyFunctions.xlam";
#else
        public static string MacrosWbName = "MyPERSONAL_Pv1.0.xlsb";
        public static string FunctionsWbName = "MyFunctions_Pv1.0.xlam";
#endif

        public static void CheckForCustomMacrosWbNames()
        {
            try
            {
                string filePath = Path.Combine(PropertiesFilesPath, "MacrosWbName.txt");
                if (File.Exists(filePath))
                {
                    string text = File.ReadAllText(filePath);
                    if (!string.IsNullOrWhiteSpace(text))
                        MacrosWbName = text.Trim();
                    else
                    {
                        if (!IsExplorerPathOpen(PropertiesFilesPath))
                            Process.Start("explorer.exe", PropertiesFilesPath);
                    }
                }
            }
            catch { }

            try
            {
                string filePath = Path.Combine(PropertiesFilesPath, "FunctionsWbName.txt");
                if (File.Exists(filePath))
                {
                    string text = File.ReadAllText(filePath);
                    if (!string.IsNullOrWhiteSpace(text))
                        FunctionsWbName = text.Trim();
                    else
                    {
                        if (!IsExplorerPathOpen(PropertiesFilesPath))
                            Process.Start("explorer.exe", PropertiesFilesPath);
                    }
                }
            }
            catch { }
        }

        public static bool IsExplorerPathOpen(string path)
        {
            Shell shell = new Shell();
            var windows = shell.Windows();
            for (int i = 0; i < windows.Count; i++)
            {
                var window = windows.Item(i);
                if (window != null && window.Path.ToLower() == path.ToLower())
                {
                    return true;
                }
            }
            return false;
        }

        public static void CopyDirectory(string sourceDir, string targetDir)
        {
            // Copy each file in the directory.
            foreach (string filePath in Directory.GetFiles(sourceDir))
            {
                string fileName = Path.GetFileName(filePath);
                string targetFilePath = Path.Combine(targetDir, fileName);
                File.Copy(filePath, targetFilePath, overwrite: true);
            }

            // Recursively copy each subdirectory.
            foreach (string directoryPath in Directory.GetDirectories(sourceDir))
            {
                string directoryName = Path.GetFileName(directoryPath);
                string targetSubDir = Path.Combine(targetDir, directoryName);
                Directory.CreateDirectory(targetSubDir);
                CopyDirectory(directoryPath, targetSubDir);
            }
        }

        public static long GetDirectorySize(string directoryPath)
        {
            long totalSize = 0;
            // Get all files in the directory and subdirectories.
            foreach (string filePath in Directory.GetFiles(directoryPath, "*", SearchOption.AllDirectories))
            {
                FileInfo fi = new FileInfo(filePath);
                totalSize += fi.Length;
            }
            return totalSize;
        }

        public static string GetPathByDialog(string initialName = "", string initialDirectory = "", string filter = "Text Files | *.txt", string defaultExt = ".txt")
        {
            SaveFileDialog saveDlg = new SaveFileDialog();

            if (!string.IsNullOrEmpty(initialDirectory))
                saveDlg.InitialDirectory = initialDirectory;
            else
                saveDlg.InitialDirectory =

            saveDlg.FileName = initialName;
            saveDlg.OverwritePrompt = true;
            saveDlg.DefaultExt = defaultExt;
            saveDlg.AddExtension = true;
            saveDlg.Filter = filter;

            if (saveDlg.ShowDialog() == DialogResult.OK)
                return saveDlg.FileName;
            else
                return null;
        }

        public static string GetValidFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return null;

            // Define the invalid characters for Windows file names
            char[] invalidChars = Path.GetInvalidFileNameChars();

            // Replace invalid characters with an underscore
            string sanitized = Regex.Replace(fileName, $"[{Regex.Escape(new string(invalidChars))}]", "_");

            return sanitized;
        }

        public static void OpenStringWithNotepad(string text)
        {
            // Create a temporary file and write the text to it
            string path = string.Empty;
            try
            {
                path = Path.GetTempFileName();
                File.WriteAllText(path, text);
            }
            catch
            {
                MessageBox.Show("Error occured while creating temporary file to open in notepad!\nOpen and paste contents manually if you still wanna open current document in notepad");
                return;
            }

            try
            {
                Process.Start(Path.Combine(Environment.GetEnvironmentVariable("programfiles"), @"Sublime Text\sublime_text.exe"), $"\"{path}\"");
            }
            catch (Exception)
            {
                try
                {
                    Process.Start(Path.Combine(Environment.GetEnvironmentVariable("programfiles"), @"Notepad++\notepad++.exe"), $"\"{path}\"");
                }
                catch (Exception)
                {
                    try
                    {
                        Process.Start(Path.Combine(Environment.GetEnvironmentVariable("programfiles(x86)"), @"Notepad++\notepad++.exe"), $"\"{path}\"");
                    }
                    catch (Exception)
                    {
                        Process.Start("notepad.exe", $"\"{path}\"");
                    }
                }
            }
        }
    }
}
