using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace ExcelEssentials.Scripts
{
    internal class Sde
    {
        public string Query;
        public int Instances;
        public List<string> Items;

        private string formattedItems => $"('{string.Join("', '", Items)}')";
        private string runString => $"-t=\"{formattedItems}\" -q=\"{Query}\" -i=\"{Instances.ToString()}\" -r=\"true\" -h=\"true\" -c=\"false\"";

        //  ./SDE Launcher 2.0.exe -t="[values]" -i="[values]" -q="[values]" -r="[values]" -h="[values]" -c="[values]"
        public Sde(string query, int instances, List<string> items)
        {
            Query = query;
            Instances = instances;
            Items = items;
        }

        public void Run()
        {
            string sde = SdePath();
            if (sde == null)
                return;

            if (Instances > 0 && Instances < 51 && !string.IsNullOrWhiteSpace(Query) && Items != null && Items.Count > 0)
            {
                Process process = new Process();
                process.StartInfo.FileName = sde;
                process.StartInfo.Arguments = runString;
                process.Start();
            }
        }

        public static string SdePath()
        {
            string path = File.ReadLines(Path.Combine(FileManager.PropertiesFilesPath, "SdePath.txt")).FirstOrDefault();
            if (string.IsNullOrWhiteSpace(path))
                return null;

            if (Directory.Exists(path))
                path = Path.Combine(path, "SDE Launcher 2.0.exe");
            else
                return null;

            if (File.Exists(path))
                return path;
            else
                return null;
        }
        
        public static List<string> SdeQueries()
        {
            string path = File.ReadLines(Path.Combine(FileManager.PropertiesFilesPath, "SdePath.txt")).FirstOrDefault();
            if (string.IsNullOrWhiteSpace(path))
                return null;

            if (Directory.Exists(path))
                path = Path.Combine(path, "Data", "SpecificFiles");
            else
                return null;

            if (Directory.Exists(path))
                return Directory.EnumerateDirectories(path).Select(p=> p.Substring(p.LastIndexOf(Path.DirectorySeparatorChar) + 1)).ToList();
            else
                return null;
        }
    }
}
