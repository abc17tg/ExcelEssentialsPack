using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelVB = Microsoft.Vbe.Interop;

namespace ExcelAddInByMarcinOlszewski.Scripts
{
    internal class Macro
    {
        public string Name;
        public string ModuleName;
        public string Code;
        public string FirstCodeLine => Code.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "";
        public string FullName => $"{ModuleName}.{Name}";

        public Macro() { }

        public Macro(string name, string moduleName, string code)
        {
            Name = name;
            ModuleName = moduleName;
            Code = code;
        }

        public static bool Exists(string macroName, string moduleName, Excel.Workbook wb)
        {
            try
            {

                var component = wb.VBProject.VBComponents.Item(moduleName);
                if (component != null && component.Type == ExcelVB.vbext_ComponentType.vbext_ct_StdModule)
                    for (int i = 1; i < component.CodeModule.CountOfLines; i++)
                        if (macroName == component.CodeModule.ProcOfLine[i, out ExcelVB.vbext_ProcKind procKind])
                            return true;

                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static string GetMacroNameForButton(string btnId, Excel.Workbook wb)
        {
            try
            {
                var xe = XElement.Load(Path.Combine(FileManager.PropertiesFilesPath, "ButtonSubroutineMapping.xml"));
                var mapping = xe.Elements("Mapping").FirstOrDefault(m => (string)m.Element("ButtonID") == btnId);
                string macroModuleName = mapping?.Element("Subroutine")?.Value;
/*                string[] temp = macroModuleName.Split('.');
                string macroName = temp[1];
                string module = temp[0];

                if (Exists(macroName, module, wb))*/
                    return macroModuleName;
/*                else
                    return null;*/
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
