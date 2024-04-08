using System;
using System.Linq;

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

    }
}
