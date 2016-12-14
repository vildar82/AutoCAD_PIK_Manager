using System;
using System.Collections.Generic;

namespace AutoCAD_PIK_Manager.Settings
{
    public class PathVariable
    {
        public List<Variable> Supports { get; set; }
        public List<Variable> PrinterConfigPaths { get; set; }
        public List<Variable> PrinterDescPaths { get; set; }
        public List<Variable> PrinterPlotStylePaths { get; set; }
        public List<Variable> ToolPalettePaths { get; set; }
        public List<Variable> ColorBookPaths { get; set; }
        public Variable TemplatePath { get; set; }
        public Variable SheetSetTemplatePath { get; set; }        
        public Variable QNewTemplateFile { get; set; }
        public Variable PageSetupOverridesTemplateFile { get; set; }

        /// <summary>
        /// Объединение настроек
        /// </summary>        
        public static PathVariable Merge(PathVariable vars1, PathVariable vars2)
        {
            if (vars1 == null) return vars2;
            if (vars2 == null) return vars1;

            vars1.Supports.AddRange(vars2.Supports);
            vars1.ToolPalettePaths.AddRange(vars2.ToolPalettePaths);
            return vars1;
        }
    }
}