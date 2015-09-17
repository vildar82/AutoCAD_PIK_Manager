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
      public Variable TemplatePath { get; set; }
      public Variable SheetSetTemplatePath { get; set; }
      public Variable QNewTemplateFile { get; set; }
      public Variable PageSetupOverridesTemplateFile { get; set; }
   }
}