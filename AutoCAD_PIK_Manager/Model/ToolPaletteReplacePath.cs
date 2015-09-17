using System;
using System.IO;
using System.Text.RegularExpressions;
using AutoCAD_PIK_Manager.Settings;

namespace AutoCAD_PIK_Manager.Model
{
   // Замена путей к настройкам в файлах инструментальных палитр atc
   internal static class ToolPaletteReplacePath
   {
      public static void Replace()
      {
         string dirPalette = Path.Combine(PikSettings.LocalSettingsFolder, "ToolPalette");
         if (Directory.Exists(dirPalette))
         {
            try
            {
               var filesAtc = Directory.GetFiles(dirPalette, "*.atc", SearchOption.AllDirectories);
               foreach (var file in filesAtc)
               {
                  ReplacePathInATC(file);
               }
            }
            catch (Exception ex)
            {
               Log.Error("Замена путей в инструментальных палитрах.", ex);
            }
         }
      }

      private static void ReplacePathInATC(string file)
      {
         string content = string.Empty;
         using (StreamReader reader = File.OpenText(file))
         {
            content = reader.ReadToEnd();
         }
         string search = "C:\\Autodesk\\AutoCAD\\Pik\\Settings";
         string replace = PikSettings.LocalSettingsFolder;

         content = ReplaceCaseInsensitive(content, search, replace);

         search = "C:/Autodesk/AutoCAD/Pik/Settings";
         replace = PikSettings.LocalSettingsFolder.Replace("\\", "/");
         content = ReplaceCaseInsensitive(content, search, replace);

         using (StreamWriter stream = new StreamWriter(file))
         {
            stream.Write(content);
         }
      }

      private static string ReplaceCaseInsensitive(string input, string search, string replacement)
      {
         string result = Regex.Replace(
             input,
             Regex.Escape(search),
             replacement.Replace("$", "$$"),
             RegexOptions.IgnoreCase
         );
         return result;
      }
   }
}