using System;
using System.IO;
using AutoCadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCAD_PIK_Manager.Settings
{
   internal static class FlexBrics
   {      
      private static string _fbLocalDir;

      internal static void Copy()
      {
         if (PikSettings.GroupFileSettings?.FlexBricsSetup == true)
         {
            try
            {
               var sourceFB = new DirectoryInfo(PikSettings.GroupFileSettings.FlexBricsFolder);
               _fbLocalDir = Path.Combine(PikSettings.LocalSettingsFolder, sourceFB.Name);
               var targetFB = new DirectoryInfo(_fbLocalDir);
               CopyAll(sourceFB, targetFB);
            }
            catch (Exception ex)
            {
               Log.Error("FlexBrics не скопирровался. ", ex);
            }
         }
      }

      internal static void Setup()
      {
         // Установка flexBrics
         // 1. Добавить папку в доверенные
         if (Directory.Exists(_fbLocalDir))
         {
            if (isAcadVerLater2013())
            {
               string trustedPath = AutoCadApp.GetSystemVariable("TRUSTEDPATHS").ToString();
               trustedPath += ";" + _fbLocalDir + "...";
               AutoCadApp.SetSystemVariable("TRUSTEDPATHS", trustedPath);
               Log.Info("FlexBrics.Setup. trustedPath ={0}", trustedPath);
            }

            // 2. Добавить в пути поиска
            dynamic preference = AutoCadApp.Preferences;
            string supPath = preference.Files.SupportPath;
            supPath = AddPath(_fbLocalDir, supPath);//Папка flexBrics
            supPath = AddPath(Path.Combine(_fbLocalDir, "dwg"), supPath);//папка dwg
            preference.Files.SupportPath = supPath;
            Log.Info("FlexBrics.Setup. SupportPath ={0}", supPath);
         }
      }

      private static bool isAcadVerLater2013()
      {
         return !(AutoCadApp.Version.Major == 19 && AutoCadApp.Version.Minor == 0);
      }

      private static string AddPath(string var, string path)
      {
         if (!path.ToUpper().Contains(var.ToUpper()))
         {
            return string.Format("{0};{1}",var, path);
         }
         return path;
      }

      private static void CopyAll(DirectoryInfo source, DirectoryInfo target)
      {         
         if (Directory.Exists(target.FullName) == false)         
            Directory.CreateDirectory(target.FullName);         
                  
         foreach (FileInfo fi in source.GetFiles())
         {            
            fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
         }         
         foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
         {
            var nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
            CopyAll(diSourceSubDir, nextTargetSubDir);
         }
      }      
   }
}
