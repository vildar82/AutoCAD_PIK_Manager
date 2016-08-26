using System;
using System.IO;
using AutoCadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCAD_PIK_Manager.Settings
{
    public static class FlexBrics
    {
        private static string _fbLocalDir;

        public static void Copy ()
        {
            if (PikSettings.GroupFileSettings?.FlexBricsSetup == true)
            {
                try
                {
                    var sourceFB = new DirectoryInfo(GetServerFlexBricsServerFolder());
                    _fbLocalDir = Path.Combine(PikSettings.LocalSettingsFolder, sourceFB.Name);
                    var targetFB = new DirectoryInfo(_fbLocalDir);
                    CopyAll(sourceFB, targetFB);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "FlexBrics не скопирровался.");
                }
            }
        }

        private static string GetServerFlexBricsServerFolder ()
        {
            var res = Path.GetFullPath(Path.Combine(PikSettings.ServerSettingsFolder, @"..\flexBrics"));
            return res;
        }        

        public static void Setup ()
        {
            // Установка flexBrics
            // 1. Добавить папку в доверенные
            if (Directory.Exists(_fbLocalDir))
            {
                if (isAcadVerLater2013())
                {
                    string trustedPath = AutoCadApp.GetSystemVariable("TRUSTEDPATHS").ToString();
                    trustedPath += ";" + _fbLocalDir + @"\...";
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

        private static bool isAcadVerLater2013 ()
        {
            Version acadVer = new Version(AutoCadApp.Version.Major, AutoCadApp.Version.Minor);
            Version acad2013Ver = new Version(19, 0);
            return acadVer > acad2013Ver;
        }

        private static string AddPath (string var, string path)
        {
            if (!path.ToUpper().Contains(var.ToUpper()))
            {
                return string.Format("{0};{1}", var, path);
            }
            return path;
        }

        private static void CopyAll (DirectoryInfo source, DirectoryInfo target)
        {
            if (Directory.Exists(target.FullName) == false)
                Directory.CreateDirectory(target.FullName);

            foreach (FileInfo fi in source.GetFiles())
            {
                try
                {
                    fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
                }
                catch { }
            }
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                var nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
                CopyAll(diSourceSubDir, nextTargetSubDir);
            }
        }
    }
}