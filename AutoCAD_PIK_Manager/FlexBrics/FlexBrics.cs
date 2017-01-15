using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using AutoCadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCAD_PIK_Manager.Settings
{
    public static class FlexBrics
    {        
        public const string FbName = "flexBrics";        

        public static bool HasFlexBrics()
        {
            return PikSettings.GroupFileSettings?.FlexBricsSetup == true;
        }

        public static string GetFBLocalDir()
        {
            return Path.Combine(PikSettings.LocalSettingsFolder, FbName);
        }

        public static string GetServerFlexBricsServerFolder ()
        {
            return Path.GetFullPath(Path.Combine(PikSettings.ServerSettingsFolder, @"..\flexBrics"));            
        }        

        public static void Setup ()
        {
            // Установка flexBrics
            // 1. Добавить папку в доверенные
            var fbLocalDir = GetFBLocalDir();
            if (Directory.Exists(fbLocalDir))
            {
                if (isAcadVerLater2013())
                {
                    string trustedPath = AutoCadApp.GetSystemVariable("TRUSTEDPATHS").ToString();
                    if (!trustedPath.ToLower().Contains(fbLocalDir.ToLower()))
                    {
                        trustedPath += ";" + fbLocalDir + @"\...";
                        AutoCadApp.SetSystemVariable("TRUSTEDPATHS", trustedPath);
                    }
                    try
                    {
                        Log.Info("FlexBrics.Setup. trustedPath ={0}", trustedPath);
                    }
                    catch { }
                }

                // 2. Добавить в пути поиска
                dynamic preference = AutoCadApp.Preferences;
                string supPath = preference.Files.SupportPath;
                supPath = AddPath(fbLocalDir, supPath);//Папка flexBrics
                supPath = AddPath(Path.Combine(fbLocalDir, "dwg"), supPath);//папка dwg
                preference.Files.SupportPath = supPath;
                try
                {
                    Log.Info("FlexBrics.Setup. SupportPath ={0}", supPath);
                }
                catch { }
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

        private static void CopyAll (DirectoryInfo source, DirectoryInfo target, CancellationToken token)
        {
            if (Directory.Exists(target.FullName) == false)
                Directory.CreateDirectory(target.FullName);

            foreach (FileInfo fi in source.GetFiles())
            {
                token.ThrowIfCancellationRequested();
                try
                {
                    fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
                }
                catch { }
            }
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                token.ThrowIfCancellationRequested();
                var nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
                CopyAll(diSourceSubDir, nextTargetSubDir, token);
            }
        }
    }
}