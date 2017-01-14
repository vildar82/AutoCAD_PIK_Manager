using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AutoCAD_PIK_Manager
{
    public static class Update
    {
        private const string CommonSettingsName = "Общие";

        public static void UpdateSettings()
        {
            try
            {
                var token = new CancellationTokenSource();
                var task = Task.Run(() =>
                {                    
                    // Проверка доступности сетевых настроек
                    if (Directory.Exists(Settings.PikSettings.ServerSettingsFolder))
                    {
                        // Копирование общих настроек
                        CopyCommonFiles(token.Token);
                        // Копирование настроек отдела
                        CopyUserGroupFiles(token.Token);
                    }
                    else
                    {
                        try
                        {
                            Log.Error($"Недоступна папка настроек на сервере {Settings.PikSettings.ServerSettingsFolder}");
                        }
                        catch { }
                    }
                }, token.Token);
                task.Wait(new TimeSpan(0, 100, 0));
                if (!task.IsCompleted)
                {
                    token.Cancel(true);
                }
            }
            catch { }
            // Копирование flexBrics если нужно
            Settings.FlexBrics.Copy();
        }

        private static void CopyUserGroupFiles(CancellationToken token)
        {
            foreach (var group in Settings.PikSettings.UserGroupsCombined)
            {
                // Проверка версии настроек отдела (UserGroup)
                if (VersionsEqal(Path.Combine(Settings.PikSettings.LocalSettingsFolder, group + ".ver"),
                    Path.Combine(Settings.PikSettings.ServerSettingsFolder, $@"{group}\{group}.ver")))
                    return;
                // Копирование настроек с сервера в локальную папку Settings
                var serverUserGroupDir = new DirectoryInfo(Path.Combine(Settings.PikSettings.ServerSettingsFolder, group));
                var localDir = new DirectoryInfo(Settings.PikSettings.LocalSettingsFolder);
                CopyFilesRecursively(serverUserGroupDir, localDir, token);
            }            
        }        

        /// <summary>
        /// Копирование общих настроек
        /// </summary>
        private static void CopyCommonFiles(CancellationToken token)
        {
            // Проверка версии общих настроек
            if (VersionsEqal(Path.Combine(Settings.PikSettings.LocalSettingsFolder, CommonSettingsName + ".ver"),
                Path.Combine(Settings.PikSettings.ServerSettingsFolder, $@"{CommonSettingsName}\{CommonSettingsName}.ver")))
                return;
            // Копирование общих настроек из папки Общие на сервере в локальную папку Settings
            var serverCommonDir = new DirectoryInfo(Path.Combine(Settings.PikSettings.ServerSettingsFolder, CommonSettingsName));
            var localDir = new DirectoryInfo(Settings.PikSettings.LocalSettingsFolder);
            CopyFilesRecursively(serverCommonDir, localDir, token);
        }

        /// <summary>
        /// Равны ли версии общих настроек локально и на сервере
        /// </summary>        
        private static bool VersionsEqal(string localVerFile, string serverVerFile)
        {
            var verServer = GetVersion(serverVerFile);
            if (verServer == null) return false;
            var verLocal = GetVersion(localVerFile);            
            return string.Equals(verLocal, verServer, StringComparison.OrdinalIgnoreCase);
        }        

        private static string GetVersion(string filePath)
        {
            try
            {
                return File.ReadLines(filePath).First();
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Копирование файлов настроек с сервера
        /// </summary>
        public static void CopyFilesRecursively(DirectoryInfo source, DirectoryInfo target, CancellationToken token)
        {
            // копрование всех папок из источника
            foreach (DirectoryInfo dir in source.GetDirectories())
            {
                token.ThrowIfCancellationRequested();
                // Если это папка с именем другого отдело, то не копировать ее
                if (IsOtherGroupFolder(dir.Name)) continue;
                CopyFilesRecursively(dir, target.CreateSubdirectory(dir.Name), token);
            }
            // копирование всех файлов из папки источника
            foreach (FileInfo f in source.GetFiles())
            {
                token.ThrowIfCancellationRequested();
                try
                {
                    f.CopyTo(Path.Combine(target.FullName, f.Name), true);
                }
                catch
                {
                    //Log.Info(ex, "CopyFilesRecursively {0}",f.FullName);
                }
            }
        }

        /// <summary>
        /// Удаление папки локальных настроек
        /// </summary>
        private static void deleteFilesRecursively(DirectoryInfo target)
        {
            if (target.Name.Equals(Commands.SystemDriveName, StringComparison.OrdinalIgnoreCase))
                return;

            var files = target.GetFiles();
            foreach (var item in files)
            {
                try
                {
                    item.Delete();
                }
                catch
                {
                    if (item.Attributes.HasFlag(FileAttributes.ReadOnly))
                    {
                        item.Attributes = FileAttributes.Normal;
                        try
                        {
                            item.Delete();
                        }
                        catch { }
                    }
                }
            }

            var dirs = target.GetDirectories();
            foreach (var item in dirs)
            {
                try
                {
                    item.Delete(true);
                }
                catch
                {
                    deleteFilesRecursively(item);
                }
            }
        }

        private static bool IsOtherGroupFolder(string name)
        {
            if (Settings.PikSettings.UserGroupsCombined.Any(g => g.Equals(name, StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }
            return Settings.PikSettings.UserGroups.Contains(name, StringComparer.OrdinalIgnoreCase);
        }
    }
}
