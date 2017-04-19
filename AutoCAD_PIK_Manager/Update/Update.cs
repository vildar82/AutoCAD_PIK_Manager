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
        internal static string CommonSettingsName = "Общие";
        private static string updateInfo = string.Empty;
        private static string verCommonLocal;
        private static string verCommonServer;
        private static string verUserGroupLocal;
        private static string verUserGroupServer;
        private static string verFBLocal;
        private static string verFBServer;
        private static bool isTester;

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
                        // Проверка - если группа пользователя тестовая (слово тест в имени группы)
                        isTester = DefineUserTester();
                        if (isTester)
                        {
                            CommonSettingsName = $@"{CommonSettingsName}_Тест";
                        }

                        // Копирование общих настроек                        
                        var filesToCopy = GetCopiedCommonFiles(token.Token);                        
                        
                        // Копирование настроек отдела                                                
                        var userFilesToCopy = GetCopiedUserGroupFiles(token.Token);
                        if (userFilesToCopy != null)
                        {
                            if (filesToCopy == null)
                                filesToCopy = userFilesToCopy;
                            else
                                filesToCopy.AddRange(userFilesToCopy);
                        }

                        // Копирование flexBrics если нужно
                        var fbFilesToCopy = GetFBCopiedFiles(token.Token);
                        if (fbFilesToCopy != null)
                        {
                            if (filesToCopy == null)
                                filesToCopy = fbFilesToCopy;
                            else
                                filesToCopy.AddRange(fbFilesToCopy);
                        }

                        CopyFiles(filesToCopy, token.Token, true);
                        verCommonLocal = verCommonServer;
                        verUserGroupLocal = verUserGroupServer;
                        verFBLocal = verFBServer;
                        updateInfo += $" Локальные версии: {CommonSettingsName} '{verCommonLocal}', Отдела '{verUserGroupLocal}'";
                        if (Settings.FlexBrics.HasFlexBrics())
                        {
                            updateInfo += $", {Settings.FlexBrics.FbName} '{verFBLocal}'";
                        }

                        // Удаление пустых папок
                        DeleteEmptyDirectory(Settings.PikSettings.LocalSettingsFolder);
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
                task.Wait(new TimeSpan(0, 1, 10));
                if (!task.IsCompleted)
                {
                    token.Cancel(true);
                }               
            }
            catch { }            
        }

        private static bool DefineUserTester()
        {
            if (Settings.PikSettings.UserGroup.IndexOf("тест", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }
            return false;
        }

        public static string GetUpdateInfo()
        {
            return updateInfo + " Конец.";
        }

        private static List<UpdateFile> GetFBCopiedFiles(CancellationToken token)
        {
            if (!Settings.FlexBrics.HasFlexBrics()) return null;

            var copyedFiles = new List<UpdateFile>();

            var serverFbDir = Settings.FlexBrics.GetServerFlexBricsServerFolder();
            var serverFB = new DirectoryInfo(serverFbDir);
            var localFbDir = Settings.FlexBrics.GetFBLocalDir();
            var localFB = new DirectoryInfo(localFbDir);

            // Проверка версии общих настроек
            var updateRequired = !VersionsEqal(Path.Combine(localFbDir, "fb.ver"),
                                    Path.Combine(serverFbDir, "fb.ver"), 
                                    out verFBLocal, out verFBServer);
            if (!updateRequired)
            {
                updateInfo += $" Версия настроек {Settings.FlexBrics.FbName} совпадает с сервером = '{verFBServer}'.";                
            }
            // Файлы FlexBrics
            return GetCopyedFiles(serverFB, localFB, token, updateRequired);
        }

        private static List<UpdateFile> GetCopiedUserGroupFiles(CancellationToken token)
        {
            var copyedFiles = new List<UpdateFile>();
            foreach (var group in Settings.PikSettings.UserGroupsCombined)
            {
                // Проверка версии настроек отдела (UserGroup)
                var updateRequired = !VersionsEqal(Path.Combine(Settings.PikSettings.LocalSettingsFolder, group + ".ver"),
                    Path.Combine(Settings.PikSettings.ServerSettingsFolder, $@"{group}\{group}.ver"),
                    out verUserGroupLocal, out verUserGroupServer);
                if (!updateRequired)
                {
                    updateInfo += $" Версия настроек отдела совпадает с сервером = '{verUserGroupServer}'.";                    
                }
                // Копирование настроек с сервера в локальную папку Settings
                var serverUserGroupDir = new DirectoryInfo(Path.Combine(Settings.PikSettings.ServerSettingsFolder, group));
                var localDir = new DirectoryInfo(Settings.PikSettings.LocalSettingsFolder);
                copyedFiles.AddRange(GetCopyedFiles(serverUserGroupDir, localDir, token, updateRequired));
            }            
            return copyedFiles;
        }        

        /// <summary>
        /// Копирование общих настроек
        /// </summary>
        private static List<UpdateFile> GetCopiedCommonFiles(CancellationToken token)
        {
            // Проверка версии общих настроек            
            var updateRequired = !VersionsEqal(Path.Combine(Settings.PikSettings.LocalSettingsFolder, CommonSettingsName + ".ver"),
                Path.Combine(Settings.PikSettings.ServerSettingsFolder, $@"{CommonSettingsName}\{CommonSettingsName}.ver"),
                out verCommonLocal, out verCommonServer);
            if (!updateRequired)
            {
                updateInfo +=$" Версия общих настроек совпадает с сервером = '{verCommonServer}'.";                
            }
            if (isTester)
            {
                // Тестовые общие настройки - нужно всегда копировать, т.к. работает политика - которая переписывает локальные файлы из общей dll
                updateRequired = true;
            }
            // Копирование общих настроек из папки Общие на сервере в локальную папку Settings
            var serverCommonDir = new DirectoryInfo(Path.Combine(Settings.PikSettings.ServerSettingsFolder, CommonSettingsName));
            var localDir = new DirectoryInfo(Settings.PikSettings.LocalSettingsFolder);
            return GetCopyedFiles(serverCommonDir, localDir, token, updateRequired);                        
        }

        /// <summary>
        /// Равны ли версии общих настроек локально и на сервере
        /// </summary>        
        public static bool VersionsEqal(string localVerFile, string serverVerFile, out string verLocal, out string verServer)
        {
            verLocal = GetVersion(localVerFile);
            verServer = GetVersion(serverVerFile);
            if (verServer == null)
            {                
                return false;
            }            
            return string.Equals(verLocal, verServer, StringComparison.OrdinalIgnoreCase);
        }        

        private static string GetVersion(string filePath)
        {
            try
            {
                return File.ReadLines(filePath)?.First()?.Trim();
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Копирование файлов настроек с сервера
        /// </summary>
        public static List<UpdateFile> GetCopyedFiles(DirectoryInfo source, DirectoryInfo target, CancellationToken token,
            bool updateRequired)
        {            
            if (source == null || !source.Exists) return null;

            var filesToCopy = new List<UpdateFile>();

            // копрование всех папок из источника
            foreach (DirectoryInfo dir in source.GetDirectories())
            {
                token.ThrowIfCancellationRequested();
                // Если это папка с именем другого отдело, то не копировать ее
                if (IsOtherGroupFolder(dir.Name)) continue;
                var files = GetCopyedFiles(dir, target.CreateSubdirectory(dir.Name), token, updateRequired);
                if (files != null)
                {
                    filesToCopy.AddRange(files);
                }
            }
            // Файлы в папке
            var sourceFiles = source.GetFiles();
            foreach (var sf in sourceFiles)
            {                
                filesToCopy.Add(new UpdateFile(sf, new FileInfo (Path.Combine(target.FullName, sf.Name)), updateRequired));
            }
            return filesToCopy;
        }

        public static void CopyFiles (List<UpdateFile> filesUpdate, CancellationToken token, bool deleteExcessFiles)
        {
            int copiedFiles = 0;            
            foreach (var fileUpdate in filesUpdate.Where(w => w.UpdateRequired))
            {
                token.ThrowIfCancellationRequested();
                try
                {
                    fileUpdate.ServerFile.CopyTo(fileUpdate.LocalFile.FullName, true);
                    copiedFiles++;
                }
                catch
                {
                    //Log.Info(ex, "CopyFilesRecursively {0}",f.FullName);
                }
            }
            updateInfo += $" Файлов скопировано {copiedFiles}.";
            // Удаление лишних файлов
            if (deleteExcessFiles && copiedFiles>0)
            {
                var localDir = new DirectoryInfo(Settings.PikSettings.LocalSettingsFolder);
                var allLocalFiles = localDir.GetFiles("*.*", SearchOption.AllDirectories);

                DeleteExcessFiles(allLocalFiles, filesUpdate.Select(s => s.LocalFile).ToArray());
            }
        }

        private static void DeleteExcessFiles(FileInfo[] allLocalFiles, FileInfo[] updateFiles)
        {
            if (allLocalFiles == null || updateFiles == null || allLocalFiles.Length < updateFiles.Length) return;
            var excessFiles = allLocalFiles.Except(updateFiles, new FileNameComparer());
            foreach (var item in excessFiles)
            {                
                DeleteFile(item);
            }
            updateInfo += $" Удалены: " + string.Join(",", excessFiles.Select(s => s.Name)) + ".";

            // Удаление пустых папок
            //DeleteEmptyFolders(Settings.PikSettings.LocalSettingsFolder);
        }

        private static void DeleteEmptyFolders(string localSettingsFolder)
        {
            if (!localSettingsFolder.EndsWith("Settings")) return;
            var folders = Directory.GetDirectories(localSettingsFolder);            
            foreach (var item in folders)
            {
                if (!Directory.EnumerateFiles(item, "*.*", SearchOption.AllDirectories).Any())
                {
                    Directory.Delete(item, true);                    
                }
            }
        }

        /// <summary>
        /// Удаление папки локальных настроек
        /// </summary>
        private static void DeleteFilesRecursively(DirectoryInfo target)
        {
            if (target.Name.Equals(Commands.SystemDriveName, StringComparison.OrdinalIgnoreCase))
                return;

            var files = target.GetFiles();
            foreach (var item in files)
            {
                DeleteFile(item);
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
                    DeleteFilesRecursively(item);
                }
            }
        }

        public static void DeleteFile(FileInfo file)
        {
            try
            {
                if (file.Name == "SettingsGroup.xml")
                    return;
                file.Delete();
            }
            catch
            {
                if (file.Attributes.HasFlag(FileAttributes.ReadOnly))
                {
                    file.Attributes = FileAttributes.Normal;
                    try
                    {
                        file.Delete();
                    }
                    catch { }
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

        private static void DeleteEmptyDirectory(string startLocation)
        {
            foreach (var directory in Directory.GetDirectories(startLocation))
            {
                DeleteEmptyDirectory(directory);
                if (!Directory.EnumerateFiles(directory).Any() &&
                    !Directory.EnumerateDirectories(directory).Any())
                {
                    Directory.Delete(directory, false);
                }
            }
        }
    }
}
