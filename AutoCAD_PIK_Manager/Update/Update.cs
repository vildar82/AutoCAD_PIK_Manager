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
                        var filesToCopy = CopyCommonFiles(token.Token);                        
                        // Копирование настроек отдела                                                
                        var userFilesToCopy = CopyUserGroupFiles(token.Token);                        
                        if (userFilesToCopy != null)
                        {
                            if (filesToCopy == null)
                                filesToCopy = userFilesToCopy;
                            else
                                filesToCopy.AddRange(userFilesToCopy);
                        }
                        CopyFiles(filesToCopy, token.Token);
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

        private static List<Tuple<FileInfo, FileInfo>> CopyUserGroupFiles(CancellationToken token)
        {
            var copyedFiles = new List<Tuple<FileInfo, FileInfo>>();
            foreach (var group in Settings.PikSettings.UserGroupsCombined)
            {
                // Проверка версии настроек отдела (UserGroup)
                if (VersionsEqal(Path.Combine(Settings.PikSettings.LocalSettingsFolder, group + ".ver"),
                    Path.Combine(Settings.PikSettings.ServerSettingsFolder, $@"{group}\{group}.ver")))
                    return null;
                // Копирование настроек с сервера в локальную папку Settings
                var serverUserGroupDir = new DirectoryInfo(Path.Combine(Settings.PikSettings.ServerSettingsFolder, group));
                var localDir = new DirectoryInfo(Settings.PikSettings.LocalSettingsFolder);
                copyedFiles.AddRange(GetCopyedFiles(serverUserGroupDir, localDir, token));
            }
            return copyedFiles;
        }        

        /// <summary>
        /// Копирование общих настроек
        /// </summary>
        private static List<Tuple<FileInfo, FileInfo>> CopyCommonFiles(CancellationToken token)
        {
            // Проверка версии общих настроек
            if (VersionsEqal(Path.Combine(Settings.PikSettings.LocalSettingsFolder, CommonSettingsName + ".ver"),
                Path.Combine(Settings.PikSettings.ServerSettingsFolder, $@"{CommonSettingsName}\{CommonSettingsName}.ver")))
                return null;
            // Копирование общих настроек из папки Общие на сервере в локальную папку Settings
            var serverCommonDir = new DirectoryInfo(Path.Combine(Settings.PikSettings.ServerSettingsFolder, CommonSettingsName));
            var localDir = new DirectoryInfo(Settings.PikSettings.LocalSettingsFolder);
            return GetCopyedFiles(serverCommonDir, localDir, token);
        }

        /// <summary>
        /// Равны ли версии общих настроек локально и на сервере
        /// </summary>        
        public static bool VersionsEqal(string localVerFile, string serverVerFile)
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
        public static List<Tuple<FileInfo, FileInfo>> GetCopyedFiles(DirectoryInfo source, DirectoryInfo target, CancellationToken token)
        {            
            if (source == null || !source.Exists) return null;

            var filesToCopy = new List<Tuple<FileInfo, FileInfo>>();

            // копрование всех папок из источника
            foreach (DirectoryInfo dir in source.GetDirectories())
            {
                token.ThrowIfCancellationRequested();
                // Если это папка с именем другого отдело, то не копировать ее
                if (IsOtherGroupFolder(dir.Name)) continue;
                var files = GetCopyedFiles(dir, target.CreateSubdirectory(dir.Name), token);
                if (files != null)
                {
                    filesToCopy.AddRange(files);
                }
            }
            // Файлы в папке
            var sourceFiles = source.GetFiles();
            foreach (var sf in sourceFiles)
            {                
                filesToCopy.Add(new Tuple<FileInfo, FileInfo> (sf, new FileInfo (Path.Combine(target.FullName, sf.Name))));
            }
            return filesToCopy;
        }

        public static void CopyFiles (List<Tuple<FileInfo, FileInfo>> filesSourceDest, CancellationToken token)
        {
            foreach (var sd in filesSourceDest)
            {
                token.ThrowIfCancellationRequested();
                try
                {
                    sd.Item1.CopyTo(sd.Item2.FullName, true);
                }
                catch
                {
                    //Log.Info(ex, "CopyFilesRecursively {0}",f.FullName);
                }               
            }
            // Удаление лишних файлов
            var localDir = new DirectoryInfo(Settings.PikSettings.LocalSettingsFolder);
            var allLocalFiles = localDir.GetFiles("*.*", SearchOption.AllDirectories);
            DeleteExcessFiles(allLocalFiles, filesSourceDest.Select(s => s.Item2).ToArray());
        }

        private static void DeleteExcessFiles(FileInfo[] localFiles, FileInfo[] serverFiles)
        {
            var excessFiles = localFiles.Except(serverFiles, new FileNameComparer());
            foreach (var item in excessFiles)
            {
                DeleteFile(item);
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
                    deleteFilesRecursively(item);
                }
            }
        }

        public static void DeleteFile(FileInfo file)
        {
            try
            {
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
    }
}
