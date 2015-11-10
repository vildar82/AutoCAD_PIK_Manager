using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;

namespace AutoCAD_PIK_Manager.Settings
{
   /// <summary>
   /// Настройки Autocad_Pik_Manager
   /// </summary>
   public static class PikSettings
   {
      private static string _curDllLocation;
      private static string _localSettingsFolder;
      private static string _serverSettingsFolder;
      private static string _serverShareSettingsFolder;
      private static SettingsGroupFile _settingsGroupFile;
      private static SettingsPikFile _settingsPikFile;
      private static string _userGroup;
      private static List<string> _userGroups;
      public const string RegAppPath = @"Software\Vildar\AutoCAD_PIK_Manager";

      /// <summary>
      /// Путь до папки настроек на локальном компьютере: c:\Autodesk\AutoCAD\Pik\Settings
      /// </summary>
      public static string LocalSettingsFolder { get { return _localSettingsFolder; } }

      /// <summary>
      /// Путь к папкее настроек на сервере. z:\AutoCAD_server\Адаптация
      /// </summary>
      public static string ServerSettingsFolder { get { return _serverSettingsFolder; } }

      /// <summary>
      /// Путь к папке с настройками программ общими для всех пользователей (share).
      /// </summary>
      public static string ServerShareSettingsFolder { get { return _serverShareSettingsFolder; } }

      public static string UserGroup { get { return _userGroup; } }

      public static List<string> UserGroups { get { return _userGroups; } }

      internal static string CurDllLocation { get { return _curDllLocation; } }

      internal static SettingsGroupFile GroupFileSettings { get { return _settingsGroupFile; } }

      internal static SettingsPikFile PikFileSettings { get { return _settingsPikFile; } }

      internal static void LoadSettings()
      {
         _curDllLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
         //Путь до папки Settings на локальном компьютере.
         _localSettingsFolder = Path.GetDirectoryName(_curDllLocation);
         _settingsPikFile = getSettings<SettingsPikFile>(Path.Combine(_curDllLocation, "SettingsPIK.xml"));
         if (_settingsPikFile == null) return;
         _serverSettingsFolder = GetExistServersettingsPath (_settingsPikFile.ServerSettingsPath);// TODO: Можно проверить доступность серверного пути, и если он недоступен, попробовать другой.
         _serverShareSettingsFolder = GetExistServersettingsPath(_settingsPikFile.ServerShareSettings);
         _userGroup = getUserGroup(GetExistServerUserListFile(_settingsPikFile.PathToUserList));
         _userGroups = getUserGroups();
         _settingsGroupFile = getSettings<SettingsGroupFile>(Path.Combine(_curDllLocation, UserGroup, "SettingsGroup.xml"));
         if (_settingsGroupFile != null) Log.Info("Загружены настройки группы {0} из {1}", UserGroup, "SettingsGroup.xml");
      }

      private static string GetExistServerUserListFile(string pathToUserList)
      {
         string res = pathToUserList;
         if (!File.Exists(res))
         {
            res = Path.Combine(@"\\ab4\CAD_Settings", pathToUserList.Substring(3));
            if (!File.Exists(res))
            {
               res = Path.Combine(@"\\dsk2.picompany.ru\project\CAD_Settings", pathToUserList.Substring(3));
               if (!File.Exists(res))
               {
                  Log.Error("Сетевой путь к файлу списка пользователей UserList2.xlsx недоступен - pathToUserList: {0}", pathToUserList);
               }
            }
         }
         return res;
      }

      internal static string GetExistServersettingsPath(string serverSettingsPath)
      {
         string res = serverSettingsPath;
         if (!Directory.Exists(res))
         {
            res = Path.Combine(@"\\dsk2.picompany.ru\project\CAD_Settings", serverSettingsPath.Substring(3));            
            if (!Directory.Exists(res))
            {
               res = Path.Combine(@"\\ab4\CAD_Settings", serverSettingsPath.Substring(3));
               if (!Directory.Exists(res))
               {
                  Log.Error("Сетевой путь к настройкам недоступен - serverSettingsPath: {0}", serverSettingsPath);
               }
            }
         }
         return res;
      }

      internal static void UpdateSettings()
      {
         // Проверка доступности сетевых настроек
         if (Directory.Exists(ServerSettingsFolder))
         {
            // Удаление локальной папки настроек
            var localSettDir = new DirectoryInfo(LocalSettingsFolder);
            deleteFilesRecursively(localSettDir);
            // Копирование настроек с сервера
            var serverSettDir = new DirectoryInfo(ServerSettingsFolder);
            localSettDir.Create();
            copyFilesRecursively(serverSettDir, localSettDir);
            // Копирование flexBrics если нужно
            FlexBrics.Copy();
         }
         else
         {
            Log.Error("Недоступна папка настроек на сервере {0}", ServerSettingsFolder);
         }
      }     

      /// <summary>
      /// Копирование файлов настроек с сервера
      /// </summary>
      private static void copyFilesRecursively(DirectoryInfo source, DirectoryInfo target)
      {
         // копрование всех папок из источника
         foreach (DirectoryInfo dir in source.GetDirectories())
         {
            // Если это папка с именем другого отдело, то не копировать ее
            if (isOtherGroupFolder(dir.Name)) continue;
            copyFilesRecursively(dir, target.CreateSubdirectory(dir.Name));
         }
         // копирование всех файлов из папки источника
         foreach (FileInfo f in source.GetFiles())
         {
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
         if (target.Name == "Settings" && target.Exists)
         {
            var dirs = target.GetDirectories();
            foreach (var item in dirs)
            {
               try
               {
                  item.Delete(true);
               }
               catch { }
            }
            var files = target.GetFiles();
            foreach (var item in files)
            {
               try
               {
                  item.Delete();
               }
               catch { }
            }
         }
      }
      private static T getSettings<T>(string file)
      {
         if (!File.Exists(file)) return default(T);
         SerializerXml ser = new SerializerXml(file);
         return ser.DeserializeXmlFile<T>();
      }

      // группа пользователя из списка сотрудников (Шифр_отдела: АР, КР-МН, КР-СБ, ВК, ОВ, и т.д.)
      private static string getUserGroup(string pathToList)
      {
         string nameGroup = "";
         try
         {
            using (var xlPackage = new ExcelPackage(new FileInfo(pathToList)))
            {
               var worksheet = xlPackage.Workbook.Worksheets[1];

               int numberRow = 2;
               while (worksheet.Cells[numberRow, 2].Text.Trim() != "")
               {
                  if (worksheet.Cells[numberRow, 2].Text.Trim().ToUpper() == Environment.UserName.ToUpper())
                  {
                     nameGroup = worksheet.Cells[numberRow, 3].Text;
                     break;
                  }
                  numberRow++;
               }
            }
            if (string.IsNullOrEmpty(nameGroup))
            {
               Log.Error("Не определена группа по файлу UserGroup2.xlsx. {0}", pathToList);
               // проверка была ли группа сохранена ранее в реестре
               nameGroup = loadUserGroupFromRegistry();
            }
            else
            {
               saveUserGroupToRegistry(nameGroup);
            }
            if (string.IsNullOrEmpty(nameGroup))
            {
               throw new Exception("IsNullOrEmpty(nameGroup)");
            }
            Log.Info("{0} Группа - {1}", Environment.UserName,nameGroup);
         }
         catch (Exception ex)
         {
            Log.Error(ex,"Не определена рабочая группа (Шифр отдела). {0}", Environment.UserName);
            throw;
         }
         //if (nameGroup == "")
         //{
         //   if (MessageBox.Show(
         //       "Ваша рабочая группа не определена!\nОбратиться за помощью к " + _settingsPikFile.NameCADManager + "?",
         //       "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
         //   {
         //      System.Diagnostics.Process.Start("mailto:" + _settingsPikFile.MailCADManager + "?subject=" + _settingsPikFile.SubjectMail + "&body=" + _settingsPikFile.BodyMail + Environment.UserName);
         //   }
         //}
         return nameGroup;
      }

      private static string loadUserGroupFromRegistry()
      {
         string res = ""; // default
         try
         {
            var keyReg = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegAppPath);
            res = (string)keyReg.GetValue("UserGroup", res);
         }
         catch { }
         return res;
      }
      private static void saveUserGroupToRegistry(string userGroup)
      {
         try
         {
            var keyReg = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegAppPath);
            keyReg.SetValue("UserGroup", userGroup, Microsoft.Win32.RegistryValueKind.String);
         }
         catch { }
      }

      private static List<string> getUserGroups()
      {
         var dirStandart = new DirectoryInfo(Path.Combine(ServerSettingsFolder, "Standart"));
         return dirStandart.GetDirectories().Select(d => d.Name).ToList();
      }

      private static bool isOtherGroupFolder(string name)
      {
         if (UserGroup.Equals(name, StringComparison.OrdinalIgnoreCase))
         {
            return false;
         }
         return _userGroups.Contains(name, StringComparer.OrdinalIgnoreCase);
      }
   }
}