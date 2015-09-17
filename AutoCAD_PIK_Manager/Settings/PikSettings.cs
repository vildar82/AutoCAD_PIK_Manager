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
      private static SettingsGroupFile _settingsGroupFile;
      private static SettingsPikFile _settingsPikFile;
      private static string _userGroup;
      private static List<string> _userGroups;

      /// <summary>
      /// Путь до папки настроек на локальном компьютере: c:\Autodesk\AutoCAD\Pik\Settings
      /// </summary>
      public static string LocalSettingsFolder { get { return _localSettingsFolder; } }

      /// <summary>
      /// Путь к папкее настроек на сервере. z:\AutoCAD_server\Адаптация
      /// </summary>
      public static string ServerSettingsFolder { get { return _serverSettingsFolder; } }

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
         _serverSettingsFolder = _settingsPikFile.ServerSettingsPath;// TODO: Можно проверить доступность серверного пути, и если он недоступен, попробовать другой.
         _userGroup = getUserGroup(_settingsPikFile.PathToUserList);
         _userGroups = getUserGroups();
         _settingsGroupFile = getSettings<SettingsGroupFile>(Path.Combine(_curDllLocation, UserGroup, "SettingsGroup.xml"));
         if (_settingsGroupFile != null) Log.Info("Загружены настройки группы {0} из {1}", UserGroup, "SettingsGroup.xml");
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
            catch (Exception ex)
            {
               Log.Error("CopyFilesRecursively " + f.FullName, ex);
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
            try
            {
               target.Delete(true);
            }
            catch
            {
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
            Log.Info("{0} Группа - {1}", Environment.UserName,nameGroup);
         }
         catch (Exception ex)
         {
            Log.Error("Не определена рабочая группа (Шифр отдела). " + Environment.UserName, ex);
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