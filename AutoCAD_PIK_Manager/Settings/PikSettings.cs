using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Serialization;
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
      
      // статический конструктор
      static PikSettings()
      {
         _curDllLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);         
      }

      internal static void LoadSettings()
      {
         _settingsPikFile = (SettingsPikFile)getSettingsPik(Path.Combine(_curDllLocation, "SettingsPIK.xml"));
      }

      /// <summary>
      /// Путь до папки настроек на локальном компьютере: c:\Autodesk\AutoCAD\Pik\Settings
      /// </summary>
      public static string LocalSettingsFolder
      {
         get
         {
            if (_localSettingsFolder == null)
            {
               //Путь до папки Settings на локальном компьютере.
               _localSettingsFolder = Path.GetPathRoot(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
            }
            return _localSettingsFolder;
         }
      }

      /// <summary>
      /// Путь к папкее настроек на сервере. z:\AutoCAD_server\Адаптация
      /// </summary>
      public static string ServerSettingsFolder
      {
         get
         {
            if (_serverSettingsFolder == null)
            {
               _serverSettingsFolder = PikFileSettings.ServerSettingsPath;
               // TODO: Можно проверить доступность серверного пути, и если он недоступен, попробовать другой.
            }
            return _serverSettingsFolder;
         }
      }

      public static string UserGroup
      {
         get
         {
            if (_userGroup == null)
            {
               _userGroup = getUserGroup(PikFileSettings.PathToUserList);
            }
            return _userGroup;
         }
      }

      public static List<string> UserGroups
      {
         get
         {
            if (_userGroups == null)
            {
               _userGroups = getUserGroups();
            }
            return _userGroups;
         }
      }

      internal static SettingsGroupFile GroupFileSettings
      {
         get
         {
            if (_settingsGroupFile == null)
            {
               _settingsGroupFile = (SettingsGroupFile)getSettingsGroup(Path.Combine(_curDllLocation, UserGroup, "SettingsPIK.xml"));
            }
            return _settingsGroupFile;
         }
      }

      internal static SettingsPikFile PikFileSettings { get { return _settingsPikFile; } }

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
            copyFilesRecursively(serverSettDir, localSettDir);
         }
         else
         {
            Log.Error("Недоступна папка настроек на сервере {0}", ServerSettingsFolder);
         }
      }

      /// <summary>
      /// Удаление папки локальных настроек
      /// </summary>
      private static void deleteFilesRecursively(DirectoryInfo target)
      {
         try
         {
            target.Delete(true);
         }
         catch (Exception ex)
         {
            Log.Error("Ошибка при удалении файлов локальных настроек.", ex);
         }
      }

      /// <summary>
      /// Копирование файлов настроек с сервера
      /// </summary>
      private static void copyFilesRecursively(DirectoryInfo source, DirectoryInfo target)
      {
         if (!(source.Exists))
         {
            return;
         }
         if (!target.Exists)
         {
            target.Create();
         }
         // копрование всех папок из источника
         foreach (DirectoryInfo dir in source.GetDirectories())
         {
            // Если это папка с именем другого отдело, то не копировать ее
            if (isOtherGroupFolder(dir.Name)) continue;
            copyFilesRecursively(dir, target);
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

      private static SettingsPikFile getSettingsPik(string file)
      {         
         SerializerXml ser = new SerializerXml(file);
         return ser.DeserializeXmlFile<SettingsPikFile>();
      }
      private static SettingsGroupFile getSettingsGroup(string file)
      {         
         SerializerXml ser = new SerializerXml(file);
         return ser.DeserializeXmlFile<SettingsGroupFile>();
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
         }
         catch (Exception ex)
         {
            Log.Error("Не определена рабочая группа (Шифр отдела). " + Environment.UserName, ex);
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
         return _userGroups.Contains(name, StringComparer.OrdinalIgnoreCase);
      }
   }
}