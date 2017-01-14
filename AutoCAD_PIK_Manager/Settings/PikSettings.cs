using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using AutoCAD_PIK_Manager.Model;
using OfficeOpenXml;
using System.Threading;
using System.Threading.Tasks;

namespace AutoCAD_PIK_Manager.Settings
{
    /// <summary>
    /// Настройки Autocad_Pik_Manager
    /// </summary>
    public static class PikSettings
    {
        readonly static List<string> pathesServerCadSettings = new List<string> {            
            @"\\dsk2.picompany.ru\project\CAD_Settings\AutoCAD_server\Адаптация",
            @"\\ab7\CAD_Settings\AutoCAD_server\Адаптация",
            @"\\ab5\CAD_Settings\AutoCAD_server\Адаптация"
        };

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
        /// <summary>
        /// Группа пользователя. Допускается указывать несколько через запятую
        /// </summary>
        public static string UserGroup { get { return _userGroup; } }
        /// <summary>
        /// Комбинация групп пользователя - например КР-МН и КР-СБ
        /// </summary>
        public static List<string> UserGroupsCombined { get; private set; }
        /// <summary>
        /// Группы пользователей определенные по папкам в папке Standart на сервере
        /// </summary>
        public static List<string> UserGroups { get { return _userGroups; } }

        public static string CurDllLocation { get { return _curDllLocation; } }

        public static SettingsGroupFile GroupFileSettings { get { return _settingsGroupFile; } }        

        public static SettingsPikFile PikFileSettings { get { return _settingsPikFile; } }        

        internal static void LoadSettings()
        {
            _curDllLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            //Путь до папки Settings на локальном компьютере.
            _localSettingsFolder = Path.GetDirectoryName(_curDllLocation);
            _settingsPikFile = getSettings<SettingsPikFile>(Path.Combine(_curDllLocation, "SettingsPIK.xml"));
            if (_settingsPikFile == null)
            {
                _settingsPikFile = SettingsPikFile.Default();
            }
            _serverSettingsFolder = GetServerSettingsPath(_settingsPikFile?.ServerSettingsPath);// TODO: Можно проверить доступность серверного пути, и если он недоступен, попробовать другой.
            _serverShareSettingsFolder = GetServerShareLibPath();
            try
            {
                // Загрузка группы юзера из файла UsersLisr2
                _userGroup = getUserGroupFromServer(GetServerUserListFile());
            }
            catch
            {
                _userGroup = getUserGroupFromLocalSide(GetServerUserListFile());
            }
            if (_userGroup == "Нет")
            {
                throw new Exceptions.NoGroupException();
            }
            UserGroupsCombined = GetUserCombinedGroups();
            _userGroups = getUserGroups();
            _settingsGroupFile = LoadSettingsGroupFiles();            
        }

        private static SettingsGroupFile LoadSettingsGroupFiles()
        {
            var sgfs = new List<SettingsGroupFile>();
            foreach (var usergroup in UserGroupsCombined)
            {
                var sgf = getSettings<SettingsGroupFile>(Path.Combine(_curDllLocation, usergroup, "SettingsGroup.xml"));
                if (sgf != null)
                {
                    sgfs.Add(sgf);
                    try
                    {
                        Log.Info($"Загружены настройки группы {usergroup} из SettingsGroup.xml");
                    }
                    catch { }
                }
            }
            return SettingsGroupFile.Merge(sgfs);
        }

        private static string GetServerShareLibPath ()
        {
            var res =  Path.GetFullPath(Path.Combine(ServerSettingsFolder, @"..\ShareSettings"));            
            return res;
        }

        private static string GetServerUserListFile()
        {            
            var res = Path.GetFullPath(Path.Combine(ServerSettingsFolder, @"..\users\userlist2.xlsx"));            
            return res;
        }

        internal static string GetServerSettingsPath(string serverSettingsPath)
        {
            string res = serverSettingsPath;
            if (!Directory.Exists(res))
            {                
                foreach (var itemServerSettPath in pathesServerCadSettings)
                {
                    if (Directory.Exists(itemServerSettPath))
                    {
                        return itemServerSettPath;                        
                    }
                }
                try
                {
                    Log.Error($"Не определен путь к сетевой папке настроек - '{serverSettingsPath}'.");
                }
                catch { }
            }
            return res;
        }

        

        private static T getSettings<T>(string file)
        {
            if (!File.Exists(file)) return default(T);
            SerializerXml ser = new SerializerXml(file);
            return ser.DeserializeXmlFile<T>();
        }

        // группа пользователя из списка сотрудников (Шифр_отдела: АР, КР-МН, КР-СБ, ВК, ОВ, и т.д.)
        private static string getUserGroupFromServer(string pathToList)
        {
            string nameGroup = "";
            // Определение группы по файлу списка пользователей на сервере
            try
            {
                var epplusDll = Path.Combine(_curDllLocation, "EPPlus.dll");
                LoadDll.LoadTry(epplusDll);
                // Копирование файла списка пользователей
                string fileTemp = Path.GetTempFileName();
                File.Copy(pathToList, fileTemp, true);
                using (var xlPackage = new ExcelPackage(new FileInfo(fileTemp)))
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
                Log.Error(ex, $"Ошибка определена рабочей группы по файлу '{pathToList}'.");
            }

            if (string.IsNullOrEmpty(nameGroup))
            {
                Log.Error($"Не определена рабочая группа по файлу UserGroup2.xlsx. {pathToList}");
                // проверка была ли группа сохранена ранее в реестре
                nameGroup = loadUserGroupFromRegistry();                
                if (string.IsNullOrEmpty(nameGroup))
                {
                    // Определение группы по текущим папкам настроек  
                    nameGroup = getCurrentGroupFromLocal();
                    if (string.IsNullOrEmpty(nameGroup))
                    {
                        Log.Error($"Не определена рабочая группа (Шифр отдела). {Environment.UserName}");
                        throw new Exception("IsNullOrEmpty(nameGroup)");
                    }
                }
            }
            else
            {
                saveUserGroupToRegistry(nameGroup);
            }            
            Log.Info($"{Environment.UserName} Группа - {nameGroup}");            
            return nameGroup;
        }

        private static string getUserGroupFromLocalSide (string pathToList)
        {
            string nameGroup = "";
            // проверка была ли группа сохранена ранее в реестре
            nameGroup = loadUserGroupFromRegistry();
            if (string.IsNullOrEmpty(nameGroup))
            {
                // Определение группы по текущим папкам настроек  
                nameGroup = getCurrentGroupFromLocal();
                if (string.IsNullOrEmpty(nameGroup))
                {
                    try
                    {
                        Log.Error($"Не определена рабочая группа (Шифр отдела). {Environment.UserName}");
                    }
                    catch { }
                    throw new Exception("IsNullOrEmpty(nameGroup)");
                }
            }
            return nameGroup;
        }

        private static string getCurrentGroupFromLocal()
        {
            try
            {
                string folderStandart = Path.Combine(LocalSettingsFolder, "Standart");
                var foldersStandart = Directory.GetDirectories(folderStandart);
                return Path.GetFileName(foldersStandart.First());
            }
            catch
            {
                return null;
            }
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
            List<string> res = new List<string>();
            try
            {
                var dirStandart = new DirectoryInfo(Path.Combine(ServerSettingsFolder, "Standart"));
                res = dirStandart.GetDirectories().Select(d => d.Name).ToList();
            }
            catch { }
            return res;
        }        

        /// <summary>
        /// Определение комбинации групп пользователя. группы могут быть перечислены через запятую - КР-МН, КР-СБ
        /// </summary>        
        private static List<string> GetUserCombinedGroups()
        {
            return UserGroup.Split(',').Select(s => s.Trim()).ToList();
        }
    }
}