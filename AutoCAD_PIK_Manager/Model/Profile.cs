using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using AutoCAD_PIK_Manager.Settings;
using Autodesk.AutoCAD.ApplicationServices;
using AutoCadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCAD_PIK_Manager.Model
{
    /// <summary>
    /// Настройка профиля ПИК в автокаде.
    /// </summary>
    internal class Profile
    {
        #region Private Fields

        private static string _localSettingsFolder;
        private static string _profileName;
        private static SettingsGroupFile _settGroupFile;
        private static SettingsPikFile _settPikFile;
        private static string _userGroup;
        private static List<string> _usersComError;

        #endregion Private Fields

        #region Public Constructors

        public Profile()
        {
            Init();
        }

        public static void Init()
        {
            _usersComError = new List<string> { "LilyuevAA", "PodnebesnovVK", "kozlovsb" }; // у BystrovDS теперь другой комп
            _settPikFile = PikSettings.PikFileSettings;
            _profileName = _settPikFile.ProfileName;
            _settGroupFile = PikSettings.GroupFileSettings;
            _userGroup = PikSettings.UserGroup;
            _localSettingsFolder = PikSettings.LocalSettingsFolder;
        }

        #endregion Public Constructors

        #region Public Methods

        /// <summary>
        /// Настройка профиля ПИК в автокаде
        /// </summary>
        public void SetProfile()
        {
            try
            {
                if (!_usersComError.Exists(u => string.Equals(u, Environment.UserName, StringComparison.OrdinalIgnoreCase)))
                {
                    dynamic preferences = AutoCadApp.Preferences;
                    object profiles = null;
                    preferences.Profiles.GetAllProfileNames(out profiles);
                    //Проверка существующих профилей
                    bool isExistProfile = ((string[])profiles).Any(x => x.Equals(_profileName));
                    if (isExistProfile)
                    {
                        if (preferences.Profiles.ActiveProfile != _profileName)
                        {
                            preferences.Profiles.ActiveProfile = _profileName;
                        }
                    }
                    else
                    {
                        preferences.Profiles.CopyProfile(preferences.Profiles.ActiveProfile, _profileName);
                        preferences.Profiles.ActiveProfile = _profileName;
                        Log.Info("Профиль {0} создан", _profileName);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка управления профилями.");
            }
            // Но настройки все равно настраиваем, даже в текущем профиле не ПИК.
            ApplySetting();
        }

        #endregion Public Methods

        #region Private Methods

        // Настройка профиля
        private void ApplySetting()
        {
            dynamic preference = AutoCadApp.Preferences;
            //IConfigurationSection con = AutoCadApp.UserConfigurationManager.OpenCurrentProfile();
            string path = string.Empty;
            Version v2015 = new Version(20, 0);
            Version cuVer = new Version(Application.Version.Major, Application.Version.Minor);

            // SupportPaths    
            SetupSupportPath();

            // PrinterConfigPaths         
            try
            {
                var varsPath = GetPaths(_settPikFile.PathVariables.PrinterConfigPaths, _settGroupFile?.PathVariables?.PrinterConfigPaths);
                path = GetPathVariable(varsPath, preference.Files.PrinterConfigPath, "");
                //Log.Warn("PrinterConfigPaths. Before - path=" + path);
                path = ExcludePikPaths(path, varsPath);
                //Log.Warn("PrinterConfigPaths. After - path=" +path);
                if (!string.IsNullOrEmpty(path))
                {
                    //if (cuVer < v2015)
                    //{
                    // Глючит печать в 2013-2014 версии.
                    // Скопировать файлы из нашей папки в первую папку из списка путей к принтерам.
                    if (_settPikFile.PathVariables.PrinterConfigPaths.Count > 0)
                    {
                        string pathPikPlotters = Path.Combine(_localSettingsFolder, _settPikFile.PathVariables.PrinterConfigPaths[0].Value);
                        string pathCurProfilePlotters = Env.GetEnv("PrinterConfigDir").Split(';').First();
                        CopyFilesToFisrtPathInCurProfile(pathPikPlotters, pathCurProfilePlotters);
                        //// Исключить наши папки из путей принтеров
                        //string pathEx = getPathWithoutOurPlotters(path, _settPikFile.PathVariables.PrinterConfigPaths);
                        //if (path != pathEx)
                        //    Env.SetEnv("PrinterConfigDir", pathEx);
                    }
                    //}
                    //else
                    //{
                    try
                    {
                        preference.Files.PrinterConfigPath = path;
                    }
                    catch
                    {
                        Env.SetEnv("PrinterConfigDir", path);
                    }

                    //}
                }
                Log.Info("PrinterConfigPath={0}", path);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "preference.Files.PrinterConfigPath = {0}", path);
            }

            // PrinterDescPaths         
            try
            {
                var varsPath = GetPaths(_settPikFile.PathVariables.PrinterDescPaths, _settGroupFile?.PathVariables?.PrinterDescPaths);
                path = GetPathVariable(varsPath, preference.Files.PrinterDescPath, "");
                path = ExcludePikPaths(path, varsPath);
                if (!string.IsNullOrEmpty(path))
                {
                    //if (cuVer < v2015)
                    //{
                    // Глючит печать в 2013-2014 версии.
                    // Скопировать файлы из нашей папки в первую папку из списка путей к принтерам.
                    if (_settPikFile.PathVariables.PrinterDescPaths.Count > 0)
                    {
                        string pathPikPrinterDesc = Path.Combine(_localSettingsFolder, _settPikFile.PathVariables.PrinterDescPaths[0].Value);
                        string pathCurProfilePrinterDesc = Env.GetEnv("PrinterDescDir").Split(';').First();
                        CopyFilesToFisrtPathInCurProfile(pathPikPrinterDesc, pathCurProfilePrinterDesc);
                        //// Исключить наши папки из путей принтеров
                        //string pathEx = getPathWithoutOurPlotters(path, _settPikFile.PathVariables.PrinterDescPaths);
                        //if (path != pathEx)
                        //    Env.SetEnv("PrinterDescDir", pathEx);
                    }
                    //}
                    //else
                    //{
                    try
                    {
                        preference.Files.PrinterDescPath = path;
                    }
                    catch
                    {
                        Env.SetEnv("PrinterDescDir", path);
                    }
                    //}
                }
                Log.Info("PrinterDescDir={0}", path);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "preference.Files.PrinterDescDir = {0}", path);
            }

            // PrinterPlotStylePaths         
            try
            {
                var varsPath = GetPaths(_settPikFile.PathVariables.PrinterPlotStylePaths, _settGroupFile?.PathVariables?.PrinterPlotStylePaths);
                path = GetPathVariable(varsPath, preference.Files.PrinterStyleSheetPath, "");
                path = ExcludePikPaths(path, varsPath);
                if (!string.IsNullOrEmpty(path))
                {
                    //if (cuVer < v2015)
                    //{
                    // Глючит печать в 2013-2014 версии.
                    // Скопировать файлы из нашей папки в первую папку из списка путей к принтерам.
                    if (_settPikFile.PathVariables.PrinterPlotStylePaths.Count > 0)
                    {
                        string pathPikPlotStyle = Path.Combine(_localSettingsFolder, _settPikFile.PathVariables.PrinterPlotStylePaths[0].Value);
                        string pathCurProfilePlotStyle = Env.GetEnv("PrinterStyleSheetDir").Split(';').First();
                        CopyFilesToFisrtPathInCurProfile(pathPikPlotStyle, pathCurProfilePlotStyle);
                        //// Исключить наши папки из путей pathCurProfilePlotStyle
                        //string pathEx = getPathWithoutOurPlotters(path, _settPikFile.PathVariables.PrinterPlotStylePaths);
                        //if (path != pathEx)
                        //    Env.SetEnv("PrinterStyleSheetDir", pathEx);
                    }
                    //}
                    //else
                    //{
                    try
                    {
                        preference.Files.PrinterStyleSheetPath = path;
                    }
                    catch
                    {
                        Env.SetEnv("PrinterStyleSheetDir", path);
                    }
                    //}
                }
                Log.Info("PrinterStyleSheetDir={0}", path);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "preference.Files.PrinterStyleSheetDir = {0}", path);
            }

            // ToolPalettePath
            try
            {
                path = GetPathVariable(GetPaths(_settPikFile.PathVariables.ToolPalettePaths, _settGroupFile?.PathVariables?.ToolPalettePaths), preference.Files.ToolPalettePath, _userGroup);
                if (!string.IsNullOrEmpty(path))
                {
                    preference.Files.ToolPalettePath = path;
                }
                Log.Info("ToolPalettePath={0}", path);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "preference.Files.ToolPalettePath = {0}", path);
            }

            //TemplatePath
            try
            {
                path = Path.Combine(_localSettingsFolder, _settPikFile.PathVariables.TemplatePath.Value, _userGroup);
                if (Directory.Exists(path))
                {
                    if (!string.IsNullOrEmpty(path))
                    {
                        Env.SetEnv("TemplatePath", path);
                    }
                    Log.Info("TemplatePath={0}", path);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Env.SetEnv(TemplatePath = {0}", path);
            }

            //PageSetupOverridesTemplateFile
            try
            {
                path = Path.Combine(_localSettingsFolder, _settPikFile.PathVariables.QNewTemplateFile.Value, _userGroup, _userGroup + ".dwt");
                if (File.Exists(path))
                {
                    Env.SetEnv("QnewTemplate", path);
                    Log.Info("QnewTemplate={0}", path);
                    preference.Files.PageSetupOverridesTemplateFile = path;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Env.SetEnv(QnewTemplate = {0}", path);
            }

            //SheetSetTemplatePath
            try
            {
                path = Path.Combine(_localSettingsFolder, _settPikFile.PathVariables.SheetSetTemplatePath.Value, _userGroup);
                if (Directory.Exists(path))
                {
                    Env.SetEnv("SheetSetTemplatePath", path);
                    Log.Info("SheetSetTemplatePath={0}", path);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Env.SetEnv(SheetSetTemplatePath = {0}", path);
            }

            // ColorBookLocation
            try
            {
                path = GetPathVariable(GetPaths(_settPikFile.PathVariables.ColorBookPaths, _settGroupFile?.PathVariables?.ColorBookPaths), preference.Files.ColorBookPath, "");
                if (!string.IsNullOrEmpty(path))
                {
                    preference.Files.ColorBookPath = path;
                }
                Log.Info("ColorBookPath={0}", path);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "preference.Files.ColorBookPath = {0}", path);
            }

            // Системные переменные
            foreach (var sysVar in _settPikFile.SystemVariables)
            {
                try
                {
                    SetSystemVariable(sysVar.Name, sysVar.Value, sysVar.IsReWrite);
                    Log.Info("Установка системной переменной {0}={1}, с перезаписью -{2}", sysVar.Name, sysVar.Value, sysVar.IsReWrite);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "Уст сис перем {0} = {1}", sysVar.Name, sysVar.Value);
                }
            }

            if (_settGroupFile?.FlexBricsSetup == true)
            {
                try
                {
                    FlexBrics.Setup();
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "FlexBrics.Setup()");
                }
            }
        }

        /// <summary>
        /// Добавление путей доступа определенных в файле settingsPik и settingsGroup
        /// </summary>
        public static void SetupSupportPath()
        {
            dynamic preference = AutoCadApp.Preferences;
            try
            {
                var path = GetPathVariable(GetPaths(_settPikFile.PathVariables.Supports, _settGroupFile?.PathVariables?.Supports), preference.Files.SupportPath, "");
                if (!string.IsNullOrEmpty(path))
                {
                    preference.Files.SupportPath = path;
                }
                Log.Info("SupportPath={0}", path);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "preference.Files.SupportPath");
            }            
        }

        private string ExcludePikPaths(string path, List<Variable> vars)
        {
            path = path.ToLower();
            foreach (var variable in vars)
            {
                var varPath = (Path.Combine(_localSettingsFolder, variable.Value) + ";").ToLower();                
                //Log.Warn($"varPath = {varPath}");
                path = path.Replace(varPath, "");
                //Log.Warn($"Remove variable {variable.Value} - path = {path}");
            }
            return path;
        }

        private string getPathWithoutOurPlotters(string path, List<Variable> vars)
        {
            Dictionary<string, string> dicPath = new Dictionary<string, string>();
            var paths = path.Split(';');
            foreach (var item in paths)
            {
                if (string.IsNullOrEmpty(item)) continue;
                dicPath.Add(item.ToUpper(), item);
            }
            foreach (var var in vars)
            {
                string varPath = Path.Combine(_localSettingsFolder, var.Value);
                if (dicPath.ContainsKey(varPath.ToUpper()))
                {
                    dicPath.Remove(varPath.ToUpper());
                }
            }
            return string.Join(";", dicPath.Values.ToArray()) + (dicPath.Count > 1 ? ";" : "");
        }

        private void CopyFilesToFisrtPathInCurProfile(string sourceFolderr, string destFolder)
        {
            if (!Directory.Exists(sourceFolderr) || !Directory.Exists(destFolder))
                return;
            DirectoryInfo dirSource = new DirectoryInfo(sourceFolderr);
            var filePlotters = dirSource.GetFiles("*.pc3", SearchOption.TopDirectoryOnly);
            foreach (var filePlotter in filePlotters)
            {
                try
                {
                    filePlotter.CopyTo(Path.Combine(destFolder, filePlotter.Name), true);
                }
                catch { }
            }
        }

        private static List<Variable> GetPaths(List<Variable> pathVars1, List<Variable> pathVars2)
        {
            var resList = new List<Variable>(pathVars1);
            var values1 = resList.Select(v => v.Value);
            if (pathVars2 != null)
            {
                foreach (var item in pathVars2)
                {
                    if (!values1.Contains(item.Value, StringComparer.OrdinalIgnoreCase))
                    {
                        resList.Add(item);
                    }
                }
            }
            return resList;
        }

        private static string GetPathVariable(List<Variable> settings, string path, string group)
        {
            string fullPath = string.Empty;
            try
            {
                bool isWrite = false;
                foreach (var setting in settings)
                {
                    string valuePath = Path.Combine(_localSettingsFolder, setting.Value, group);
                    if (Directory.Exists(valuePath))
                    {
                        isWrite = setting.IsReWrite;
                        if ((!path.ToUpper().Contains(valuePath.ToUpper())) || (isWrite))
                        {
                            fullPath += valuePath + ";";
                        }
                    }
                }
                if (!isWrite)
                {
                    fullPath = path + (path.EndsWith(";") ? "" : ";") + fullPath;
                }
            }
            catch { }

            fullPath = getOnlyExistsPaths(fullPath);

            return fullPath;
        }

        // Удаление несуществующих путей.
        private static string getOnlyExistsPaths(string fullPath)
        {
            Dictionary<string, string> existsPath = new Dictionary<string, string>();
            string pathUpper;
            var paths = fullPath.Split(';');
            if (paths.Length <= 1)
            {
                return fullPath;
            }
            foreach (var path in paths)
            {
                if (string.IsNullOrWhiteSpace(path))
                {
                    continue;
                }
                pathUpper = path.ToUpper();
                if (existsPath.ContainsKey(pathUpper))
                {
                    continue;
                }
                try
                {
                    FileAttributes attr = File.GetAttributes(path);
                    if (attr.HasFlag(FileAttributes.Directory))
                    {
                        if (Directory.Exists(path))
                        {
                            existsPath.Add(pathUpper, path);
                        }
                    }
                    else
                    {
                        if (File.Exists(path))
                        {
                            existsPath.Add(pathUpper, path);
                        }
                    }
                }
                catch
                {
                    // Это не путь
                    continue;
                }
            }
            return string.Join(";", existsPath.Values.ToArray()) + (existsPath.Count > 1 ? ";" : "");
        }

        private void SetSystemVariable(string name, string value, bool isReWrite)
        {
            object nameVar = null;
            try
            {
                nameVar = AutoCadApp.GetSystemVariable(name);
            }
            catch
            {
                return;
            }

            if (name.ToUpper() == "TRUSTEDPATHS")
            {
                var paths = value.Split(';').Select(p => Path.Combine(_localSettingsFolder, p));
                value = String.Join(";", paths);
                if (isReWrite)
                {
                    nameVar = value;
                }
                else
                {
                    nameVar = getOnlyExistsPaths(nameVar?.ToString());
                    if (string.IsNullOrEmpty(nameVar?.ToString()))
                    {
                        nameVar = value;
                    }
                    else
                    {
                        nameVar += ";" + value;
                    }
                }
                nameVar = getOnlyExistsPaths(nameVar?.ToString());
                Log.Info("TRUSTEDPATHS = {0}", nameVar);
            }
            else
            {
                if (isReWrite)
                {
                    nameVar = value;
                }
                else
                {
                    if ((nameVar == null) || (Convert.ToString(nameVar) == ""))
                    {
                        nameVar = value;
                    }
                    else if (!value.ToUpper().Contains((nameVar.ToString().Remove(nameVar.ToString().Length - 1, 1)).ToUpper()))
                    {
                        nameVar += ";" + value;
                    }
                }
            }

            try
            {
                AutoCadApp.SetSystemVariable(name, nameVar);
            }
            catch
            {
                try
                {
                    AutoCadApp.SetSystemVariable(name, Convert.ToInt16(nameVar));
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "SetSystemVariable {0}", name);
                }
            }
        }

        #endregion Private Methods

        ///// <summary>
        ///// Дефолтные настройки профиля.
        ///// Возможно лучше вообще ничего не делать, чем портить профиль человеку. Один фиг работать не будеть.!!!???
        ///// </summary>
        //private static void CreateTemplateSettings()
        //{
        //   PathVariable pathVar = new PathVariable();
        //   List<Variable> supports = new List<Variable>();
        //   supports.Add(new Variable { Name = "SupportPaths", Value = @"C:\Autodesk\AutoCAD\Pik\Settings\Fonts" });
        //   supports.Add(new Variable { Name = "SupportPaths", Value = @"C:\Autodesk\AutoCAD\Pik\Settings\Support" });
        //   pathVar.Supports = supports;
        //   List<Variable> printerConfigPaths = new List<Variable>();
        //   printerConfigPaths.Add(new Variable { Name = "PrinterConfigPaths", Value = @"C:\Autodesk\AutoCAD\Pik\Settings\Plotters", IsReWrite = false });
        //   pathVar.PrinterConfigPaths = printerConfigPaths;

        //   List<Variable> printerDescPaths = new List<Variable>();
        //   printerDescPaths.Add(new Variable { Name = "PrinterDescPaths", Value = @"C:\Autodesk\AutoCAD\Pik\Settings\Plotters\PMP files", IsReWrite = false });
        //   pathVar.PrinterDescPaths = printerDescPaths;

        //   List<Variable> printerPlotStylePaths = new List<Variable>();
        //   printerPlotStylePaths.Add(new Variable { Name = "PrinterPlotStylePaths", Value = @"C:\Autodesk\AutoCAD\Pik\Settings\Plotters\Plot Styles", IsReWrite = false });
        //   pathVar.PrinterPlotStylePaths = printerPlotStylePaths;

        //   List<Variable> toolPalettePaths = new List<Variable>();
        //   toolPalettePaths.Add(new Variable { Name = "ToolPalettePaths", Value = @"C:\Autodesk\AutoCAD\Pik\Settings\ToolPalette\", IsReWrite = false });
        //   pathVar.ToolPalettePaths = toolPalettePaths;

        //   pathVar.TemplatePath = new Variable
        //   {
        //      Name = "TemplatePath",
        //      Value = @"C:\Autodesk\AutoCAD\Pik\Settings\Template\",
        //      IsReWrite = true
        //   };

        //   pathVar.SheetSetTemplatePath = new Variable
        //   {
        //      Name = "SheetSetTemplatePath",
        //      Value = @"C:\Autodesk\AutoCAD\Pik\Settings\Sheet set\",
        //      IsReWrite = true
        //   };

        //   pathVar.PageSetupOverridesTemplateFile = new Variable
        //   {
        //      Name = "PageSetupOverridesTemplateFile",
        //      Value = @"C:\Autodesk\AutoCAD\Pik\Settings\Sheet set\",
        //      IsReWrite = true
        //   };

        //   pathVar.QNewTemplateFile = new Variable
        //   {
        //      Name = "QNewTemplateFile",
        //      Value = @"C:\Autodesk\AutoCAD\Pik\Settings\Template\",
        //      IsReWrite = true
        //   };

        //   List<SystemVariable> sysVar = new List<SystemVariable>();
        //   for (int i = 0; i < 5; i++)
        //   {
        //      SystemVariable sys = new SystemVariable();
        //      sys.Name = "Name" + i.ToString();
        //      sys.Value = "Value" + i.ToString();
        //      sys.IsReWrite = true;
        //      sysVar.Add(sys);
        //   }
        //   SettingsPikFile settings = new SettingsPikFile();
        //   settings.ProfileName = "ПИК";
        //   settings.LocalSettingsPath = @"C:\Autodesk\AutoCAD\Pik\Settings\";
        //   settings.ServerSettingsPath = @"Z:\AutoCAD_server\Адаптация\";
        //   settings.PathToUserList = @"Z:\Settings\Users\UserList2.xlsx";
        //   settings.PathVariables = pathVar;
        //   settings.SystemVariables = sysVar;
        //   List<Group> listGroup = new List<Group>();
        //   listGroup.Add(new Group() { Name = "Название отдела", Code = "Шифр отдела" });
        //   settings.Groups = listGroup;
        //   settings.NameCADManager = "Вильдару Хисяметдинову";
        //   settings.MailCADManager = "mailto:OstaninAM@pik.ru";
        //   settings.SubjectMail = "444";
        //   settings.BodyMail = "555";
        //   SerializerXml ser = new SerializerXml(@"c:\temp\!Acad_adds\Plugins\NET-VIL\AutoCAD_PIK_Manager\AutoCAD_PIK_Manager\bin\Debug\SettingsPIK.xml");
        //   ser.SerializeList(settings);
        //}
    }
}