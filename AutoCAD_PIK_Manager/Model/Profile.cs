﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using AutoCAD_PIK_Manager.Settings;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using AutoCadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Threading.Tasks;
using System.Threading;

namespace AutoCAD_PIK_Manager.Model
{
    /// <summary>
    /// Настройка профиля ПИК в автокаде.
    /// </summary>
    internal class Profile
    {     
        private static string _localSettingsFolder;
        private static string _profileName;
        private static SettingsGroupFile _settGroupFile;
        private static SettingsPikFile _settPikFile;
        private static string _userGroup;
        private static List<string> _usersComError;
        public bool SetProfilePIK { get; set; } = true;
        public bool SetToolPalette { get; set; } = true;
        public bool SetTemplate { get; set; } = true;
        public List<SystemVariable> SetSysVars { get; set; }

        public Profile()
        {
            Init();
            SetSysVars = _settPikFile.SystemVariables;
        }

        public static void Init()
        {
            _usersComError = new List<string> { "LilyuevAA", "tishchenkoag" };// { "LilyuevAA", "PodnebesnovVK", "kozlovsb", "tishchenkoag" };
            _settPikFile = PikSettings.PikFileSettings;
            _profileName = _settPikFile.ProfileName;
            _settGroupFile = PikSettings.GroupFileSettings;
            _userGroup = PikSettings.UserGroupsCombined.First();
            _localSettingsFolder = PikSettings.LocalSettingsFolder;            
        }        

        /// <summary>
        /// Настройка профиля ПИК в автокаде
        /// </summary>
        public void SetProfile()
        {
            try
            {
                if (SetProfilePIK && !_usersComError.Exists(u => string.Equals(u, Environment.UserName, StringComparison.OrdinalIgnoreCase)))
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
                        Log.Info($"Профиль {_profileName} создан");
                    }
                }
            }
            catch (Exception ex)
            {
                try
                {
                    Log.Error(ex, "Ошибка управления профилями.");
                }
                catch { }
            }
            // Но настройки все равно настраиваем, даже в текущем профиле не ПИК.
            ApplySetting();
        }        

        // Настройка профиля
        private void ApplySetting()
        {
            //dynamic preference = AutoCadApp.Preferences;            
            //Version v2015 = new Version(20, 0);
            //Version cuVer = new Version(Application.Version.Major, Application.Version.Minor);

            SetupSupportPath();
            PrinterConfigPaths();
            ToolPalettePath();

            //TemplatePath
            if (SetTemplate)
            {
                TemplatePath();
                PageSetupOverridesTemplateFile();
                SheetSetTemplatePath();
            }

            // ColorBookLocation
            ColorBookLocation();

            // Системные переменные
            SetSystemVariables();
            // Установка флексбрикс если задано
            FlexBrics();
        }

        private static void FlexBrics()
        {
            if (_settGroupFile?.FlexBricsSetup == true)
            {
                try
                {
                    Settings.FlexBrics.Setup();
                }
                catch (Exception ex)
                {
                    try
                    {
                        Log.Error(ex, "FlexBrics.Setup()");
                    }
                    catch { }
                }
            }
        }

        private void SetSystemVariables()
        {
            if (SetSysVars == null) return;
            foreach (var sysVar in SetSysVars)
            {
                try
                {
                    SetSystemVariable(sysVar.Name, sysVar.Value, sysVar.IsReWrite);
                    Log.Info($"Установка системной переменной {sysVar.Name}={sysVar.Value}, с перезаписью -{sysVar.IsReWrite}");
                }
                catch (Exception ex)
                {
                    try
                    {
                        Log.Error(ex, $"Уст сис перем {sysVar.Name} = {sysVar.Value}");
                    }
                    catch { }
                }
            }
        }

        private void ColorBookLocation()
        {
            try
            {
                dynamic preference = AutoCadApp.Preferences;
                string path = GetPathVariable(GetPaths(_settPikFile.PathVariables.ColorBookPaths, 
                    _settGroupFile?.PathVariables?.ColorBookPaths), preference.Files.ColorBookPath, "ColorBookPath", false);
                if (!string.IsNullOrEmpty(path))
                {
                    preference.Files.ColorBookPath = path;
                }
                Log.Info($"ColorBookPath={path}");
            }
            catch (Exception ex)
            {
                try
                {
                    Log.Error(ex, $"preference.Files.ColorBookPath");
                }
                catch { }
            }            
        }

        private void SheetSetTemplatePath()
        {
            try
            {
                dynamic preference = AutoCadApp.Preferences;
                string path = Path.Combine(_localSettingsFolder, _settPikFile.PathVariables.SheetSetTemplatePath.Value, _userGroup);
                if (Directory.Exists(path))
                {
                    Env.SetEnv("SheetSetTemplatePath", path);
                    Log.Info($"SheetSetTemplatePath={path}");
                }
            }
            catch (Exception ex)
            {
                try
                {
                    Log.Error(ex, $"Env.SetEnv(SheetSetTemplatePath)");
                }
                catch { }
            }            
        }

        private void PageSetupOverridesTemplateFile()
        {
            try
            {
                dynamic preference = AutoCadApp.Preferences;
                string path = Path.Combine(_localSettingsFolder, _settPikFile.PathVariables.QNewTemplateFile.Value, _userGroup, _userGroup + ".dwt");
                if (File.Exists(path))
                {
                    Env.SetEnv("QnewTemplate", path);
                    Log.Info($"QnewTemplate={path}");
                    preference.Files.PageSetupOverridesTemplateFile = path;
                }
            }
            catch (Exception ex)
            {
                try
                {
                    Log.Error(ex, $"Env.SetEnv(QnewTemplate)");
                }
                catch { }
            }
        }

        private void TemplatePath()
        {
            try
            {
                dynamic preference = AutoCadApp.Preferences;
                string path = Path.Combine(_localSettingsFolder, _settPikFile.PathVariables.TemplatePath.Value, _userGroup);
                if (Directory.Exists(path))
                {
                    if (!string.IsNullOrEmpty(path))
                    {
                        try
                        {
                            Env.SetEnv("TemplatePath", path);
                        }
                        catch
                        {
                            preference.Files.TemplateDwgPath = path;
                        }
                    }
                    Log.Info($"TemplatePath={path}");
                }
            }
            catch (Exception ex)
            {
                try
                {
                    Log.Error(ex, $"Env.SetEnv(TemplatePath)");
                }
                catch { }
            }            
        }

        private void ToolPalettePath()
        {
            if (!SetToolPalette) return;
            try
            {
                dynamic preference = AutoCadApp.Preferences;
                string path = GetPathVariable(GetPaths(_settPikFile.PathVariables.ToolPalettePaths, 
                    _settGroupFile?.PathVariables?.ToolPalettePaths), preference.Files.ToolPalettePath, "ToolPalettePath", true);
                if (!string.IsNullOrEmpty(path))
                {
                    preference.Files.ToolPalettePath = path;
                }
                Log.Info($"ToolPalettePath={path}");
            }
            catch (Exception ex)
            {
                try
                {
                    Log.Error(ex, $"preference.Files.ToolPalettePath");
                }
                catch { }
            }
        }

        private void PrinterConfigPaths()
        {
            try
            {
                dynamic preference = AutoCadApp.Preferences;
                var varsPath = GetPaths(_settPikFile.PathVariables.PrinterConfigPaths, _settGroupFile?.PathVariables?.PrinterConfigPaths);
                //path = GetPathVariable(varsPath, preference.Files.PrinterConfigPath, "");
                var curPlottersPaths = preference.Files.PrinterConfigPath;
                //Log.Warn("PrinterConfigPaths. Before - path=" + path);
                string path = ExcludePikPaths(curPlottersPaths, varsPath);
                //Log.Warn("PrinterConfigPaths. After - path=" +path);
                if (string.IsNullOrEmpty(path))
                {
                    // Добавить стандартный путь к плоттерам  
                    try
                    {
                        Log.Error($"Не определен стандартный путь к плоттерам!!.");
                    }
                    catch { }
                }
                else
                {

                    //if (cuVer < v2015)
                    //{
                    // Глючит печать в 2013-2014 версии.
                    // Скопировать файлы из нашей папки в первую папку из списка путей к принтерам.
                    if (_settPikFile.PathVariables.PrinterConfigPaths.Count > 0)
                    {
                        string pathPikPlotters = Path.Combine(_localSettingsFolder, _settPikFile.PathVariables.PrinterConfigPaths[0].Value);                        
                        string pathCurProfilePlotters = path.Split(';').First();                        
                        //try
                        //{
                        //    pathCurProfilePlotters = Env.GetEnv("PrinterConfigDir").Split(';').First();
                        //}
                        //catch
                        //{
                        //    pathCurProfilePlotters = path.Split(';').First();
                        //    if (string.IsNullOrEmpty(pathCurProfilePlotters))
                        //    {
                        //        // Стандартный путь к плоттерам
                        //    }
                        //}
                        CopyFilesToFisrtPathInCurProfile(pathPikPlotters, pathCurProfilePlotters);
                        try
                        {
                            Log.Info($"Скопированы плоттеры из папки {pathPikPlotters}, в папку {pathCurProfilePlotters}.");
                        }
                        catch { }
                        //// Исключить наши папки из путей принтеров
                        //string pathEx = getPathWithoutOurPlotters(path, _settPikFile.PathVariables.PrinterConfigPaths);
                        //if (path != pathEx)
                        //    Env.SetEnv("PrinterConfigDir", pathEx);
                    }

                    if (path != curPlottersPaths)
                    {
                        try
                        {
                            preference.Files.PrinterConfigPath = path;
                        }
                        catch
                        {
                            Env.SetEnv("PrinterConfigDir", path);
                        }
                    }
                }
                try
                {
                    Log.Info($"PrinterConfigPath={path}");
                }
                catch { }
            }
            catch (Exception ex)
            {
                try
                {
                    Log.Error(ex, $"preference.Files.PrinterConfigPath.");
                }
                catch { }
            }            
        }

        /// <summary>
        /// Добавление путей доступа определенных в файле settingsPik и settingsGroup
        /// </summary>
        public static void SetupSupportPath()
        {            
            try
            {
                dynamic preference = AutoCadApp.Preferences;
                string path = GetPathVariable(GetPaths(_settPikFile.PathVariables.Supports, 
                    _settGroupFile?.PathVariables?.Supports), preference.Files.SupportPath, "SupportPath", false);

                //// Копирование файов из папки Support в папку appdata/roamable Support пользователя
                //var supportPikPath = Path.Combine(_localSettingsFolder, _settPikFile.PathVariables.Supports.First(s=>s.Value == "Support").Value);
                //var roamPath = Path.Combine(HostApplicationServices.Current.RoamableRootFolder, "Support");
                //PikSettings.CopyFilesRecursively(new DirectoryInfo(supportPikPath),new DirectoryInfo ( roamPath));                

                if (!string.IsNullOrEmpty(path))
                {
                    preference.Files.SupportPath = path;
                }
                Log.Info($"SupportPath={path}");
            }
            catch (Exception ex)
            {
                try
                {
                    Log.Error(ex, "preference.Files.SupportPath");
                }
                catch { }
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
            DirectoryInfo dirDest = new DirectoryInfo(destFolder);

            try
            {
                var token = new CancellationTokenSource();
                var task = Task.Run(() => {

                    var files = Update.GetCopyedFiles(dirSource, dirDest, token.Token, true);
                    Update.CopyFiles(files, token.Token, false);
                });
                task.Wait(new TimeSpan(0,0,3));
                if (!task.IsCompleted)
                {
                    token.Cancel(true);
                }
            }
            catch { }
            
            //var filePlotters = dirSource.GetFiles("*.pc3", SearchOption.TopDirectoryOnly);
            //foreach (var filePlotter in filePlotters)
            //{
            //    try
            //    {
            //        filePlotter.CopyTo(Path.Combine(destFolder, filePlotter.Name), true);
            //    }
            //    catch { }
            //}
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

        private static string GetPathVariable(List<Variable> settings, string path, string namePathsForLog, bool withGroup)
        {
            string fullPath = string.Empty;
            try
            {
                bool isWrite = false;
                foreach (var setting in settings)
                {                    
                    if (withGroup)
                    {
                        if (setting.IsReWrite)
                        {
                            path = "";
                        }
                        foreach (var group in PikSettings.UserGroupsCombined)
                        {
                            var valuePath = Path.Combine(_localSettingsFolder, setting.Value, group);
                            //if (Directory.Exists(valuePath))
                            //{                                
                                if (!path.ToUpper().Contains(valuePath.ToUpper()))
                                {
                                    fullPath += valuePath + ";";
                                }
                            //}
                        }                        
                    }
                    else
                    {
                        var valuePath = Path.Combine(_localSettingsFolder, setting.Value);
                        //if (Directory.Exists(valuePath))
                        //{
                            isWrite = setting.IsReWrite;
                            if ((!path.ToUpper().Contains(valuePath.ToUpper())) || (isWrite))
                            {
                                fullPath += valuePath + ";";
                            }
                        //}
                    }                    
                }
                if (!isWrite)
                {
                    fullPath = path + (path.EndsWith(";") ? "" : ";") + fullPath;
                }
            }
            catch { }

            fullPath = getOnlyExistsPaths(fullPath, namePathsForLog);

            return fullPath;
        }

        // Удаление несуществующих путей.
        private static string getOnlyExistsPaths(string fullPath, string namePathsForLog)
        {
            List<string> deletedPath = new List<string>  ();
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
                    bool isAdded = false;
                    FileAttributes attr = File.GetAttributes(path);
                    if (attr.HasFlag(FileAttributes.Directory))
                    {
                        //if (Directory.Exists(path))
                        //{
                            existsPath.Add(pathUpper, path);
                            isAdded = true;
                        //}
                    }
                    else
                    {
                        //if (File.Exists(path))
                        //{
                            existsPath.Add(pathUpper, path);
                            isAdded = true;
                        //}
                    }
                    if (!isAdded)
                    {
                        deletedPath.Add(path);
                    }
                }
                catch
                {
                    // Это не путь
                    continue;
                }
            }
            if (deletedPath.Count != 0)
            {
                try
                {
                    Log.Error($"Удаленные пути из {namePathsForLog}: {string.Join(";", deletedPath)}");
                }
                catch { }
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
                    nameVar = getOnlyExistsPaths(nameVar?.ToString(), "TRUSTEDPATHS");
                    if (string.IsNullOrEmpty(nameVar?.ToString()))
                    {
                        nameVar = value;
                    }
                    else
                    {
                        nameVar += ";" + value;
                    }
                }
                nameVar = getOnlyExistsPaths(nameVar?.ToString(), "TRUSTEDPATHS");
                try
                {
                    Log.Info("TRUSTEDPATHS = {0}", nameVar);
                }
                catch { }
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
                    try
                    {
                        Log.Error(ex, $"SetSystemVariable {name}");
                    }
                    catch { }
                }
            }
        }       

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