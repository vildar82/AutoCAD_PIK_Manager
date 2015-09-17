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
      private string _localSettingsFolder;
      private string _profileName;
      private SettingsGroupFile _settGroupFile;
      private SettingsPikFile _settPikFile;
      private string _userGroup;
      public Profile()
      {
         _settPikFile = PikSettings.PikFileSettings;
         _profileName = _settPikFile.ProfileName;
         _settGroupFile = PikSettings.GroupFileSettings;
         _userGroup = PikSettings.UserGroup;
         _localSettingsFolder = PikSettings.LocalSettingsFolder;
      }

      static public void AddPath(string var, string path)
      {
         StringBuilder oldpath = new StringBuilder(Env.GetEnv(var));
         oldpath.Append(";"); oldpath.Append(path);
         Env.SetEnv(var, oldpath.ToString());
      }

      /// <summary>
      /// Настройка профиля ПИК в автокаде
      /// </summary>
      public void SetProfile()
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
            ApplySetting();
            Log.Info("Профиль {0} обновлен", _profileName);
         }
         else
         {
            preferences.Profiles.CopyProfile(preferences.Profiles.ActiveProfile, _profileName);
            preferences.Profiles.ActiveProfile = _profileName;
            ApplySetting();
            Log.Info("Профиль {0} создан", _profileName);
         }
      }

      // Настройка профиля
      private void ApplySetting()
      {
         dynamic preference = AutoCadApp.Preferences;
         IConfigurationSection con = AutoCadApp.UserConfigurationManager.OpenCurrentProfile();
         string path = string.Empty;

         // SupportPaths
         path = GetPathVariable(GetPaths(_settPikFile.PathVariables.Supports, _settGroupFile?.PathVariables?.Supports), preference.Files.SupportPath, "");
         try
         {
            preference.Files.SupportPath = path;
         }
         catch { }

         // PrinterConfigPaths
         //path = GetPathVariable(GetPaths(_settPikFile.PathVariables.PrinterConfigPaths, _settGroupFile == null ? null : _settGroupFile.PathVariables.PrinterConfigPaths), preference.Files.PrinterConfigPath, "");
         path = GetPathVariable(GetPaths(_settPikFile.PathVariables.PrinterConfigPaths, _settGroupFile?.PathVariables?.PrinterConfigPaths), Env.Ver == 20 ? preference.Files.PrinterConfigPath : Env.GetEnv("PrinterConfigDir"), "");
         try
         {
            if (Env.Ver == 20)
               preference.Files.PrinterConfigPath = path;
            else
               Env.SetEnv("PrinterConfigDir", path);
         }
         catch { con.OpenSubsection("General").WriteProperty("PrinterConfigDir", path); }

         // PrinterDescPaths
         //path = GetPathVariable(GetPaths(_settPikFile.PathVariables.PrinterDescPaths, _settGroupFile == null ? null : _settGroupFile.PathVariables.PrinterDescPaths), preference.Files.PrinterDescPath, "");
         path = GetPathVariable(GetPaths(_settPikFile.PathVariables.PrinterDescPaths, _settGroupFile?.PathVariables?.PrinterDescPaths), Env.Ver == 20 ? preference.Files.PrinterDescPath : Env.GetEnv("PrinterDescDir"), "");
         try
         {
            if (Env.Ver == 20)
               preference.Files.PrinterDescPath = path;
            else
               Env.SetEnv("PrinterDescDir", path);
         }
         catch { con.OpenSubsection("General").WriteProperty("PrinterDescDir", path); }

         // PrinterPlotStylePaths
         //path = GetPathVariable(GetPaths(_settPikFile.PathVariables.PrinterPlotStylePaths, _settGroupFile == null ? null : _settGroupFile.PathVariables.PrinterPlotStylePaths), preference.Files.PrinterStyleSheetPath, "");
         path = GetPathVariable(GetPaths(_settPikFile.PathVariables.PrinterPlotStylePaths, _settGroupFile?.PathVariables?.PrinterPlotStylePaths), Env.Ver == 20 ? preference.Files.PrinterStyleSheetPath : Env.GetEnv("PrinterStyleSheetDir"), "");
         try
         {
            if (Env.Ver == 20)
               preference.Files.PrinterStyleSheetPath = path;
            else
               Env.SetEnv("PrinterStyleSheetDir", path);
         }
         catch { con.OpenSubsection("General").WriteProperty("PrinterStyleSheetDir", path); }

         // ToolPalettePath
         path = GetPathVariable(GetPaths(_settPikFile.PathVariables.ToolPalettePaths, _settGroupFile?.PathVariables?.ToolPalettePaths), preference.Files.ToolPalettePath, _userGroup);
         try
         {
            preference.Files.ToolPalettePath = path;
         }
         catch { con.OpenSubsection("General").WriteProperty("ToolPalettePath", path); }

         //TemplatePath
         path = Path.Combine(_localSettingsFolder, _settPikFile.PathVariables.TemplatePath.Value, _userGroup);
         if (Directory.Exists(path))         
            preference.Files.TemplateDwgPath = path + "\\";

         //PageSetupOverridesTemplateFile
         path = Path.Combine(_localSettingsFolder, _settPikFile.PathVariables.QNewTemplateFile.Value, _userGroup, _userGroup + ".dwt");
         if (File.Exists(path))
         {            
            preference.Files.PageSetupOverridesTemplateFile = path; //Path.Combine(_localSettingsFolder, _settPikFile.PathVariables.PageSetupOverridesTemplateFile.Value);                                                                           
            preference.Files.QNewTemplateFile = path;
         }

         //SheetSetTemplatePath
         path = Path.Combine(_localSettingsFolder, _settPikFile.PathVariables.SheetSetTemplatePath.Value, _userGroup);
         if (Directory.Exists(path))
            Env.SetEnv("SheetSetTemplatePath", path);
         //con.OpenSubsection("General").WriteProperty("SheetSetTemplatePath", path);
         //con.OpenSubsection("General").WriteProperty("TemplatePath", _settingsPIK.PathVariables.SheetSetTemplatePath.Value + _userGroup);

         foreach (var sysVar in _settPikFile.SystemVariables)
         {
            SetSystemVariable(sysVar.Name, sysVar.Value, sysVar.IsReWrite);
         }

         FlexBrics.Setup();
      }

      private List<Variable> GetPaths(List<Variable> pathVars1, List<Variable> pathVars2)
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
      private string GetPathVariable(List<Variable> settings, string path, string group)
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
               fullPath += path;
            }
            else
            {
               fullPath = fullPath.Remove(fullPath.Length - 1, 1);
            }
         }
         catch { }

         return fullPath;
      }

      private void SetSystemVariable(string name, string value, bool isReWrite)
      {
         if (name.ToUpper() == "TRUSTEDPATHS")
         {
            var paths = value.Split(';').Select(p => Path.Combine(_localSettingsFolder, p));
            value = String.Join(";", paths);
         }

         object nameVar = null;
         try
         {
            nameVar = AutoCadApp.GetSystemVariable(name);
         }
         catch
         {
            return;
         }
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
               Log.Error("SetSystemVariable " + name, ex);
            }
         }
      }

      /// <summary>
      /// Дефолтные настройки профиля.
      /// Возможно лучше вообще ничего не делать, чем портить профиль человеку. Один фиг работать не будеть.!!!???
      /// </summary>
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