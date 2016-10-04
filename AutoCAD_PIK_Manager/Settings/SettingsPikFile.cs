using System;
using System.Collections.Generic;

namespace AutoCAD_PIK_Manager.Settings
{
    /// <summary>
    /// Настройки ПИК из файла SettingsPIK.xml
    /// </summary>
    [Serializable]
    public class SettingsPikFile
    {
        public string ProfileName { get; set; }// Имя для профиля в автокаде

        //public string LocalSettingsPath { get; set; }//Путь к папке локольных настроек: Autodesk\AutoCAD\Pik\Settings
        public string ServerSettingsPath { get; set; }  //Путь к папке настроек на сервере: z:\AutoCAD_server\Адаптация        
        public PathVariable PathVariables { get; set; }// Пути доступа (support paths, printers, templates)
        public List<SystemVariable> SystemVariables { get; set; }//Системные переменные
        public List<Group> Groups { get; set; }//Группы (Шифр отдела - Полное имя отдела).
        public string LoginCADManager { get; set; }//логин CAD-менеджера - UserName
        public string NameCADManager { get; set; }//Имя CAD-менеджера      
        public string MailCADManager { get; set; }
        public string SubjectMail { get; set; }
        public string BodyMail { get; set; }

        public static SettingsPikFile Default ()
        {
            SettingsPikFile res = new SettingsPikFile() {
                ProfileName = "ПИК",
                ServerSettingsPath = @"\\dsk2.picompany.ru\project\CAD_Settings\AutoCAD_server\Адаптация",
                PathVariables = new PathVariable {
                    Supports = new List<Variable> {
                            new Variable { Name= "SupportPaths", Value = "Fonts" },
                            new Variable { Name= "SupportPaths", Value = "Support" } },
                    ToolPalettePaths = new List<Variable> { new Variable { Name = "ToolPalettePaths", Value = "ToolPalette", IsReWrite = true } },
                    TemplatePath = new Variable { Name = "TemplatePath", Value = "Template", IsReWrite = true },
                    QNewTemplateFile= new Variable { Name = "QNewTemplateFile", Value = "Template", IsReWrite = true },
                }
            };
            return res;
        }
    }
}