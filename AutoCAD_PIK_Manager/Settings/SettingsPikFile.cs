﻿using System;
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
      public string ServerSettingsPath { get; set; }//Путь к папке настроек на сервере: z:\AutoCAD_server\Адаптация
      public string PathToUserList { get; set; }//Путь к файлу со списком ФИО и соответствующих шифров отделов: z:\Settings\Users\UserList2.xlsx
      public string ServerLogPath { get; set; }//Путь к папке z:\AutoCAD_server\ShareSettings\AutoCAD_PIK_Manager\Logs

      public PathVariable PathVariables { get; set; }// Пути доступа (support paths, printers, templates)
      public List<SystemVariable> SystemVariables { get; set; }//Системные переменные
      public List<Group> Groups { get; set; }//Группы (Шифр отдела - Полное имя отдела).
      public string NameCADManager { get; set; }//Имя CAD-менеджера
      public string MailCADManager { get; set; }
      public string SubjectMail { get; set; }
      public string BodyMail { get; set; }
   }
}