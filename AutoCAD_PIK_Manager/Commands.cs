using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using AutoCAD_PIK_Manager.Model;
using AutoCAD_PIK_Manager.Settings;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;

[assembly: ExtensionApplication(typeof(AutoCAD_PIK_Manager.Commands))]
[assembly: CommandClass(typeof(AutoCAD_PIK_Manager.Commands))]

namespace AutoCAD_PIK_Manager
{
   public class Commands : IExtensionApplication
   {
      private static string _about;

      private static string About
      {
         get
         {
            if (_about == null)
               _about = "\nПрограмма настройки AutoCAD_Pik_Manager, версия: " + Assembly.GetExecutingAssembly().GetName().Version +
                 "\nПользоватль: " + Environment.UserName + ", Группа: " + PikSettings.UserGroup;
            return _about;
         }
      }

      [CommandMethod("PIK", "PIK_Manager_About", CommandFlags.Modal)]
      public static void AboutCommand()
      {
         Document doc = Application.DocumentManager.MdiActiveDocument;
         doc.Editor.WriteMessage("\n{0}", About);
      }

      public void Initialize()
      {
         // Исключения в Initialize проглотит автокад, без выдачи сообщений.
         // Програ загружена в автокад.
         try
         {
            PikSettings.LoadSettings();
         }
         catch (System.Exception ex)
         {
            Log.Error(ex, "LoadSettings");
            Log.Info("AutoCAD_PIK_Manager загружен с ошибками. Версия {0}. Настройки не загружены из {1}", Assembly.GetExecutingAssembly().GetName().Version, PikSettings.CurDllLocation);
            Log.Info("Версия автокада - {0}", Application.Version.ToString());
            Log.Info("Путь к сетевой папке настроек - {0}", PikSettings.ServerSettingsFolder ?? "нет");
            throw; // Не допускаются ошибки при загрузке настроек. Последствия непредсказуемы. Нужно подойти и разобраться на месте.
         }
         // Запись в лог
         Log.Info("AutoCAD_PIK_Manager загружен. Версия {0}. Настройки загружены из {1}", Assembly.GetExecutingAssembly().GetName().Version, PikSettings.CurDllLocation);
         Log.Info("Путь к сетевой папке настроек - {0}", PikSettings.ServerSettingsFolder ?? "нет");
         Log.Info("Версия автокада - {0}", Application.Version.ToString());

         // Если есть другие запущеннык автокады, то пропускаем копирование файлов с сервера, т.к. многие файлы уже заняты другим процессом автокада.
         if (!IsProcessAny())
         {
            // Обновление настроек с сервера (удаление и копирование)
            try
            {
               PikSettings.UpdateSettings();
               Log.Info("Настройки обновлены.");
            }
            catch (System.Exception ex)
            {
               Log.Error(ex, "Ошибка обновления настроек PikSettings.UpdateSettings();");
            }
            try
            {
               PikSettings.LoadSettings(); // Перезагрузка настроек (могли обновиться файлы настроек на сервере)
                                           // Замена путей к настройкам в файлах инструментальных палитр
               ToolPaletteReplacePath.Replace();
               Log.Info("Настройки загружены.");
            }
            catch (System.Exception ex)
            {
               Log.Error(ex, "Ошибка загрузки настроек PikSettings.LoadSettings();");
            }
         }
         try
         {
            // Настройка профиля ПИК в автокаде
            Profile profile = new Profile();
            profile.SetProfile();
            Log.Info("Профиль ПИК установлен.");
         }
         catch (System.Exception ex)
         {
            Log.Error(ex, "Ошибка настройки профиля SetProfile().");
         }
      }

      public void Terminate()
      {
         // Обновление программы (копирование AutoCAD_PIK_Manager.dll)
         string updater = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "UpdatePIKManager.exe");
         string sourceDllPikManager = Path.Combine(PikSettings.ServerSettingsFolder, "Dll\\AutoCAD_PIK_Manager.dll");
         string destDllPikManager = Path.Combine(PikSettings.LocalSettingsFolder, "Dll\\AutoCAD_PIK_Manager.dll");
         string arg = string.Format("\"{0}\" \"{1}\"", sourceDllPikManager, destDllPikManager);
         Log.Info("Запущена программа обновления UpdatePIKManager с аргументами: sourceDllPikManager - {0}, destDllPikManager - {1}", sourceDllPikManager, destDllPikManager);
         Process.Start(updater, arg);
      }

      private static bool IsProcessAny()
      {
         Process[] acadProcess = Process.GetProcessesByName("acad");
         if (acadProcess.Count() > 1)
         {
            Log.Info("Несколько процессов Acad = {0}", acadProcess.Count());
            return true;
         }
         return false;
      }
   }
}