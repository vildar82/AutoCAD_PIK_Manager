using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using AutoCAD_PIK_Manager.Settings;
using Autodesk.AutoCAD.Runtime;
using NLog;
using NLog.Config;

[assembly: ExtensionApplication(typeof(AutoCAD_PIK_Manager.Commands))]
[assembly: CommandClass(typeof(AutoCAD_PIK_Manager.Commands))]

namespace AutoCAD_PIK_Manager
{
   public class Commands : IExtensionApplication
   {
      public void Initialize()
      {
         // Програ загружена в автокад.

         // Запись в лог
         Log.Info("AutoCAD_PIK_Manager загружен - Initialize() start.");         
         
         PikSettings.LoadSettings(); 
         
         // Если есть другие запущеннык автокады, то пропускаем копирование файлов с сервера, т.к. многие файлы уже заняты другим процессом автокада.
         if (!IsProcessAny())
         {
            // Обновление настроек с сервера (удаление и копирование)
            PikSettings.UpdateSettings();
            Log.Info("Настройки обновлены");
         }
         // Настройка профиля ПИК в автокаде
         Profile profile = new Profile();
         profile.SetProfile();
      }

      public void Terminate()
      {
         try
         {
            // Обновление программы (копирование AutoCAD_PIK_Manager.dll)
            //Process.Start(@"C:\Autodesk\AutoCAD\Pik\Settings\Dll\UpdatePIKManager.exe");
         }
         catch { }
      }

      private static bool IsProcessAny()
      {
         //logger.Info("IsProcessAny");
         Process[] acadProcess = Process.GetProcessesByName("acad");
         if (acadProcess.Count() > 1) return true;
         return false;
      }
   }
}