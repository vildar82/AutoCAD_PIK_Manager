﻿using System;
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
        public const string Group = "PIK";
        private static string _about;
        private static string _err = string.Empty;

        public static readonly string SystemDriveName = Path.GetPathRoot(Environment.SystemDirectory);

        private static string About
        {
            get
            {
                if (_about == null)
                    _about = "\nПрограмма настройки AutoCAD_Pik_Manager, версия: " + Assembly.GetExecutingAssembly().GetName().Version +
                      "\nПользоватль: " + Environment.UserName + ", Группа: " + PikSettings.UserGroup + 
                      $"\nПуть к серверу настроек = {PikSettings.ServerSettingsFolder}";
                return _about;
            }
        }

        [CommandMethod(Group, "PIK_Manager_About", CommandFlags.Modal)]
        public static void AboutCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\n{0}", About);
            if (!string.IsNullOrEmpty(_err))
            {
                doc.Editor.WriteMessage("\n{0}", _err);
            }
        }

        public void Initialize()
        {
            // Исключения в Initialize проглотит автокад, без выдачи сообщений. При этом сборка не будет загружена!!!.         
            try
            {
                PikSettings.LoadSettings();
                // Запись в лог
                Log.Info($"AutoCAD_PIK_Manager загружен. Версия {Assembly.GetExecutingAssembly().GetName().Version}. Настройки загружены из {PikSettings.CurDllLocation}");
                Log.Info($"Путь к сетевой папке настроек - {PikSettings.ServerSettingsFolder}");
                Log.Info($"Версия автокада - {Application.Version.ToString()}");
                Log.Info($"Версия среды .NET Framework - {Environment.Version}");

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
                        _err += ex.Message;
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
                        _err += ex.Message;
                    }
                }
                try
                {
                    Profile profile = new Profile();
                    // Настройка профиля ПИК в автокаде
                    if (PikSettings.UserGroup == "ГП")
                    {
                        profile.SetProfilePIK = false;
                        profile.SetTemplate = false;
                        profile.SetToolPalette = false;
                        profile.SetSysVars = PikSettings.GroupFileSettings?.SystemVariables;
                    }
                    profile.SetProfile();
                    Log.Info("Профиль настроен.");
                    //else
                    //{
                    //// Загрузка сбороки ГП                        
                    //string gpdll = Path.Combine(PikSettings.LocalSettingsFolder, @"Script\NET\ГП\PIK_GP_Acad.dll");
                    //LoadDll.Load(gpdll);                        
                    //}
                }
                catch (System.Exception ex)
                {
                    Log.Error(ex, "Ошибка настройки профиля SetProfile().");
                    _err += ex.Message;
                }
            }
            catch (Settings.Exceptions.NoGroupException)
            {
                // Пользователь без группы.
                return;
            }
            catch (System.Exception ex)
            {
                Log.Error(ex, "LoadSettings");
                Log.Info($"AutoCAD_PIK_Manager загружен с ошибками. Версия {Assembly.GetExecutingAssembly().GetName().Version}. Настройки не загружены из {PikSettings.CurDllLocation}");
                Log.Info($"Версия автокада - {Application.Version.ToString()}");
                Log.Info($"Путь к сетевой папке настроек - {PikSettings.ServerSettingsFolder}");
                _err += ex.Message;
            }

            // Загрузка библиотек
            try
            {
                LoadDll.Load(Path.Combine(PikSettings.CurDllLocation, "AcadLib.dll"));
            }
            catch (System.Exception ex)
            {
                Log.Error(ex, "Ошибка загрузки библиотеки AcadLib.");
                _err += ex.Message;
            }
        }

        public void Terminate()
        {
            // Обновление программы (копирование AutoCAD_PIK_Manager.dll)
            string updater = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "UpdatePIKManager.exe");
            string sourceDllPikManager = string.Empty;
            string destDllPikManager = string.Empty;
            try
            {
                sourceDllPikManager = Path.Combine(PikSettings.ServerSettingsFolder, "Dll\\AutoCAD_PIK_Manager.dll");
            }
            catch
            {
                sourceDllPikManager = @"\\dsk2.picompany.ru\project\CAD_Settings\AutoCAD_server\Адаптация\Dll\AutoCAD_PIK_Manager.dll";
            }
            try
            {
                destDllPikManager = Path.Combine(PikSettings.LocalSettingsFolder, "Dll\\AutoCAD_PIK_Manager.dll");
            }
            catch
            {
                destDllPikManager = Path.Combine(SystemDriveName, @"Autodesk\AutoCAD\Pik\Settings\Dll\AutoCAD_PIK_Manager.dll");
            }
            string roamableFolder = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.RoamableRootFolder.TrimEnd(new char[] { '\\', '/' });
            string arg = string.Format("\"{0}\" \"{1}\" \"{2}\" \"{3}\"", sourceDllPikManager, destDllPikManager, roamableFolder, PikSettings.PikFileSettings.ProfileName);
            Log.Info("Запущена программа обновления UpdatePIKManager с аргументом: {0}", arg);
            ProcessStartInfo startInfo = new ProcessStartInfo(updater);
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.Arguments = arg;
            Process.Start(startInfo);
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