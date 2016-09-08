using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace UpdatePIKManager
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            // args[0] - серверный файл AutoCAD_PIK_Manager (Z:\AutoCAD_server\Адаптация\Dll\AutoCAD_PIK_Manager.dll)
            // args[1] - локальный файл AutoCAD_PIK_Manager (C:\Autodesk\AutoCAD\PIK\Dll\AutoCAD_PIK_Manager.dll)
            // args[2] - RoamablyRootFolder - C:\Users\khisyametdinovvt\AppData\Roaming\Autodesk\AutoCAD 2016\R20.1\rus\
            // args[3] - имя профиля (ПИК)
            Trace.WriteLine(string.Format("UpdatePIKManager запущен. {0}", DateTime.Now));
            Trace.WriteLine(string.Format("Кол аргументов {0}", args.Length));
            Trace.WriteLine("Аргументы: " + String.Join(" ", args));
            //if (args == null || args.Length < 2)
            //{
            //   Trace.WriteLine("Неверные аргументы");
            //   return;
            //}

            try
            {
                Update(args);
                // Сброс сортировки кнопок в инструментальных палитрах.            
                if (args.Length >= 4)
                {
                    SortingToolPalette sortPalette = new SortingToolPalette(args[2], args[3]);                    
                    sortPalette.ResetSorting();
                    Trace.WriteLine("Сортировка кнопок в палитрах сброшена");
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine("Исключение: " + ex.Message);
            }
        }

        private static void Update(string[] args)
        {
            string sourceFile = string.Empty;
            string destFile = string.Empty;
            try
            {
                sourceFile = Path.GetFullPath(args[0]);// @"Z:\AutoCAD_server\Адаптация\Dll\AutoCAD_PIK_Manager.dll";
            }
            catch
            {
                sourceFile = @"\\dsk2.picompany.ru\project\CAD_Settings\AutoCAD_server\Адаптация\Dll\AutoCAD_PIK_Manager.dll";
            }
            try
            {
                destFile = Path.GetFullPath(args[1]);// @"C:\Autodesk\AutoCAD\PIK\Dll\AutoCAD_PIK_Manager.dll";
            }
            catch
            {
                destFile = Path.Combine(Path.GetPathRoot(Environment.SystemDirectory), @"Autodesk\AutoCAD\Pik\Settings\Dll\AutoCAD_PIK_Manager.dll");
            }
            Trace.WriteLine("sourceFile " + sourceFile);            
            Trace.WriteLine("destFile " + destFile);

            if (!File.Exists(sourceFile) || !File.Exists(destFile))
            {
                Trace.WriteLine("Не существует одного из путей.");
                return;
            }

            int i = 0;
            while (i < 5)
            {
                try
                {
                    // Удаление epplus.dll - скрытого
                    DeleteHidenEpplus(destFile);
                    //File.Copy(sourceFile, destFile, true);
                    // копирование файлов в папке
                    copyFiles(Path.GetDirectoryName(sourceFile), Path.GetDirectoryName(destFile));
                    Trace.WriteLine("Скопировалось");
                    break;
                }
                catch
                {
                    Thread.Sleep(5000);//Подождать пока автокад закроется (2015 закрывается очень долго).
                    i++;
                }
            }
        }

        private static void DeleteHidenEpplus(string destFile)
        {
            string fileEpplus = Path.Combine(Path.GetDirectoryName(destFile), "EPPlus.dll");
            if (File.Exists(fileEpplus))
            {
                try
                {
                    var attrs = File.GetAttributes(fileEpplus);
                    if ((attrs & FileAttributes.Hidden) == FileAttributes.Hidden)
                    {
                        Trace.WriteLine("Epplus скрытый");
                        File.SetAttributes(fileEpplus, FileAttributes.Normal);
                    }
                    Trace.WriteLine("Удаление Epplus");
                    File.Delete(fileEpplus);
                }
                catch (Exception ex)
                {
                    Trace.Write(ex.ToString());
                }
            }
            else
            {
                Trace.WriteLine(string.Format("Файла Epplus не существует - {0}", fileEpplus));
            }
        }

        private static void copyFiles(string sourseDir, string destDir)
        {
            Trace.WriteLine(string.Format("Копирование файлов из {0} в {1}", sourseDir, destDir));
            var dirSource = new DirectoryInfo(sourseDir);
            var filesSource = dirSource.GetFiles("*.dll");
            foreach (var file in filesSource)
            {
                Trace.WriteLine(string.Format("Копирование файла из {0} в {1}", file.FullName, Path.Combine(destDir, file.Name)));
                file.CopyTo(Path.Combine(destDir, file.Name), true);
            }
        }
    }
}