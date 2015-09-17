﻿using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace UpdatePIKManager
{
   class Program
   {
      static void Main(string[] args)
      {
         try
         {
            Trace.WriteLine(string.Format("UpdatePIKManager запущен. {0}", DateTime.Now));
            Trace.WriteLine("Аргументы: " + String.Join(" ", args));
            if (args == null || args.Length < 2)
            {
               Trace.WriteLine("Неверные аргументы");
               return;
            }

            string sourceFile = Path.GetFullPath(args[0]);// @"Z:\AutoCAD_server\Адаптация\Dll\AutoCAD_PIK_Manager.dll";
            Trace.WriteLine("sourceFile " +sourceFile);
            string destFile = Path.GetFullPath(args[1]);// @"C:\Autodesk\AutoCAD\PIK\Dll\AutoCAD_PIK_Manager.dll";
            Trace.WriteLine("destFile " +destFile);

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
                  File.Copy(sourceFile, destFile, true);
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
         catch (Exception ex)
         {
            Trace.WriteLine("Исключение: " + ex.Message);
         }
      }
   }
}
