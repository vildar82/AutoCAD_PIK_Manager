using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace AutoCAD_PIK_Manager.Services
{
   /// <summary>
   /// Загрузка вспомогательных сборок
   /// </summary>
   public static class LoadService 
   {
      private static bool isLoadedSpecBlocks = false;

      /// <summary>
      /// Загрузка сборки SpecBlocks.dll
      /// </summary>
      public static void LoadSpecBlocks()
      {
         if (isLoadedSpecBlocks)
         {
            return;
         }
         // Загрузка сборки SpecBlocks
         var dllSpecBlocks = Path.Combine(Settings.PikSettings.LocalSettingsFolder, @"Script\NET\SpecBlocks\SpecBlocks.dll");
         if (File.Exists(dllSpecBlocks))
         {
            try
            {
               Assembly.LoadFrom(dllSpecBlocks);
               isLoadedSpecBlocks = true;
            }
            catch (Exception ex)
            {
               throw ex;
            }
         }
         else
         {
            throw new Exception($"Не найден файл {dllSpecBlocks}.");
         }
      }
   }
}
