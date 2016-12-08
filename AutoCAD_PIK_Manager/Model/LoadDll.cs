using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace AutoCAD_PIK_Manager.Model
{
   public class LoadDll
   {
      public static void LoadTry(string file)
      {
            try
            {
                if (File.Exists(file))
                {
                    Assembly.LoadFrom(file);
                    //AcadLib.Comparers.StringsNumberComparer comparer = new AcadLib.Comparers.StringsNumberComparer ();
                }
            }
            catch (Exception ex)
            {
                try
                {
                    Log.Error(ex, $"Ошибка загрузки файла {file}");
                }
                catch { }                
            }
        }

        public static void LoadRefs ()
        {
            var curDllLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            LoadTry(Path.Combine(curDllLocation, "NLog.dll"));
            LoadTry(Path.Combine(curDllLocation, "EPPlus.dll"));
        }
    }
}
