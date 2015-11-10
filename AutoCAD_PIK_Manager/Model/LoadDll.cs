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
      public static void Load(string file)
      {
         if (File.Exists(file))
         {
            Assembly.LoadFile(file);
            //AcadLib.Comparers.StringsNumberComparer comparer = new AcadLib.Comparers.StringsNumberComparer ();
         }
      }
   }
}
