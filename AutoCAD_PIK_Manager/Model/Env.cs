using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using AutoCadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCAD_PIK_Manager.Model
{ 
   public static class Env
   {
      public static int Ver = AutoCadApp.Version.Major;
      // AutoCAD 2007...2012
      [DllImport("acad.exe", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Unicode, EntryPoint = "acedGetEnv")]
      extern static private Int32 acedGetEnv12(string var, StringBuilder val);
      // AutoCAD 2013...
      [DllImport("accore.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Unicode, EntryPoint = "acedGetEnv")]
      extern static private Int32 acedGetEnv13(string var, StringBuilder val);

      static public string GetEnv(string var)
      {
         StringBuilder val = new StringBuilder(16536);
         if (Ver <= 18) acedGetEnv12(var, val); else acedGetEnv13(var, val);
         return val.ToString();
      }
      
      // AutoCAD 2007...2012
      [DllImport("acad.exe", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Unicode, EntryPoint = "acedSetEnv")]
      extern static private Int32 acedSetEnv12(string var, string val);
      // AutoCAD 2013...
      [DllImport("accore.dll", CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Unicode, EntryPoint = "acedSetEnv")]
      extern static private Int32 acedSetEnv13(string var, string val);

      static public void SetEnv(string var, string val)
      {
         if (Ver <= 12) acedSetEnv12(var, val); else acedSetEnv13(var, val);
      }
   }
}
