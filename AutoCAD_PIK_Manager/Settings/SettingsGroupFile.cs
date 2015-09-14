using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCAD_PIK_Manager.Settings
{
   /// <summary>
   /// Настройки отдела. Индивидуальные особенности отдела.
   /// </summary>
   [Serializable]
   public class SettingsGroupFile
   {
      // Дополнительные пути доступа
      public PathVariable PathVariables { get; set; }// Пути доступа (support paths, printers, templates)
      public List<SystemVariable> SystemVariables { get; set; }//Системные переменные
   }
}
