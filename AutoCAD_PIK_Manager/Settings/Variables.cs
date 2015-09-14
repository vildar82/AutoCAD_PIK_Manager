using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCAD_PIK_Manager.Settings
{
   public abstract class Variables
   {
      public string Name { get; set; }
      public string Value { get; set; }
      public bool IsReWrite { get; set; }      
   }
}
