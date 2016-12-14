using System;
using System.Linq;
using System.Collections.Generic;

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
        public bool FlexBricsSetup { get; set; }
        public string FlexBricsFolder { get; set; }

        /// <summary>
        /// Объединение настроек групп в одну общую
        /// </summary>
        /// <param name="sgfs">Группы настроек разных разделов</param>        
        public static SettingsGroupFile Merge(List<SettingsGroupFile> sgfs)
        {
            if (sgfs == null || sgfs.Count == 0) return null;
            if (sgfs.Count == 1) return sgfs[0];

            var f = sgfs[0];
            foreach (var item in sgfs.Skip(1))
            {
                if (item.FlexBricsSetup)
                {
                    f.FlexBricsSetup = true;
                    f.FlexBricsFolder = item.FlexBricsFolder;                    
                }
                f.PathVariables = PathVariable.Merge(f.PathVariables, item.PathVariables);

                if (f.SystemVariables == null)
                {
                    f.SystemVariables = item.SystemVariables;
                }
                else
                {
                    if (item.SystemVariables!= null && item.SystemVariables.Any())
                    {
                        f.SystemVariables.AddRange(item.SystemVariables);
                    }
                }
            }
            return f;
        }
    }
}