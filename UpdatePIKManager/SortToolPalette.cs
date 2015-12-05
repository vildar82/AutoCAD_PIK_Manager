using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.XPath;

namespace UpdatePIKManager
{
   /// <summary>
   /// Сортировка кнопок на инструментальной палитре в профиле ПИК
   /// Автокад должен быть закрыт.
   /// </summary>
   public class SortingToolPalette
   {
      string acadRoamableRootFolder;
      string profileName;

      public SortingToolPalette(string acadRoamableRootFolder, string profileName)
      {
         this.acadRoamableRootFolder = acadRoamableRootFolder;
         this.profileName = profileName;
      }

      /// <summary>
      /// Сброс порядка инструментов во всех палитрах в профиле
      /// </summary>            
      public void ResetSorting()
      {
         string awsProfile = getProfileFile();
         if (!File.Exists(awsProfile))
         {
            throw new Exception(string.Format("Файл aws профиля {0} не найден. Путь поиска {1}", profileName, awsProfile));
         }
         ResetToolOrder(awsProfile);
      }

      private string getProfileFile()
      {
         return Path.Combine(acadRoamableRootFolder, "Support\\Profiles", profileName, "Profile.aws");
      }

      public static void ResetToolOrder(string awsProfile)
      {
         // Удаление элементов порядка инструментов из файла aws во всех палитрах
         // Profile\StorageRoot\ToolPaletteScheme\ToolPaletteSets\ToolPaletteSet\CAcTcUiToolPaletteSet
         // \ToolPalettes\CAcTcUiToolPalette\CatalogView\ToolOrder - удалить все элементы Tool
         XDocument doc = XDocument.Load(awsProfile);
         var palettes = doc.XPathSelectElement("/Profile/StorageRoot/ToolPaletteScheme/ToolPaletteSets/ToolPaletteSet/CAcTcUiToolPaletteSet/ToolPalettes")?.Descendants("ToolPalette");
         if (palettes == null)
         {
            throw new Exception("Не найдены элементы ToolPalette в файле aws");
         }
         // /ToolPalette/CAcTcUiToolPalette/CatalogView/ToolOrder
         foreach (var item in palettes)
         {
            var toolOrder = item.XPathSelectElement("CAcTcUiToolPalette/CatalogView/ToolOrder");
            if (toolOrder != null)
            {
               toolOrder.RemoveNodes();
            }
         }
         doc.Save(awsProfile);
      }
   }
}
