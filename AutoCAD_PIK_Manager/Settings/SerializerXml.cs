using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

namespace AutoCAD_PIK_Manager.Settings
{
   internal class SerializerXml
   {
      private string _settingsFile;      

      public SerializerXml(string settingsFile)
      {
         _settingsFile = settingsFile;         
      }

      public void SerializeList<T>(T settings)
      {
         using (FileStream fs = new FileStream(_settingsFile, FileMode.Create, FileAccess.Write))
         {
            XmlSerializer ser = new XmlSerializer(typeof(T) );
            ser.Serialize(fs, settings);
         }
      }

      public T DeserializeXmlFile<T>()
      {         
         XmlSerializer ser = new XmlSerializer(typeof(T));
         using (XmlReader reader = XmlReader.Create(_settingsFile))
         {
            try
            {
               return (T)ser.Deserialize(reader);
            }
            catch (Exception ex)
            {
               Log.Fatal("DeserializeXmlFile " + _settingsFile, ex);
               throw;
            }            
         }         
      }
   }
}