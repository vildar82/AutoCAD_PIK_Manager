using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCAD_PIK_Manager
{
    public class UpdateFile
    {
        public UpdateFile(FileInfo serverFile, FileInfo localFile, bool updateRequired)
        {
            ServerFile = serverFile;
            LocalFile = localFile;
            UpdateRequired = updateRequired;
        }

        public FileInfo ServerFile { get; set; }
        public FileInfo LocalFile { get; set; }
        public bool UpdateRequired { get; set; }
    }
}
