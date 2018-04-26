using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace AnythingToPPTX.Utils
{
    public class FileUtils
    {
        public static void DeleteFolder(string directoryPath)
        {
            foreach (string d in Directory.GetFileSystemEntries(directoryPath))
            {
                if (File.Exists(d))
                {
                    FileInfo fi = new FileInfo(d);
                    if (fi.Attributes.ToString().IndexOf("ReadOnly") != -1)
                        fi.Attributes = FileAttributes.Normal;
                    File.Delete(d);
                }
                else
                    DeleteFolder(d);
            }
            Directory.Delete(directoryPath);
        }
    }
}
