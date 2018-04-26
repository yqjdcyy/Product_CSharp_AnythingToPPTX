using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using AnythingToPPTX.Utils;

namespace ImageToPPTX
{
    class Program
    {
        static void Main(string[] args)
        {
            if(args == null || args.Length < 2)
            {
                Console.WriteLine("error:第一个参数为转换文件列表，以逗号格开，不能省略");
                return;
            }
            string path = args[1];
            string fileList = args[0];
            var files = GetFileList(fileList);
            var ppt = ImageToPPTX(path, files);
            Console.WriteLine("return:{0}",ppt);
        }

        private static string ImageToPPTX(string path, List<string> fileList)
        {
            if(fileList.Count <= 0)
            {
                return string.Empty;
            }
            ImageMegerToPPTXUtils converter = new ImageMegerToPPTXUtils();

            return converter.convert(path, fileList);
        }

        public static List<string> GetFileList(string strFileList)
        {
            if (string.IsNullOrEmpty(strFileList))
            {
                return new List<string>(0);
            }
            return new List<string>(strFileList.Split(','));
        }
    }
}
