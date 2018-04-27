using AnythingToPPTX.Entity;
using AnythingToPPTX.Utils;
using System;
using System.Collections.Generic;
using System.IO;

namespace AnythingToPPTX
{
    class Program
    {
        static void Main(string[] args)
        {
            // check
            if (args.Length <= 0)
            {
                Console.WriteLine("*.exe images.path, pptx.store.path, template.path, isOpenFolderAfterConvert");
                return;
            }

            // print
            int size = args.Length;
            Console.WriteLine("Arguments=");
            for (int i = 0; i < size; i++)
            {
                Console.WriteLine("\t{0}", args[i]);
            }

            // init
            String path = args[0];
            String dest = null;
            String template = null;
            bool bOpenFolder = true;
            if (size > 1)
                dest = args[1];
            if (size > 2)
                template = args[2];
            if (size > 3)
                bOpenFolder = Boolean.Parse(args[3]);


            // convert
            try
            {
                ImageToPPTX(path, dest, template, bOpenFolder);
            }
            catch (Exception e)
            {
                Console.WriteLine(String.Format("Convert fail: {0}", e.StackTrace));
            }
        }

        private static void PDFToPPTX()
        {
            PDFToPPTXUtils utils = new PDFToPPTXUtils();
            ImageToPPTXTemplateUtils converter = new ImageToPPTXTemplateUtils();
            string path = @"D:\资料备份\资料\工作\测试\ppt\customer\GMIC\pdf";
            utils.convert(path + "\\1-（非公开）-简仁贤-竹间智能.pdf", path + "\\image");
            //converter.convert(path + "\\ppt", utils.convert(path + "\\ZMER_BP.pdf", path + "\\image"));
            //System.Diagnostics.Process.Start("explorer.exe", path + "\\ppt");
        }

        private static void PDFToPPTX_FromJava()
        {
            ImageToPPTXTemplateUtils converter = new ImageToPPTXTemplateUtils();
            List<PPTPage> pageList = new List<PPTPage>();
            string path = @"D:\资料备份\资料\工作\测试\pdf\ppt";

            pageList.Add(new PPTPage() { Cover = @"D:\资料备份\资料\工作\测试\pdf\image\ZMER_BP.pdf_java\dpi-96-jpeg\0014.jpeg", PageUrlList = new List<PageUrl>() });
            pageList.ForEach(p =>
            {
                p.PageUrlList.Add(new PageUrl() { url = "http://www.sina.com", angle = 0, origin = new System.Drawing.Size() { Width = 0, Height = 0 }, size = new System.Drawing.Size() { Width = 100, Height = 50 } });
                p.PageUrlList.Add(new PageUrl() { url = "http://www.google.com", angle = 0, origin = new System.Drawing.Size() { Width = 100, Height = 100 }, size = new System.Drawing.Size() { Width = 100, Height = 100 } });
            });
            
            converter.convert(path, pageList);
        }

        private static void ImageToPPTX()
        {

            List<String> imgList = new List<String>();
            ImageToPPTXTemplateUtils converter = new ImageToPPTXTemplateUtils();
            string path = @"D:\data\soft\wechat\WeChat Files\yqjdcyy\Files\dpi-96-jpeg";

            foreach (string d in Directory.GetFileSystemEntries(path))
            {
                if (!Directory.Exists(d) && File.Exists(d))
                {
                    imgList.Add(d);
                }
            }
            converter.convert(path + @"\output", imgList);
            System.Diagnostics.Process.Start("explorer.exe", path + "\\output");
        }

        private static void ImageToPPTX(String path, String dest, String temp, Boolean bOpen)
        {
            // init
            List<String> list = new List<String>();
            ImageToPPTXTemplateUtils converter = new ImageToPPTXTemplateUtils();
            dest = String.IsNullOrEmpty(dest) ? (path + @"\output") : dest;

            // switch
            if (File.Exists(path))
            {
                list.Add(path);
            }else if (Directory.Exists(path))
            {
                foreach (string d in Directory.GetFileSystemEntries(path))
                {
                    if (!Directory.Exists(d) && File.Exists(d))
                    {
                        list.Add(d);
                    }
                }
            }
            else
            {
                throw new Exception("错误的文件路径");
            }

            converter.convert(dest, list, temp);
            if (bOpen)
                System.Diagnostics.Process.Start("explorer.exe", dest.Replace("\\\\", "\\"));
        }
    }
}
