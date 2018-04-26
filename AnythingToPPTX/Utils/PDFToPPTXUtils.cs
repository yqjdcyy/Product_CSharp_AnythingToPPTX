using AnythingToPPTX.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using System.IO;
using Ghostscript.NET;
using Spire.Pdf;

namespace AnythingToPPTX.Utils
{
    public class PDFToPPTXUtils
    {
        public List<PPTPage> convert(string pdfPath, string outputPath)
        {
            List<PPTPage> pageList = new List<PPTPage>();

            if (string.IsNullOrEmpty(pdfPath) || !System.IO.File.Exists(pdfPath))
                throw new Exception("pdf is not exists");
            if (string.IsNullOrEmpty(outputPath))
                throw new Exception("output path is not confirm");

            outputPath = System.IO.Path.Combine(outputPath, System.IO.Path.GetFileName(pdfPath));
            if (System.IO.Directory.Exists(outputPath))
                FileUtils.DeleteFolder(outputPath);
            System.IO.Directory.CreateDirectory(outputPath);

            pageList = SplitePDF(pdfPath, outputPath);

            return pageList;
        }

        /// <summary> BYTESCOUT_MANY_IMAGE_FROMAT 均有头像重复异常
        /// 
        ///     BMP     26368KB 4000x 2250  小
        ///     GIF     3005KB  4000x 2250  小
        ///     JPEG    880KB   4000x 2250  大
        ///     PNG     3245KB  4000x 2250  大
        ///     TIFF    4589KB  4000x 2250  大
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="outputPath"></param>
        /// <returns></returns>
        List<PPTPage> bytescoutPDF_Type_Support(string filepath, string outputPath)
        {
            List<PPTPage> pages = new List<PPTPage>();
            Bytescout.PDFRenderer.RasterRenderer render = new Bytescout.PDFRenderer.RasterRenderer();

            render.LoadDocumentFromFile(filepath);

            render.RegistrationKey = "demo";
            render.RegistrationName = "demo";

            int length = render.GetPageCount();

            render.RenderPageToFile(14, Bytescout.PDFRenderer.RasterOutputFormat.BMP, "pdf_14.bmp");
            render.RenderPageToFile(14, Bytescout.PDFRenderer.RasterOutputFormat.GIF, "pdf_14.gif");
            render.RenderPageToFile(14, Bytescout.PDFRenderer.RasterOutputFormat.JPEG, "pdf_14.jpeg");
            render.RenderPageToFile(14, Bytescout.PDFRenderer.RasterOutputFormat.PNG, "pdf_14.png");
            render.RenderPageToFile(14, Bytescout.PDFRenderer.RasterOutputFormat.TIFF, "pdf_14.tiff");

            return pages;
        }

        /// <summary> BYTESCOUT_JPEG 头像重复异常
        /// 分辨率：4000x 2250
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="outputPath"></param>
        /// <returns></returns>
        List<PPTPage> bytescoutPDF(string filepath, string outputPath)
        {
            List<PPTPage> pages = new List<PPTPage>();
            Bytescout.PDFRenderer.RasterRenderer render = new Bytescout.PDFRenderer.RasterRenderer();

            render.LoadDocumentFromFile(filepath);

            render.RegistrationKey = "demo";
            render.RegistrationName = "demo";

            int length = render.GetPageCount();
            for (int i = 0; i < length; )
            {
                string path= string.Format("{0}/img-{1}.jpeg", outputPath, i);

                System.Drawing.Image img = render.RenderPageToImage(i);
                img.Save(path, System.Drawing.Imaging.ImageFormat.Jpeg);

                pages.Add(new PPTPage() { Cover = path });
                Console.WriteLine("PDF TO IMAGES - {0}/{1}", ++i, length);
            }

            return pages;
        }

        /// <summary> SPIRE_JPEG 左上角有水印
        ///     分辨率：1280x 720 
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="outputPath"></param>
        /// <returns></returns>
        List<PPTPage> spirePDF(string filepath, string outputPath)
        {
            List<PPTPage> pages = new List<PPTPage>();
            Spire.Pdf.PdfDocument doc = new Spire.Pdf.PdfDocument(filepath);

            int pageCount = doc.Pages.Count;
            int width = (int)doc.PageSettings.Width;
            int height = (int)doc.PageSettings.Height;

            for (int i = 0; i < pageCount;)
            {
                string path = string.Format("{0}/img-{1}.jpeg", outputPath, i);

                System.Drawing.Image img = doc.SaveAsImage(i);
                img.Save(path, System.Drawing.Imaging.ImageFormat.Jpeg);

                pages.Add(new PPTPage() { Cover = path });
                Console.WriteLine("PDF TO IMAGES - {0}/{1}", ++i, pageCount);
            }

            return pages;
        }

        /// <summary> iTextSharp_PNG 无法正常进行展出操作
        /// 
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="outputPath"></param>
        /// <returns></returns>
        List<PPTPage> SplitePDF(string filepath, string outputPath)
        {
            List<PPTPage> pages = new List<PPTPage>();
            iTextSharp.text.pdf.PdfReader reader = null;
            int currentPage = 1;
            int pageCount = 0;

            System.Text.UTF8Encoding encoding = new System.Text.UTF8Encoding();
            reader = new iTextSharp.text.pdf.PdfReader(filepath);
            reader.RemoveUnusedObjects();
            pageCount = reader.NumberOfPages;

            for (int i = 1; i <= pageCount; )
            {
                string outfile = System.IO.Path.Combine(outputPath, i + ".png");
                iTextSharp.text.Document doc = new iTextSharp.text.Document(reader.GetPageSizeWithRotation(currentPage));
                iTextSharp.text.pdf.PdfCopy pdfCpy = new iTextSharp.text.pdf.PdfCopy(doc, new System.IO.FileStream(outfile, System.IO.FileMode.Create));
                doc.Open();
                for (int j = 1; j <= 1; j++)
                {
                    iTextSharp.text.pdf.PdfImportedPage page = pdfCpy.GetImportedPage(reader, currentPage);
                    pdfCpy.SetFullCompression();
                    pdfCpy.AddPage(page);
                    currentPage += 1;
                }

                Console.WriteLine("PDF TO IMAGES - {0}/{1}", ++i, pageCount);
                pages.Add(new PPTPage() { Cover = outfile });

                pdfCpy.Flush();
                doc.Close();
                pdfCpy.Close();
                reader.Close();
            }

            return pages;
        }

        public void LoadImage(string filepath, string destpath)
        {
            PdfReader reader = new iTextSharp.text.pdf.PdfReader(filepath);
            GhostscriptPngDevice dev = new GhostscriptPngDevice(GhostscriptPngDeviceType.Png256);
            dev.GraphicsAlphaBits = GhostscriptImageDeviceAlphaBits.V_4;
            dev.TextAlphaBits = GhostscriptImageDeviceAlphaBits.V_4;
            dev.ResolutionXY = new GhostscriptImageDeviceResolution(290, 290);
            dev.InputFiles.Add(filepath);
            dev.Pdf.FirstPage = 0;
            dev.Pdf.LastPage = reader.NumberOfPages;
            dev.CustomSwitches.Add("-dDOINTERPOLATE");
            dev.OutputPath = destpath + "%03d.jpg";
            dev.Process();
        }

        private void Process(string input, string output, int startPage, int endPage)
        {
            GhostscriptVersionInfo _gs_verssion_info = GhostscriptVersionInfo.GetLastInstalledVersion();
            Ghostscript.NET.Processor.GhostscriptProcessor processor = new Ghostscript.NET.Processor.GhostscriptProcessor(_gs_verssion_info, true);
            processor.StartProcessing(CreateTestArgs(input, output, startPage, endPage), new ConsoleStdIO(true, true, true));
        }

        private string[] CreateTestArgs(string inputPath, string outputPath, int pageFrom, int pageTo)
        {
            List<string> gsArgs = new List<string>();

            gsArgs.Add("-q");
            gsArgs.Add("-dSAFER");
            gsArgs.Add("-dBATCH");
            gsArgs.Add("-dNOPAUSE");
            gsArgs.Add("-dNOPROMPT");
            gsArgs.Add(@"-sFONTPATH=" + System.Environment.GetFolderPath(System.Environment.SpecialFolder.Fonts));
            gsArgs.Add("-dFirstPage=" + pageFrom.ToString());
            gsArgs.Add("-dLastPage=" + pageTo.ToString());
            gsArgs.Add("-sDEVICE=png16m");
            gsArgs.Add("-r72");
            gsArgs.Add("-sPAPERSIZE=a4");
            gsArgs.Add("-dNumRenderingThreads=" + Environment.ProcessorCount.ToString());
            gsArgs.Add("-dTextAlphaBits=4");
            gsArgs.Add("-dGraphicsAlphaBits=4");
            gsArgs.Add(@"-sOutputFile=" + outputPath);
            gsArgs.Add(@"-f" + inputPath);

            return gsArgs.ToArray();
        }
    }
}
