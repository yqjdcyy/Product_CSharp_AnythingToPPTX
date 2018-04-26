using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnythingToPPTX.Utils
{
    class ImageInfoUtils
    {

        public static int RATE = 9525;

        public List<String> filter(List<String> imgList)
        {
            List<String> leftImgList = new List<string>();
            string[] extList = { ".jpg", ".jpeg", ".png" };

            foreach (var path in imgList)
            {
                if (String.IsNullOrEmpty(path))
                    continue;

                if (!File.Exists(path))
                    continue;

                String ext = Path.GetExtension(path).ToLower();
                if (!extList.Contains(ext))
                    continue;

                leftImgList.Add(path);
            }

            return leftImgList;
        }

        public Size listMaxSize(List<String> imgList)
        {
            Size max = new Size();
            List<String> leftImgList = filter(imgList);

            foreach (var path in leftImgList)
            {
                Size cur = getSize(path);
                max.Height = max.Height >= cur.Height ? max.Height : cur.Height;
                max.Width = max.Width >= cur.Width ? max.Width : cur.Width;
            }

            return max;
        }

        public Size listMaxPPTSize(List<String> imgList)
        {
            Size max = listMaxSize(imgList);

            max.Width *= RATE;
            max.Height *= RATE;
            return max;
        }

        public Size getSize(String imgPath)
        {
            try
            {
                using (System.Drawing.Image img = System.Drawing.Image.FromFile(imgPath))
                {
                    return img.Size;
                }
            }
            catch (Exception) { }
            return new Size() { Width = 0, Height = 0 };
        }

        public Size getPPTSize(String imgPath)
        {
            try
            {
                using (System.Drawing.Image img = System.Drawing.Image.FromFile(imgPath))
                {
                    return new Size()
                    {
                        Width = img.Width * RATE,
                        Height = img.Height * RATE
                    };
                }
            }
            catch (Exception) { }
            return new Size() { Width = 0, Height = 0 };
        }
    }
}
