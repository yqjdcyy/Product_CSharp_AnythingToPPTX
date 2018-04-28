using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using System.IO;
using DocumentFormat.OpenXml.Validation;
using System.Drawing.Imaging;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using AnythingToPPTX.Entity;

namespace AnythingToPPTX.Utils
{
    public class ImageToPPTXTemplateUtils
    {
        private static String DEFAULT_TEMPLATE_PATH = System.IO.Path.Combine(System.Environment.CurrentDirectory, "Template\\template.pptx");

        public string convert(String path, List<String> imgList)
        {
            return convert(path,imgList,DEFAULT_TEMPLATE_PATH);
        }
        public string convert(String path, List<String> imgList, String tempPath)
        {
            // init
            ImageInfoUtils utils = new ImageInfoUtils();
            imgList = utils.filter(imgList);
            Size max = utils.listMaxPPTSize(imgList);
            tempPath = String.IsNullOrEmpty(tempPath) ? DEFAULT_TEMPLATE_PATH : tempPath;

            // check
            if (0 == max.Width || 0 == max.Height)
                throw new Exception("The format of resource images is invalid");
            if (!File.Exists(tempPath))
                throw new Exception(String.Format("The pointed PPTX template[{0}] is not exists", tempPath));

            return convert(path, tempPath, imgList, max);
        }

        private string convert(String path, String pptTempPath, List<String> imgList, Size size)
        {
            // init
            String filePath = String.Empty;
            while (String.IsNullOrEmpty(filePath) || System.IO.File.Exists(filePath))
            {
                string guid = Guid.NewGuid().ToString();
                filePath = System.IO.Path.Combine(path, guid + ".pptx");
            }

            try
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                    Console.WriteLine(String.Format("Create Folder[{0}]", path));
                }
                File.Copy(pptTempPath, filePath);
                Console.WriteLine(String.Format("Copy Template PPTX from [{0}] to [{1}]", pptTempPath, filePath));
                fill(filePath, imgList);
            }
            catch (Exception ex)
            {
                Console.WriteLine(String.Format("Convert fail: {0}",ex.ToString()));
                //if (File.Exists(filePath))
                    //File.Delete(filePath);
            }

            return filePath;
        }

        private void fill(string filePath, List<string> list)
        {
            int imgId = 915;
            Size max = new ImageInfoUtils().listMaxPPTSize(list);

            using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
            {
                PresentationPart presentationPart = presentationDocument.PresentationPart;

                // reset slide.size
                presentationPart.Presentation.SlideSize = new SlideSize() { Cx = max.Width, Cy = max.Height, Type = SlideSizeValues.Custom };

                // fill
                foreach (string p in list)
                {
                    Slide slide = initSlide(presentationPart, imgId.ToString());
                    fill(slide, max, p, imgId++);
                    Console.WriteLine(String.Format("Insert image[{0}] to slide[{1}]", p, imgId));
                    slide.Save();
                }

                presentationDocument.PresentationPart.Presentation.Save();
                Console.WriteLine("Save PPTX");
            }
        }

        private void fill(Slide slide, Size maxSize, string imagePath, int imgId)
        {
            P.Picture picture = new P.Picture();
            string embedId = string.Empty;
            string imageExt = getImageType(imagePath);
            Size imgSize = new ImageInfoUtils().getPPTSize(imagePath);
            embedId = "rId" + imgId.ToString();

            P.NonVisualPictureProperties nonVisualPictureProperties = new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Picture " + imgId },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks() { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties());
            picture.Append(nonVisualPictureProperties);

            UseLocalDpi useLocalDpi = new UseLocalDpi() { Val = false };
            useLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
            BlipExtension blipExtension = new BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };
            blipExtension.Append(useLocalDpi);
            BlipExtensionList blipExtensionList = new BlipExtensionList();
            blipExtensionList.Append(blipExtension);

            Stretch stretch = new Stretch();
            FillRectangle fillRectangle = new FillRectangle();
            stretch.Append(fillRectangle);

            P.ShapeProperties shapeProperties = new P.ShapeProperties()
            {
                Transform2D = new A.Transform2D()
                {
                    Offset = new A.Offset() { X = (maxSize.Width - imgSize.Width) / 2, Y = (maxSize.Height - imgSize.Height) / 2 },
                    Extents = new A.Extents() { Cx = imgSize.Width, Cy = imgSize.Height }
                }
            };
            shapeProperties.Append(new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle, AdjustValueList = new A.AdjustValueList() });

            Blip blip = new Blip() { Embed = embedId };
            blip.Append(blipExtensionList);
            P.BlipFill blipFill = new P.BlipFill() { Blip = blip };
            blipFill.Append(stretch);
            picture.Append(blipFill);

            picture.Append(shapeProperties);

            slide.CommonSlideData.ShapeTree.AppendChild(picture);

            ImagePart imagePart = slide.SlidePart.AddNewPart<ImagePart>(imageExt, embedId);
            FileStream fileStream = new FileStream(imagePath, FileMode.Open);
            imagePart.FeedData(fileStream);
            fileStream.Close();
        }

        public string convert(String path, List<PPTPage> pageList)
        {
            ImageInfoUtils utils = new ImageInfoUtils();
            List<String> imgList = new List<string>();
            pageList.ForEach(p =>
            {
                imgList.Add(p.Cover);
            });
            imgList = utils.filter(imgList);
            Size max = utils.listMaxPPTSize(imgList);
            String pptTempPath = System.Environment.CurrentDirectory + "\\..\\..\\Template\\template.pptx";

            if (0 == max.Width || 0 == max.Height)
                throw new Exception("image list are all invalid");
            if (!File.Exists(pptTempPath))
                throw new Exception("template file is not exists");

            return convert(path, pptTempPath, pageList, max);
        }

        private string convert(String path, String pptTempPath, List<PPTPage> pageList, Size size)
        {
            String filePath = String.Empty;
            while (String.IsNullOrEmpty(filePath) || System.IO.File.Exists(filePath))
            {
                string guid = Guid.NewGuid().ToString();
                filePath = System.IO.Path.Combine(path, guid + ".pptx");
            }

            try
            {
                if (!System.IO.Directory.Exists(path))
                    System.IO.Directory.CreateDirectory(path);
                File.Copy(pptTempPath, filePath);
                insert(filePath, pageList, size);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
                if (File.Exists(filePath))
                    File.Delete(filePath);
            }

            return filePath;
        }

        private void insert(string filePath, List<PPTPage> pageList, Size max)
        {
            int imgId = 915;
            using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
            {
                PresentationPart presentationPart = presentationDocument.PresentationPart;
                presentationPart.Presentation.SlideSize = new SlideSize() { Cx = max.Width, Cy = max.Height, Type = SlideSizeValues.Custom };

                int i = 0;
                int length = pageList.Count;
                foreach (PPTPage page in pageList)
                {
                    Slide slide = initSlide(presentationPart, imgId.ToString());
                    insertPage(slide, max, page, ref imgId);
                    slide.Save();

                    Console.WriteLine("IMAGE TO PPTX - {0}/{1}", ++i, length);
                }

                presentationDocument.PresentationPart.Presentation.Save();
            }
        }

        private void insertPage(Slide slide, Size maxSize, PPTPage page, ref int objId)
        {
            insertImage(slide, maxSize, page.Cover, ++objId);
            if (null != page.PageUrlList && page.PageUrlList.Count > 0)
            {
                foreach (PageUrl pageUrl in page.PageUrlList)
                {
                    insertLink(slide, pageUrl, ++objId);
                }
            }
        }

        private void insertLink(Slide slide, PageUrl pageUrl, int objId)
        {
            slide.SlidePart.AddHyperlinkRelationship(new System.Uri(pageUrl.url, System.UriKind.Absolute), true, "rId" + objId);

            P.Shape shape = new P.Shape();
            P.NonVisualShapeProperties nonVisualShapeProperties1 = new P.NonVisualShapeProperties()
            {
                NonVisualDrawingProperties = new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "矩形 1", HyperlinkOnClick = new A.HyperlinkOnClick() { Id = "rId" + objId } },
                NonVisualShapeDrawingProperties = new P.NonVisualShapeDrawingProperties(),
                ApplicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties()
            };

            P.ShapeProperties shapeProperties = new P.ShapeProperties()
            {
                Transform2D = new A.Transform2D()
                {
                    Offset = new A.Offset() { X = pageUrl.origin.Width * ImageInfoUtils.RATE, Y = pageUrl.origin.Height * ImageInfoUtils.RATE },
                    Extents = new A.Extents() { Cx = pageUrl.size.Width * ImageInfoUtils.RATE, Cy = pageUrl.size.Height * ImageInfoUtils.RATE }
                }
            };

            A.PresetGeometry presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle, AdjustValueList = new A.AdjustValueList() };
            A.NoFill noFill = new A.NoFill();
            A.Outline outline = new A.Outline();
            A.NoFill noFill2 = new A.NoFill();
            outline.Append(noFill2);

            shapeProperties.Append(presetGeometry);
            shapeProperties.Append(noFill);
            shapeProperties.Append(outline);

            P.ShapeStyle shapeStyle1 = new P.ShapeStyle();

            A.LineReference lineReference = new A.LineReference() { Index = (UInt32Value)2U };
            A.SchemeColor schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.Shade shade1 = new A.Shade() { Val = 50000 };
            schemeColor.Append(shade1);
            lineReference.Append(schemeColor);

            A.FillReference fillReference = new A.FillReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            fillReference.Append(schemeColor2);

            A.EffectReference effectReference = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            effectReference.Append(schemeColor3);

            A.FontReference fontReference = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };
            fontReference.Append(schemeColor4);

            shapeStyle1.Append(lineReference);
            shapeStyle1.Append(fillReference);
            shapeStyle1.Append(effectReference);
            shapeStyle1.Append(fontReference);

            P.TextBody textBody = new P.TextBody();
            A.BodyProperties bodyProperties = new A.BodyProperties() { RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle = new A.ListStyle();

            A.Paragraph paragraph = new A.Paragraph();
            A.ParagraphProperties paragraphProperties = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };
            A.EndParagraphRunProperties endParagraphRunProperties = new A.EndParagraphRunProperties() { Language = "zh-CN", AlternativeLanguage = "en-US" };

            paragraph.Append(paragraphProperties);
            paragraph.Append(endParagraphRunProperties);

            textBody.Append(bodyProperties);
            textBody.Append(listStyle);
            textBody.Append(paragraph);

            shape.Append(nonVisualShapeProperties1);
            shape.Append(shapeProperties);
            shape.Append(shapeStyle1);
            shape.Append(textBody);

            slide.CommonSlideData.ShapeTree.AppendChild(shape);
        }

        private void insertImage(Slide slide, Size maxSize, string imagePath, int imgId)
        {
            P.Picture picture = new P.Picture();
            string embedId = string.Empty;
            string imageExt = getImageType(imagePath);
            Size imgSize = new ImageInfoUtils().getPPTSize(imagePath);
            embedId = "rId" + imgId.ToString();

            P.NonVisualPictureProperties nonVisualPictureProperties = new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Picture " + imgId },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks() { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties());
            picture.Append(nonVisualPictureProperties);

            UseLocalDpi useLocalDpi = new UseLocalDpi() { Val = false };
            useLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
            BlipExtension blipExtension = new BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };
            blipExtension.Append(useLocalDpi);
            BlipExtensionList blipExtensionList = new BlipExtensionList();
            blipExtensionList.Append(blipExtension);

            Stretch stretch = new Stretch();
            FillRectangle fillRectangle = new FillRectangle();
            stretch.Append(fillRectangle);

            P.ShapeProperties shapeProperties = new P.ShapeProperties()
            {
                Transform2D = new A.Transform2D()
                {
                    Offset = new A.Offset() { X = (maxSize.Width - imgSize.Width) / 2, Y = (maxSize.Height - imgSize.Height) / 2 },
                    Extents = new A.Extents() { Cx = imgSize.Width, Cy = imgSize.Height }
                }
            };
            shapeProperties.Append(new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle, AdjustValueList = new A.AdjustValueList() });

            Blip blip = new Blip() { Embed = embedId };
            blip.Append(blipExtensionList);
            P.BlipFill blipFill = new P.BlipFill() { Blip = blip };
            blipFill.Append(stretch);
            picture.Append(blipFill);

            picture.Append(shapeProperties);

            slide.CommonSlideData.ShapeTree.AppendChild(picture);

            ImagePart imagePart = slide.SlidePart.AddNewPart<ImagePart>(imageExt, embedId);
            FileStream fileStream = new FileStream(imagePath, FileMode.Open);
            imagePart.FeedData(fileStream);
            fileStream.Close();
        }

        private Slide initSlide(PresentationPart presentationPart, string layoutName)
        {
            UInt32 slideId = 256U;
            var slideIdList = presentationPart.Presentation.SlideIdList;
            SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.First();
            SlideLayoutPart slideLayoutPart = slideMasterPart.SlideLayoutParts.First();

            if (slideIdList == null)
            {
                presentationPart.Presentation.SlideIdList = new SlideIdList();
                slideIdList = presentationPart.Presentation.SlideIdList;
            }

            slideId += Convert.ToUInt32(slideIdList.Count());
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
            slide.Save(slidePart);

            slidePart.AddPart<SlideLayoutPart>(slideLayoutPart);
            slidePart.Slide.CommonSlideData = (CommonSlideData)slideLayoutPart.SlideLayout.CommonSlideData.Clone();
            SlideId newSlideId = presentationPart.Presentation.SlideIdList.AppendChild<SlideId>(new SlideId());

            newSlideId.Id = slideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            return getSlideByRelationShipId(presentationPart, newSlideId.RelationshipId);
        }

        private Slide getSlideByRelationShipId(PresentationPart presentationPart, StringValue relationshipId)
        {
            SlidePart slidePart = presentationPart.GetPartById(relationshipId) as SlidePart;
            if (slidePart != null)
            {
                return slidePart.Slide;
            }
            return null;
        }

        private String getImageType(string filePath)
        {
            string imageExt = System.IO.Path.GetExtension(filePath);
            if (imageExt.Equals("jpg", StringComparison.OrdinalIgnoreCase))
                return "image/jpeg";

            return "image/png";
        }
    }
}
