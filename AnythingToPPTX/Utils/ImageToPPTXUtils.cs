using System;
using System.Collections.Generic;
using System.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using System.IO;

namespace AnythingToPPTX.Utils
{
    class ImageToPPTXUtils
    {
        public string convert(String path, List<String> imgList)
        {
            ImageInfoUtils utils = new ImageInfoUtils();
            imgList = utils.filter(imgList);
            Size max = utils.listMaxSize(imgList);


            if (0 == max.Width || 0 == max.Height)
                throw new Exception("image list are all invalid");

            return convert(path, imgList, max);
        }

        private string convert(String path, List<String> imgList, Size size)
        {
            String filePath = String.Empty;
            while (String.IsNullOrEmpty(filePath) || System.IO.File.Exists(filePath))
            {
                string guid = Guid.NewGuid().ToString();
                filePath = System.IO.Path.Combine(path, guid + ".pptx");
            }

            try
            {
                using (PresentationDocument presentationDocument = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation))
                {
                    PresentationPart presentationPart = presentationDocument.AddPresentationPart();
                    presentationPart.Presentation = new Presentation();
                    insert(presentationPart, imgList);
                    presentationDocument.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.StackTrace);
                throw new Exception("fail to convert image to pptx");
            }

            return filePath;
        }

        private void insert(PresentationPart presentationPart, List<String> imgList)
        {
            int idx = 1;
            uint uid = UInt32.MaxValue;
            var slideParts = presentationPart.SlideParts;

            string slideMasterRid = "rId" + idx;
            SlideMasterIdList slideMasterIdList = new SlideMasterIdList(new SlideMasterId() { Id = uid, RelationshipId = slideMasterRid });
            SlideIdList slideIdList = new SlideIdList();
            SlideSize slideSize = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
            NotesSize notesSize = new NotesSize() { Cx = 6858000, Cy = 9144000 };
            DefaultTextStyle defaultTextStyle = new DefaultTextStyle();
            presentationPart.Presentation.Append(slideMasterIdList, slideIdList, slideSize, notesSize, defaultTextStyle);

            SlideLayoutPart layoutPart = null;
            SlideMasterPart masterPart = null;
            ThemePart themePart = null;

            foreach (string imgPath in imgList)
            {
                String imgIdx = "rId" + (900 + idx);
                String slideIdx = "rId" + (idx + 1);
                String themeIdx = "rId" + (idx + 4);

                var slidePart = CreateSlidePart(presentationPart, slideIdx, uid = uid - 10);
                if (null == layoutPart)
                {
                    layoutPart = CreateSlideLayoutPart(slidePart, slideMasterRid, uid = uid - 10);
                    masterPart = CreateSlideMasterPart(layoutPart, slideMasterRid, uid = uid - 10);
                    themePart = CreateTheme(masterPart, themeIdx);
                    masterPart.AddPart(layoutPart, slideMasterRid);
                    presentationPart.AddPart(masterPart, slideMasterRid);
                    presentationPart.AddPart(themePart, themeIdx);
                }

                //insert(slidePart, imgPath, imgIdx, uid = uid - 10);
                idx += 5;
            }
            presentationPart.Presentation.Save();
        }

        private void insert(SlidePart slidePart, string imagePath, string imgIdx, uint uid)
        {
            P.Picture picture = new P.Picture();
            string embedId = imgIdx;
            P.NonVisualPictureProperties nonVisualPictureProperties = new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties() { Id = uid--, Name = "Picture 5" },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks() { NoChangeAspect = true }),
                new ApplicationNonVisualDrawingProperties());
            P.BlipFill blipFill = new P.BlipFill();

            BlipExtensionList blipExtensionList = new BlipExtensionList();
            BlipExtension blipExtension = new BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };
            UseLocalDpi useLocalDpi = new UseLocalDpi() { Val = false };
            useLocalDpi.AddNamespaceDeclaration("a14",
                "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension.Append(useLocalDpi);
            blipExtensionList.Append(blipExtension);
            Blip blip = new Blip() { Embed = embedId };
            blip.Append(blipExtensionList);

            Stretch stretch = new Stretch();
            FillRectangle fillRectangle = new FillRectangle();
            stretch.Append(fillRectangle);

            blipFill.Append(blip);
            blipFill.Append(stretch);

            // TODO calc the size
            A.Transform2D transform2D = new A.Transform2D();
            A.Offset offset = new A.Offset() { X = 457200L, Y = 1524000L };
            A.Extents extents = new A.Extents() { Cx = 8229600L, Cy = 5029200L };
            transform2D.Append(offset);
            transform2D.Append(extents);

            A.PresetGeometry presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList = new A.AdjustValueList();
            presetGeometry.Append(adjustValueList);

            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            shapeProperties.Append(transform2D);
            shapeProperties.Append(presetGeometry);

            picture.Append(nonVisualPictureProperties);
            picture.Append(blipFill);
            picture.Append(shapeProperties);
            slidePart.Slide.CommonSlideData.ShapeTree.AppendChild(picture);

            var ext = System.IO.Path.GetExtension(imagePath).Substring(1);
            ext = ext.Equals("png", StringComparison.OrdinalIgnoreCase) ? "image/png" : "image/jpeg";
            ImagePart imagePart = slidePart.AddNewPart<ImagePart>(ext, embedId);
            using (FileStream fileStream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(fileStream);
            }
        }

        private static SlidePart CreateSlidePart(PresentationPart presentationPart, string slideIdx, uint uid)
        {
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>(slideIdx);
            slidePart.Slide = new Slide(new CommonSlideData(
                            new P.NonVisualGroupShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                                new P.NonVisualGroupShapeDrawingProperties(),
                                new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new TransformGroup())));
            presentationPart.Presentation.SlideIdList.Append(new SlideId() { Id = uid--, RelationshipId = slideIdx });
            return slidePart;
        }

        private static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart, string layoutIdx, uint uid)
        {
            SlideLayoutPart layoutPart = slidePart.AddNewPart<SlideLayoutPart>(layoutIdx);
            SlideLayout slideLayout = new SlideLayout(new CommonSlideData(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new TransformGroup()))));
            layoutPart.SlideLayout = slideLayout;
            return layoutPart;
        }

        private static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart, string masterIdx, uint uid)
        {
            SlideMasterPart slideMasterPart = slideLayoutPart.AddNewPart<SlideMasterPart>(masterIdx);
            SlideMaster slideMaster = new SlideMaster(new CommonSlideData(new ShapeTree(new ShapeTree(
              new P.NonVisualGroupShapeProperties(
              new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
              new P.NonVisualGroupShapeDrawingProperties(),
              new ApplicationNonVisualDrawingProperties()),
              new GroupShapeProperties(new TransformGroup())))));
            slideMasterPart.SlideMaster = slideMaster;

            return slideMasterPart;
        }

        private static ThemePart CreateTheme(SlideMasterPart slideMasterPart, string themeIdx)
        {
            ThemePart themePart = slideMasterPart.AddNewPart<ThemePart>(themeIdx);
            A.Theme theme = new A.Theme() { Name = "Office Theme" };

            A.ThemeElements themeElements1 = new A.ThemeElements(
            new A.ColorScheme(
              new A.Dark1Color(new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
              new A.Light1Color(new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
              new A.Dark2Color(new A.RgbColorModelHex() { Val = "1F497D" }),
              new A.Light2Color(new A.RgbColorModelHex() { Val = "EEECE1" }),
              new A.Accent1Color(new A.RgbColorModelHex() { Val = "4F81BD" }),
              new A.Accent2Color(new A.RgbColorModelHex() { Val = "C0504D" }),
              new A.Accent3Color(new A.RgbColorModelHex() { Val = "9BBB59" }),
              new A.Accent4Color(new A.RgbColorModelHex() { Val = "8064A2" }),
              new A.Accent5Color(new A.RgbColorModelHex() { Val = "4BACC6" }),
              new A.Accent6Color(new A.RgbColorModelHex() { Val = "F79646" }),
              new A.Hyperlink(new A.RgbColorModelHex() { Val = "0000FF" }),
              new A.FollowedHyperlinkColor(new A.RgbColorModelHex() { Val = "800080" })) { Name = "Office" },
              new A.FontScheme(
              new A.MajorFont(
              new A.LatinFont() { Typeface = "Calibri" },
              new A.EastAsianFont() { Typeface = "" },
              new A.ComplexScriptFont() { Typeface = "" }),
              new A.MinorFont(
              new A.LatinFont() { Typeface = "Calibri" },
              new A.EastAsianFont() { Typeface = "" },
              new A.ComplexScriptFont() { Typeface = "" })) { Name = "Office" },
              new A.FormatScheme(
              new A.FillStyleList(
              new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor }),
              new A.GradientFill(
                new A.GradientStopList(
                new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 50000 },
                  new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 37000 },
                 new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 35000 },
                new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 15000 },
                 new A.SaturationModulation() { Val = 350000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }
                ),
                new A.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new A.NoFill(),
              new A.PatternFill(),
              new A.GroupFill()),
              new A.LineStyleList(
              new A.Outline(
                new A.SolidFill(
                new A.SchemeColor(
                  new A.Shade() { Val = 95000 },
                  new A.SaturationModulation() { Val = 105000 }) { Val = A.SchemeColorValues.PhColor }),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = A.LineCapValues.Flat,
                  CompoundLineType = A.CompoundLineValues.Single,
                  Alignment = A.PenAlignmentValues.Center
              },
              new A.Outline(
                new A.SolidFill(
                new A.SchemeColor(
                  new A.Shade() { Val = 95000 },
                  new A.SaturationModulation() { Val = 105000 }) { Val = A.SchemeColorValues.PhColor }),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = A.LineCapValues.Flat,
                  CompoundLineType = A.CompoundLineValues.Single,
                  Alignment = A.PenAlignmentValues.Center
              },
              new A.Outline(
                new A.SolidFill(
                new A.SchemeColor(
                  new A.Shade() { Val = 95000 },
                  new A.SaturationModulation() { Val = 105000 }) { Val = A.SchemeColorValues.PhColor }),
                new A.PresetDash() { Val = A.PresetLineDashValues.Solid })
              {
                  Width = 9525,
                  CapType = A.LineCapValues.Flat,
                  CompoundLineType = A.CompoundLineValues.Single,
                  Alignment = A.PenAlignmentValues.Center
              }),
              new A.EffectStyleList(
              new A.EffectStyle(
                new A.EffectList(
                new A.OuterShadow(
                  new A.RgbColorModelHex(
                  new A.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new A.EffectStyle(
                new A.EffectList(
                new A.OuterShadow(
                  new A.RgbColorModelHex(
                  new A.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
              new A.EffectStyle(
                new A.EffectList(
                new A.OuterShadow(
                  new A.RgbColorModelHex(
                  new A.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
              new A.BackgroundFillStyleList(
              new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor }),
              new A.GradientFill(
                new A.GradientStopList(
                new A.GradientStop(
                  new A.SchemeColor(new A.Tint() { Val = 50000 },
                    new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                new A.GradientStop(
                  new A.SchemeColor(new A.Tint() { Val = 50000 },
                    new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                new A.GradientStop(
                  new A.SchemeColor(new A.Tint() { Val = 50000 },
                    new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 }),
                new A.LinearGradientFill() { Angle = 16200000, Scaled = true }),
              new A.GradientFill(
                new A.GradientStopList(
                new A.GradientStop(
                  new A.SchemeColor(new A.Tint() { Val = 50000 },
                    new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
                new A.GradientStop(
                  new A.SchemeColor(new A.Tint() { Val = 50000 },
                    new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 }),
                new A.LinearGradientFill() { Angle = 16200000, Scaled = true }))) { Name = "Office" });

            // new instance
            theme.Append(new A.ThemeElements());
            theme.Append(new A.ObjectDefaults());
            theme.Append(new A.ExtraColorSchemeList());

            themePart.Theme = theme;
            return themePart;

        }
    }
}