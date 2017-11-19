using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace 画像を挿入する
{
    class Program
    {
        // 写真を貼る位置
        struct Position
        {
            public int X;
            public int Y;

            public Position( int X, int Y )
            {
                this.X = X;
                this.Y = Y;
            }
        }

        struct ImagePath
        {
            public string Path;
            public RotateFlipType Rot;
            
            public ImagePath( string path, RotateFlipType rot)
            {
                this.Path = path;
                this.Rot = rot;
            }
        }

        // 写真を貼る位置
        private static readonly Position[] POSITION = new Position[] { 
           new Position(108000,1321200)
            ,new Position(2348915,1321200)
            ,new Position(4590465,1321200)
            ,new Position(105460,5083575)
            ,new Position(2348915,5083575)
            ,new Position(4590465,5083575)
        };

        static void Main(string[] args)
        {
            //string fileName = @"C:\Users\t\Documents\07_OpenXML\template.pptx";
            //string fileNameCopy = @"C:\Users\t\Documents\07_OpenXML\template_1.pptx";
            string basePath = System.AppDomain.CurrentDomain.BaseDirectory;
            string fileName = basePath + "◆◆◆邸　事前写真.pptx";
            string fileNameCopy = basePath + "xxx邸　事前写真.pptx";
            System.IO.File.Copy( fileName, fileNameCopy, true );
            //string[] imagePath = {@"C:\Users\t\Documents\07_OpenXML\image.png",
            //                        @"C:\Users\t\Documents\07_OpenXML\image2.png"};
            List<ImagePath> imagePathList = null;
            imagePathList = ChkRotation(args);
            AddImage(fileNameCopy, imagePathList);
        }

        private static List<ImagePath> ChkRotation( string[] paths )
        {
            List<ImagePath> retList = new List<ImagePath>();

            foreach (string path in paths)
            {
                // 元の画像を開く
                using (var origin = new Bitmap(path))
                {
                    var rotation = RotateFlipType.RotateNoneFlipNone;

                    // 画像に付与されているEXIF情報を列挙する
                    foreach (var item in origin.PropertyItems)
                    {
                        if (item.Id != 0x0112)
                            continue;

                        // IFD0 0x0112; Orientationの値を調べる
                        switch (item.Value[0])
                        {
                            case 3:
                                // 時計回りに180度回転しているので、180度回転して戻す
                                rotation = RotateFlipType.Rotate180FlipNone;
                                break;
                            case 6:
                                // 時計回りに270度回転しているので、90度回転して戻す
                                rotation = RotateFlipType.Rotate90FlipNone;
                                break;
                            case 8:
                                // 時計回りに90度回転しているので、270度回転して戻す
                                rotation = RotateFlipType.Rotate270FlipNone;
                                break;
                        }
                    }
                    retList.Add(new ImagePath(path, rotation));
                }
            }
            return retList;
        }

        private static void AddImage(string file, List<ImagePath> image)
        {
            using (var presentation = PresentationDocument.Open(file, true))
            {
                var slideCount = presentation.PresentationPart.SlideParts.Count();
                PresentationPart presentationPart = presentation.PresentationPart;
                OpenXmlElementList slideIds = presentationPart.Presentation.SlideIdList.ChildElements;

                //var slideParts = presentation
                //    .PresentationPart
                //    .SlideParts.ToArray<SlidePart>();
 
                int cnt = 0;

                int j = 0;
                string relId = (slideIds[j] as SlideId).RelationshipId;
                SlidePart slidePart = (SlidePart)presentationPart.GetPartById(relId);

                foreach ( ImagePath imgPath in image )
                {
                    ImagePart part = slidePart
                            .AddImagePart(ImagePartType.Png);

                    using (var stream = File.OpenRead(imgPath.Path))
                    {
                        part.FeedData(stream);
                    }
                    var tree = slidePart
                                .Slide
                                .Descendants<DocumentFormat.OpenXml.Presentation.ShapeTree>()
                                .First();
                    var picture = new DocumentFormat.OpenXml.Presentation.Picture();


                    picture.NonVisualPictureProperties = new DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties();
                    picture.NonVisualPictureProperties.Append(new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties
                    {
                        Name = "My Shape",
                        Id = (UInt32)tree.ChildElements.Count - 1
                    });

                    var nonVisualPictureDrawingProperties = new DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties();
                    nonVisualPictureDrawingProperties.Append(new DocumentFormat.OpenXml.Drawing.PictureLocks()
                    {
                        NoChangeAspect = true
                    });
                    picture.NonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
                    picture.NonVisualPictureProperties.Append(new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties());

                    var blipFill = new DocumentFormat.OpenXml.Presentation.BlipFill();
                    var blip1 = new DocumentFormat.OpenXml.Drawing.Blip()
                    {
                        Embed = slidePart.GetIdOfPart(part)
                    };
                    var blipExtensionList1 = new DocumentFormat.OpenXml.Drawing.BlipExtensionList();
                    var blipExtension1 = new DocumentFormat.OpenXml.Drawing.BlipExtension()
                    {
                        Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                    };
                    var useLocalDpi1 = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi()
                    {
                        Val = false
                    };
                    useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                    blipExtension1.Append(useLocalDpi1);
                    blipExtensionList1.Append(blipExtension1);
                    blip1.Append(blipExtensionList1);
                    var stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
                    stretch.Append(new DocumentFormat.OpenXml.Drawing.FillRectangle());
                    blipFill.Append(blip1);
                    blipFill.Append(stretch);
                    picture.Append(blipFill);

                    picture.ShapeProperties = new DocumentFormat.OpenXml.Presentation.ShapeProperties();
                    picture.ShapeProperties.Transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D();

                    int rotation = 0;
                    switch (imgPath.Rot)
                    {
                        case RotateFlipType.RotateNoneFlipNone:
                            rotation = 0;
                            break;
                        case RotateFlipType.Rotate180FlipNone: // 時計回りに180度回転しているので、180度回転して戻す
                            rotation = 60000 * 180;
                            break;
                        case RotateFlipType.Rotate90FlipNone: // 時計回りに270度回転しているので、90度回転して戻す
                            rotation = 60000 * 90;
                            break;
                        case RotateFlipType.Rotate270FlipNone: // 時計回りに90度回転しているので、270度回転して戻す
                            rotation = 60000 * 270;
                            break;
                        default:
                            rotation = 0;
                            break;
                    }
                    picture.ShapeProperties.Transform2D.Rotation = rotation;
                    picture.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Offset
                    {
                        X = POSITION[cnt%6].X,
                        Y = POSITION[cnt%6].Y,
                    });

                    // 縦向き
                    if(imgPath.Rot == RotateFlipType.RotateNoneFlipNone || imgPath.Rot == RotateFlipType.Rotate180FlipNone){
                        picture.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Extents
                        {
                            Cx = 3600 * 6 * 100,
                            Cy = 3600 * 8 * 100,
                        });
                    }
                    else // 横向き
                    {
                        picture.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Extents
                        {
                            Cx = 3600 * 8 * 100,
                            Cy = 3600 * 6 * 100,
                        });
                    }


                    picture.ShapeProperties.Append(new DocumentFormat.OpenXml.Drawing.PresetGeometry
                    {
                        Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
                    });

                    tree.Append(picture);

                    if (cnt%6 == 5)
                    {
                        j++;
                        relId = (slideIds[j] as SlideId).RelationshipId;
                        slidePart = (SlidePart)presentationPart.GetPartById(relId);
                    }
                    cnt++;
                }



            }
        }
    }
}
