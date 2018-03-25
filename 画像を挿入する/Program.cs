using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Drawing2 = DocumentFormat.OpenXml.Drawing;

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

        class Size
        {
            public int Width = 0;
            public int Height = 0;
            public static int _6cm = 3600 * 6 * 100;
            public static int _8cm = 3600 * 8 * 100;

        }

        struct ImagePath
        {
            public string Path;
            public RotateFlipType Rot;
            public Size Size;
            
            public ImagePath( string path, RotateFlipType rot, Size size)
            {
                this.Path = path;
                this.Rot = rot;
                this.Size = size;
            }
        }

        // 写真を貼る位置
        private static readonly Position[] POSITION0 = new Position[] { 
           new Position(104775, 1322425)
            ,new Position(2359195,1322425)
            ,new Position(4613615,1322425)
            ,new Position(104775,5364048)
            ,new Position(2359195,5364048)
            ,new Position(4613615,5364048)
        };
        private static readonly Position[] POSITION90 = new Position[] { 
           new Position(-229594, 1673928)
            ,new Position(2011321,1673928)
            ,new Position(4252871,1673928)
            ,new Position(-229594,5364048)
            ,new Position(2011321,5364048)
            ,new Position(4252871,5364048)
        };

        static string GetSaveFileName( string[] args, int i_nameOrdate )
        {
            String retFileName = "xxx";
            if(i_nameOrdate == 2)
            {
                retFileName = "xxx";
            }
            else
            {
                DateTime dtNow = DateTime.Now;
                retFileName = dtNow.Year + "年xx月xx日";
            }

            if(args.Length == 0)
            {
                return retFileName;
            }

            foreach( string filePath in args)
            {
                string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
                char[] charSeparators = new char[] { '_', '＿' };
                string[] fileNameList = fileName.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries);
                if( fileNameList.Length == 1)
                {
                    continue;
                }
                else
                { 
                    if( fileNameList.Length > 1 && i_nameOrdate == 2 )
                    {
                        retFileName = fileNameList[i_nameOrdate-1];
                    }
                    else if( fileNameList.Length > 2 && i_nameOrdate == 3)
                    {
                        retFileName = fileNameList[i_nameOrdate - 1];
                    }
                    break;
                }
            }

            return retFileName;
        }

        static int IsExpirationDate()
        {
            // NTPサーバへの接続用UDP生成
            System.Net.Sockets.UdpClient objSck;
            System.Net.IPEndPoint ipAny =
                new System.Net.IPEndPoint(System.Net.IPAddress.Any, 0);
            objSck = new System.Net.Sockets.UdpClient(ipAny);

            // NTPサーバへのリクエスト送信
            Byte[] sdat = new Byte[48];
            sdat[0] = 0xB;
            Byte[] rdat = null;
            try
            {
                objSck.Send(sdat, sdat.GetLength(0), "time.windows.com", 123);

                // NTPサーバから日時データ受信
                rdat = objSck.Receive(ref ipAny);
            }
            catch( Exception e)
            {
                Console.WriteLine(e.ToString());
                MessageBox.Show("ネットワークエラー","エラー",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return 99;
            }
            finally
            {
                objSck.Close();
            }

            // 1900年1月1日からの経過時間(日時分秒)
            long lngAllS; // 1900年1月1日からの経過秒数
            long lngD;    // 日
            long lngH;    // 時
            long lngM;    // 分
            long lngS;    // 秒

            // 1900年1月1日からの経過秒数計算
            lngAllS = (long)(
                      rdat[40] * Math.Pow(2, (8 * 3)) +
                      rdat[41] * Math.Pow(2, (8 * 2)) +
                      rdat[42] * Math.Pow(2, (8 * 1)) +
                      rdat[43]);

            // 1900年1月1日からの経過(日時分秒)計算
            lngD = lngAllS / (24 * 60 * 60); // 日
            lngS = lngAllS % (24 * 60 * 60); // 残りの秒数
            lngH = lngS / (60 * 60);         // 時
            lngS = lngS % (60 * 60);         // 残りの秒数
            lngM = lngS / 60;                // 分
            lngS = lngS % 60;                // 秒

            // 現在の日時(DateTime)計算
            DateTime dtTime = new DateTime(1900, 1, 1);
            dtTime = dtTime.AddDays(lngD);
            dtTime = dtTime.AddHours(lngH);
            dtTime = dtTime.AddMinutes(lngM);
            dtTime = dtTime.AddSeconds(lngS);

            // グリニッジ標準時から日本時間への変更
            dtTime = dtTime.AddHours(9);

            // 現在の日時の比較
            return "20171231".CompareTo(dtTime.ToString("yyyyMMdd")); 
        }

        static void Main(string[] args)
        {
            if(args.Count() == 0)
            {
                MessageBox.Show("画像のパスを指定して実行してください。","実行エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //int isValid = 0;
            //isValid = IsExpirationDate();
            //if( isValid < 0)
            //{
            //    MessageBox.Show("有効期限が切れました。", "実行エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    return;
            //}else if( isValid == 99)
            //{
            //    return;
            //}

            string basePath = System.AppDomain.CurrentDomain.BaseDirectory;

            string templateFileName = "◆◆◆邸　着工前写真.pptx";
            string fileName = basePath + templateFileName;

            // ファイル名
            string replaceName = GetSaveFileName(args, 2);
            // 撮影日付
            string insertDate = GetSaveFileName(args, 3);

            // テンプレートから事前写真をCopyする
            string fileNameCopy = fileName.Replace("◆◆◆",replaceName);
            System.IO.File.Copy( fileName, fileNameCopy, true );

            // 写真の向き・回転
            List<ImagePath> imagePathList = null;
            imagePathList = ChkRotation(args);

            // 写真の添付
            AddImage(fileNameCopy, imagePathList);

            // お客様名と撮影日を追加する
            InsertNameAndDate(fileNameCopy, replaceName, insertDate);

        }

        private static void InsertNameAndDate( string pptxPath, string insertName, string insertDate)
        {
            using( PresentationDocument ppt = PresentationDocument.Open(pptxPath, true))
            {
                if(ppt == null)
                {
                    throw new ArgumentNullException("presentationDocument");
                }

                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;

                for( int index = 0; index < slideIds.Count(); index++)
                {
                    // スライドを取得する
                    string relId = (slideIds[index] as SlideId).RelationshipId;
                    SlidePart slide = (SlidePart)part.GetPartById(relId);

                    if(slide != null)
                    {
                        ShapeTree tree = slide.Slide.CommonSlideData.ShapeTree;

                        //１番目の<s:sp>を取得する
                        Shape shape = tree.GetFirstChild<Shape>();

                        if( shape != null)
                        {
                            TextBody textBody = shape.TextBody;
                            IEnumerable<Drawing2.Paragraph> paragraphs = textBody.Descendants<Drawing2.Paragraph>();

                            foreach( Drawing2.Paragraph paragraph in paragraphs)
                            {
                                foreach( var text in paragraph.Descendants<Drawing2.Text>())
                                {
                                    if(text.Text.Contains("様邸"))
                                    {
                                        text.Text = insertName + text.Text;
                                    }
                                    else if (text.Text.Contains("年月日"))
                                    {
                                        text.Text = text.Text.Replace("年月日",insertDate);
                                    }
                                }
                            }
                        }
                        slide.Slide.Save();
                    }
                }

            }

            return;
        }

        private static List<ImagePath> ChkRotation( string[] paths )
        {
            List<ImagePath> retList = new List<ImagePath>();

            foreach (string path in paths)
            {
                // 元の画像を開く
                using (Bitmap origin = new Bitmap(path))
                {
                    var height = origin.Size.Height;
                    var width = origin.Size.Width;

                    // 事前写真に貼り付ける画像のサイズ
                    Size size = new Size();
                    if (height > width)
                    {
                        // 縦長画像
                        size.Width = Size._6cm ;
                        size.Height = height * ( Size._6cm / width );
                    }
                    else
                    {
                        // 横長画像
                        size.Width = width * ( Size._6cm / height );
                        size.Height = Size._6cm;
                    }

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
                    retList.Add(new ImagePath(path, rotation, size));
                }
            }
            return retList;
        }

        private static void AddImage(string file, List<ImagePath> image)
        {
            using (var presentationDocument = PresentationDocument.Open(file, true))
            {
                var slideCount = presentationDocument.PresentationPart.SlideParts.Count();
                SlideIdList slideIdList = presentationDocument.PresentationPart.Presentation.SlideIdList;
                Presentation presentation = presentationDocument.PresentationPart.Presentation;
                PresentationPart presentationPart = presentationDocument.PresentationPart;
                OpenXmlElementList slideIds = presentationPart.Presentation.SlideIdList.ChildElements;

                int cnt = 0;    // 画像の添付数

                int j = 0;  // 画像添付スライド位置
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
                    nonVisualPictureDrawingProperties.Append(new Drawing2.PictureLocks()
                    {
                        NoChangeAspect = true
                    });
                    picture.NonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
                    picture.NonVisualPictureProperties.Append(new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties());

                    var blipFill = new DocumentFormat.OpenXml.Presentation.BlipFill();
                    var blip1 = new Drawing2.Blip()
                    {
                        Embed = slidePart.GetIdOfPart(part)
                    };
                    var blipExtensionList1 = new Drawing2.BlipExtensionList();
                    var blipExtension1 = new Drawing2.BlipExtension()
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
                    var stretch = new Drawing2.Stretch();
                    stretch.Append(new Drawing2.FillRectangle());
                    blipFill.Append(blip1);
                    blipFill.Append(stretch);
                    picture.Append(blipFill);

                    picture.ShapeProperties = new DocumentFormat.OpenXml.Presentation.ShapeProperties();
                    picture.ShapeProperties.Transform2D = new Drawing2.Transform2D();

                    int rotation = 0;
                    switch (imgPath.Rot)
                    {
                        case RotateFlipType.RotateNoneFlipNone:
                            rotation = 0;
                            picture.ShapeProperties.Transform2D.Append(new Drawing2.Offset
                            {
                                X = POSITION0[cnt%6].X,
                                Y = POSITION0[cnt%6].Y,
                            });
                            break;
                        case RotateFlipType.Rotate180FlipNone: // 時計回りに180度回転しているので、180度回転して戻す
                            rotation = 60000 * 180;
                            picture.ShapeProperties.Transform2D.Append(new Drawing2.Offset
                            {
                                X = POSITION0[cnt%6].X,
                                Y = POSITION0[cnt%6].Y,
                            });
                            break;
                        case RotateFlipType.Rotate90FlipNone: // 時計回りに270度回転しているので、90度回転して戻す
                            rotation = 60000 * 90;
                            picture.ShapeProperties.Transform2D.Append(new Drawing2.Offset
                            {
                                X = POSITION90[cnt%6].X,
                                Y = POSITION90[cnt%6].Y,
                            });
                            break;
                        case RotateFlipType.Rotate270FlipNone: // 時計回りに90度回転しているので、270度回転して戻す
                            rotation = 60000 * 270;
                            picture.ShapeProperties.Transform2D.Append(new Drawing2.Offset
                            {
                                X = POSITION90[cnt%6].X,
                                Y = POSITION90[cnt%6].Y,
                            });
                            break;
                        default:
                            rotation = 0;
                            picture.ShapeProperties.Transform2D.Append(new Drawing2.Offset
                            {
                                X = POSITION0[cnt%6].X,
                                Y = POSITION0[cnt%6].Y,
                            });
                            break;
                    }
                    picture.ShapeProperties.Transform2D.Rotation = rotation;

                    // 縦向き
                    picture.ShapeProperties.Transform2D.Append(new Drawing2.Extents
                    {
                        Cx = imgPath.Size.Width,
                        Cy = imgPath.Size.Height,
                    });

                    picture.ShapeProperties.Append(new Drawing2.PresetGeometry
                    {
                        Preset = Drawing2.ShapeTypeValues.Rectangle
                    });

                    tree.Append(picture);

                    if (cnt%6 == 5)
                    {
                        if( j < slideCount - 1)
                        {
                            j++;
                            relId = (slideIds[j] as SlideId).RelationshipId;
                            slidePart = (SlidePart)presentationPart.GetPartById(relId);
                        }
                        else
                        {
                            // 画像ループを抜ける
                            break;
                        }
                    }
                    cnt++;
                }

                // スライドを削除する
                for( int i = slideCount-1; i > j; i--)
                {
                    //Console.WriteLine(i);
                    SlideId slideId = slideIds[i] as SlideId;
                    string slideRelId = slideId.RelationshipId;
                    slideIdList.RemoveChild(slideId);

                    if( presentation.CustomShowList != null)
                    {
                        // Iterate through the list of custom shows.
                        foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
                        {
                            if (customShow.SlideList != null)
                            {
                                // Declare a link list of slide list entries.
                                LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                                foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                                {
                                    // Find the slide reference to remove from the custom show.
                                    if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                                    {
                                        slideListEntries.AddLast(slideListEntry);
                                    }
                                }

                                // Remove all references to the slide from the custom show.
                                foreach (SlideListEntry slideListEntry in slideListEntries)
                                {
                                    customShow.SlideList.RemoveChild(slideListEntry);
                                }
                            }
                        }
                    }
                    presentation.Save();

                    SlidePart slidePart2 = presentationPart.GetPartById(slideRelId) as SlidePart;

                    presentationPart.DeletePart(slidePart2);

                }

            }
        }
    }
}
