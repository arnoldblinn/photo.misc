using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using System.Web;
using System.Net;
using System.IO;
using System.Xml;
using Msn.Framework;

namespace Msn.PhotoMix
{
    public class ImageUtil
    {
        public static Size GetFinalSize(Size targetSize, Size currentSize, bool grow, bool shrink)
        {
            Size finalSize = new Size(0, 0);

            // Check for growing and shrinking the image
            if (currentSize.Width < targetSize.Width && currentSize.Height < targetSize.Height)
            {
                if (!grow)
                {
                    finalSize.Width = currentSize.Width;
                    finalSize.Height = currentSize.Height;
                }
                else
                {
                    if (currentSize.Width / targetSize.Width > currentSize.Height / targetSize.Height)
                    {
                        finalSize.Width = targetSize.Width;
                        finalSize.Height = targetSize.Width * currentSize.Height / currentSize.Width;
                    }
                    else
                    {
                        finalSize.Height = targetSize.Height;
                        finalSize.Width = targetSize.Height * currentSize.Width / currentSize.Height;
                    }
                }
            }
            else if (shrink)
            {
                if (currentSize.Width / targetSize.Width > currentSize.Height / targetSize.Height)
                {
                    finalSize.Width = targetSize.Width;
                    finalSize.Height = targetSize.Width * currentSize.Height / currentSize.Width;
                }
                else
                {
                    finalSize.Height = targetSize.Height;
                    finalSize.Width = targetSize.Height * currentSize.Width / currentSize.Height;
                }
            }
            else
            {
                finalSize.Width = currentSize.Width;
                finalSize.Height = currentSize.Height;
            }

            return finalSize;
        }

        public static void SaveJpeg(string path, Bitmap img, long quality)
        {
            // Encoder parameter for image quality
            EncoderParameter qualityParam =
               new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, (long)quality);

            // Jpeg image codec
            ImageCodecInfo jpegCodec = getEncoderInfo("image/jpeg");

            if (jpegCodec == null)
                return;

            EncoderParameters encoderParams = new EncoderParameters(1);
            encoderParams.Param[0] = qualityParam;

            img.Save(path, jpegCodec, encoderParams);
        }

        private static ImageCodecInfo getEncoderInfo(string mimeType)
        {
            // Get image codecs for all image formats
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();

            // Find the correct image codec
            for (int i = 0; i < codecs.Length; i++)
                if (codecs[i].MimeType == mimeType)
                    return codecs[i];
            return null;
        }

        public static Bitmap LoadImageFromUrl(string url)
        {
            try
            {
                WebClient webClient = new WebClient();
                byte[] buffer = webClient.DownloadData(url);
                MemoryStream stream = new MemoryStream();
                stream.Write(buffer, 0, buffer.Length);
                Bitmap bitmap = new Bitmap(stream);

                return bitmap;
            }
            catch (Exception)
            {
            }

            return null;
        }

        public static void ImageToAd16x9(
            string imageName,
            string fileName)
        {
            Bitmap bitmap = new Bitmap(800, 450);
            Graphics graphics = Graphics.FromImage(bitmap);
            SolidBrush brush = new SolidBrush(Color.Black);

            Image image = Image.FromFile(imageName);

            Size targetSize = new Size(620, 400);

            Size finalSize = ImageUtil.GetFinalSize(targetSize, image.Size, false, true);
            Rectangle rect = new Rectangle(20, 25, (int)finalSize.Width, (int)finalSize.Height);

            // Draw the image            
            graphics.DrawImage(image, rect);

            // Draw a black box where the ad would be
            graphics.DrawRectangle(new Pen(brush), 20 + 620 + 20, 25, 120, 240);

            bitmap.Save(fileName);
        }

        public static void ImageToAd4x3(
            string imageName,
            string fileName)
        {
            Bitmap bitmap = new Bitmap(800, 600);
            Graphics graphics = Graphics.FromImage(bitmap);
            SolidBrush brush = new SolidBrush(Color.Black);

            Image image = Image.FromFile(imageName);

            Size targetSize = new Size(620, 400);
            Size finalSize = ImageUtil.GetFinalSize(targetSize, image.Size, false, true);
            Rectangle rect = new Rectangle(90, 36, (int)finalSize.Width, (int)finalSize.Height);

            // Draw the image
            graphics.DrawImage(image, rect);

            // Draw a black box where the ad would be
            graphics.DrawRectangle(new Pen(brush), 36, 36 + 37 + 400, 728, 90);

            bitmap.Save(fileName);
        }

        public static string GetCompiledImageDirectory(string subdirectory)
        {
            string path = Config.GetSetting("CompiledImageDirectory") + (subdirectory != null ? "\\" + subdirectory + "\\" : "");

            Directory.CreateDirectory(path);

            return path;
        }

        public static void AddXmlElement(XmlNode node, string elementName, string elementValue)
        {
            XmlNode child = node.OwnerDocument.CreateElement(elementName);
            child.InnerText = elementValue;
            node.AppendChild(child);
        }
    }
}
