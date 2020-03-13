using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Xml;

using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

using Msn.Framework;

namespace Msn.PhotoMix.SlideShow
{
    public class CompiledTextFeedItem
    {
        // Hash and unique id of the feed that this item is in
        private int compiledTextFeedHash = 0;
        private Guid compiledTextFeedGuid;

        // Title and description of this item
        private string title = null;
        private string description = null;

        // Hash of the item (title + description)
        private int compiledTextFeedItemHash = 0;
        
        // Date the item was referenced in a compile
        private DateTime compiledDate;

        // Date the image was generated
        private DateTime imageGeneratedDate;

        // Time to live for text feeds
        static private int textFeedImageTTL = Convert.ToInt32(Config.GetSetting("TextFeedImageTTL"));

        public string Title
        {
            get { return this.title; }
        }

        public string Description
        {
            get { return this.description; }
        }

        public int CompiledTextFeedItemHash
        {
            get { return this.compiledTextFeedItemHash; }
        }

        //
        // Load
        //
        // Loads the object from the database, and if it doesn't exist, creates it
        static public CompiledTextFeedItem Load(int compiledTextFeedHash, Guid compiledTextFeedGuid, string title, string description, int compiledTextFeedItemHash)
        {
            CompiledTextFeedItem compiledTextFeedItem = new CompiledTextFeedItem();
            compiledTextFeedItem.compiledTextFeedHash = compiledTextFeedHash;
            compiledTextFeedItem.compiledTextFeedGuid = compiledTextFeedGuid;
            if ((title != null) || (description != null))
                compiledTextFeedItem.compiledTextFeedItemHash = ((string)(title + description)).GetHashCode();
            else
                compiledTextFeedItem.compiledTextFeedItemHash = compiledTextFeedItemHash;
            compiledTextFeedItem.compiledDate = DateTime.Now;
            compiledTextFeedItem.imageGeneratedDate = DateTime.Now;

            using (PhotoMixQuery query = new PhotoMixQuery("SelectCompiledTextFeedItem"))
            {
                query.Parameters.Add("@CompiledTextFeedHash", SqlDbType.Int).Value = compiledTextFeedItem.compiledTextFeedHash;
                query.Parameters.Add("@CompiledTextFeedGuid", SqlDbType.UniqueIdentifier).Value = compiledTextFeedItem.compiledTextFeedGuid;
                query.Parameters.Add("@CompiledTextFeedItemHash", SqlDbType.Int).Value = compiledTextFeedItem.compiledTextFeedItemHash;
                if ((title != null) || (description != null))
                {
                    query.Parameters.Add("@Title", SqlDbType.VarChar).Value = String.IsNullOrEmpty(title) ? (Object)DBNull.Value : (Object)title;
                    query.Parameters.Add("@Description", SqlDbType.Text).Value = String.IsNullOrEmpty(description) ? (Object)DBNull.Value : (Object)description;
                    query.Parameters.Add("@CompiledDate", SqlDbType.DateTime).Value = compiledTextFeedItem.compiledDate;
                }
                else
                    query.Parameters.Add("@ImageGeneratedDate", SqlDbType.DateTime).Value = compiledTextFeedItem.imageGeneratedDate;

                if (query.Reader.Read())
                {
                    compiledTextFeedItem.title = query.Reader.IsDBNull(0) ? null : query.Reader.GetString(0);
                    compiledTextFeedItem.description = query.Reader.IsDBNull(1) ? null : query.Reader.GetString(1);
                    compiledTextFeedItem.compiledDate = query.Reader.IsDBNull(2) ? DateTime.MinValue : query.Reader.GetDateTime(2);
                    compiledTextFeedItem.imageGeneratedDate = query.Reader.IsDBNull(3) ? DateTime.MinValue : query.Reader.GetDateTime(3);
                }
            }

            return compiledTextFeedItem;
        }

        //
        // LoadForCompile
        //
        // Given a feed identifier (hash and guid) and data unique to an item in a text rss feed, will
        // either load or create data necessary to later generate an image representing this item
        //
        static public CompiledTextFeedItem LoadForCompile(int compiledTextFeedHash, Guid compiledTextFeedGuid, string title, string description)
        {
            return Load(compiledTextFeedHash, compiledTextFeedGuid, title, description, 0);
        }

        //
        // LoadForImage
        //
        // Given an item hash and feed identifier (hash and guid), will load/create an instance of 
        // this object in preparation for generation of an image.
        //
        static public CompiledTextFeedItem LoadForImage(int compiledTextFeedItemHash, int compiledTextFeedHash, Guid compiledTextFeedGuid)
        {
            return Load(compiledTextFeedHash, compiledTextFeedGuid, null, null, compiledTextFeedItemHash);
        }

        //
        // GenerateImage
        //
        // Will generate the bitmap representing this text item
        //
        public Bitmap GenerateImage(string fileName, SlideShowImageSize imageSize, string templateDirectory)
        {
            int width = SlideShow.slideShowImageWidths[(int)imageSize];
            int height = SlideShow.slideShowImageHeights[(int)imageSize];            

            CompiledTextFeed compiledTextFeed = CompiledTextFeed.GetCompiledTextFeed(compiledTextFeedHash, compiledTextFeedGuid);

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.AppendChild(xmlDoc.CreateElement("textfeeditem"));
            ImageUtil.AddXmlElement(xmlDoc.DocumentElement, "channeltitle", compiledTextFeed.Title);
            if (!String.IsNullOrEmpty(compiledTextFeed.LogoImageUrl))
                ImageUtil.AddXmlElement(xmlDoc.DocumentElement, "logoimg", compiledTextFeed.LogoImageUrl);
            if (!String.IsNullOrEmpty(this.title))
                ImageUtil.AddXmlElement(xmlDoc.DocumentElement, "itemtitle", this.title);
            if (!String.IsNullOrEmpty(this.description))
                ImageUtil.AddXmlElement(xmlDoc.DocumentElement, "itemdesc", this.description);
            if (compiledTextFeed.RenderAd)
                ImageUtil.AddXmlElement(xmlDoc.DocumentElement, "adurl", Config.GetSetting("AdUrl"));
            ImageUtil.AddXmlElement(xmlDoc.DocumentElement, "width", width.ToString());
            ImageUtil.AddXmlElement(xmlDoc.DocumentElement, "height", height.ToString());
            Bitmap bitmap = WebPageBitmap.LoadXsl(templateDirectory + "textfeeditem.xsl", xmlDoc, width, height);
            ImageUtil.SaveJpeg(fileName, bitmap, 100);

            return bitmap;
        }

        //
        // GetImageFileName
        //
        // Will get the image file name representing the item
        //
        public static string GetImageFileName(int compiledTextFeedHash, Guid compiledTextFeedGuid, int compiledTextFeedItemHash, SlideShowImageSize imageSize, string templateDirectory, bool bypassCaches)
        {
            string fileName = ImageUtil.GetCompiledImageDirectory("TextFeed") + compiledTextFeedGuid.ToString() + "_" + compiledTextFeedItemHash.ToString() + "_" + ((int)imageSize).ToString() + ".jpg";
            if (!bypassCaches && MiscUtil.TTLFileExists(fileName, CompiledTextFeedItem.textFeedImageTTL))
            {
                return fileName;
            }
            else
            {
                CompiledTextFeedItem compiledTextFeedItem = CompiledTextFeedItem.LoadForImage(compiledTextFeedItemHash, compiledTextFeedHash, compiledTextFeedGuid);

                Bitmap bitmap = compiledTextFeedItem.GenerateImage(fileName, imageSize, templateDirectory);

                return fileName;
            }
        }
    }

}
