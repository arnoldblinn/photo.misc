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
    public class CompiledWebPage
    {
        // Hash of the web page Url
        private int compiledWebPageHash;
        
        // Unique id for this particular feed
        private Guid compiledWebPageGuid;

        // Url of the web page
        private string url;

        // Date that the feed was referenced in a compile
        private DateTime compiledDate;

        // Date that the image was generated
        private DateTime imageGeneratedDate;

        // Date that the data for the image was last fetched
        private DateTime fetchDataDate;

        // Time to live for web pages
        static private int webPageFetchTTL = Convert.ToInt32(Config.GetSetting("WebPageFetchTTL"));
        static private int webPageImageTTL = Convert.ToInt32(Config.GetSetting("WebPageImageTTL"));

        public CompiledWebPage()
        {
        }

        public int CompiledWebPageHash
        {
            get { return this.compiledWebPageHash; }
        }

        public Guid CompiledWebPageGuid
        {
            get { return this.compiledWebPageGuid; }
        }

        public string Url
        {
            get { return this.url; }
        }

        public DateTime FetchDataDate
        {
            get { return this.fetchDataDate; }
        }

        //
        // FetchWebPageData
        //
        // Gets the web page image for the selected url,
        // then stores the page in the web page cache
        //
        static private void FetchWebPageData(string url, Guid compiledWebPageGuid)
        {
            try
            {
                string filename = ImageUtil.GetCompiledImageDirectory("WebPageFetchCache") + compiledWebPageGuid + ".jpg";

                Bitmap webPageBitmap = WebPageBitmap.Fetch(new Uri(url).AbsoluteUri, 800, 600);
                
                webPageBitmap.Save(filename, ImageFormat.Jpeg);
            }
            catch (Exception)
            {
            }
        }

        static private bool CheckFetchWebPageData(Guid compiledWebPageGuid)
        {
            string filename = ImageUtil.GetCompiledImageDirectory("WebPageFetchCache") + compiledWebPageGuid + ".jpg";
            if (MiscUtil.TTLFileExists(filename, webPageFetchTTL))
                return true;
            else
                return false;
        }

        //
        // Load
        //
        // Loads the object from the database, and if it doesn't exist, creates it
        //
        static private CompiledWebPage Load(string url, int compiledWebPageHash, Guid compiledWebPageGuid, bool forImage, DateTime dateContext, bool bypassCaches)
        {
            CompiledWebPage compiledWebPage = null;

            using (PhotoMixQuery query = new PhotoMixQuery("SelectCompiledWebPage"))
            {

                if (forImage)
                {
                    query.Parameters.Add("@ImageGeneratedDate", SqlDbType.DateTime).Value = dateContext;
                    query.Parameters.Add("@CompiledWebPageGuid", SqlDbType.UniqueIdentifier).Value = compiledWebPageGuid;
                    query.Parameters.Add("@CompiledWebPageHash", SqlDbType.Int).Value = compiledWebPageHash;
                }
                else
                {
                    query.Parameters.Add("@CompiledDate", SqlDbType.DateTime).Value = dateContext;
                    query.Parameters.Add("@Url", SqlDbType.NVarChar).Value = url;
                    query.Parameters.Add("@CompiledWebPageGuid", SqlDbType.UniqueIdentifier).Value = Guid.NewGuid();
                    query.Parameters.Add("@CompiledWebPageHash", SqlDbType.Int).Value = url.GetHashCode();
                }                    

                if (query.Reader.Read())
                {
                    compiledWebPage = new CompiledWebPage();   
                    compiledWebPage.compiledWebPageHash = query.Reader.GetInt32(0);
                    compiledWebPage.compiledWebPageGuid = query.Reader.GetGuid(1);
                    compiledWebPage.url = query.Reader.GetString(2);
                    compiledWebPage.compiledDate = query.Reader.GetDateTime(3);
                    compiledWebPage.imageGeneratedDate = query.Reader.IsDBNull(4) ? DateTime.MinValue : query.Reader.GetDateTime(4);
                    compiledWebPage.fetchDataDate = query.Reader.IsDBNull(5) ? DateTime.MinValue : query.Reader.GetDateTime(5);
                }
            }

            if (compiledWebPage != null &&
                (bypassCaches || compiledWebPage.fetchDataDate.AddMinutes(CompiledWebPage.webPageFetchTTL) < dateContext) || !CheckFetchWebPageData(compiledWebPage.compiledWebPageGuid)
                )
            {
                FetchWebPageData(compiledWebPage.url, compiledWebPage.compiledWebPageGuid);
                string sql = "update CompiledWebPages " +
                      "set FetchDataDate = @FetchDataDate " +
                      "where CompiledWebPageGuid = @CompiledWebPageGuid";
                using (PhotoMixQuery query2 = new PhotoMixQuery(sql, CommandType.Text))
                {
                    query2.Parameters.Add("@CompiledWebPageGuid", SqlDbType.UniqueIdentifier).Value = compiledWebPage.compiledWebPageGuid;
                    query2.Parameters.Add("@FetchDataDate", SqlDbType.DateTime).Value = dateContext;
                    query2.Execute();
                }

                compiledWebPage.fetchDataDate = dateContext;
            }


            return compiledWebPage;
        }

        //
        // LoadForCompile
        //
        // Given data unique to a text rss feed, will either create a new compiled reference
        // or update an existing compile reference
        //
        static public CompiledWebPage LoadForCompile(string url, DateTime dateContext)
        {
            return Load(url, 0, Guid.Empty, false, dateContext, false);            
        }

        //
        // LoadForImage
        //
        // Given a text feed hash (for webstore lookup) and unique id (unique to the feed),
        // looks up and loads an instance of this class for image generation.
        //
        static public CompiledWebPage LoadForImage(int compiledWebPageHash, Guid compiledWebPageGuid, DateTime dateContext, bool bypassCaches)
        {
            return Load(null, compiledWebPageHash, compiledWebPageGuid, true, dateContext, bypassCaches);            
        }

        //
        // GetImageFileName
        //
        // Will get the image file name representing the the web page, and will generate the image if necessary
        //
        public static string GetImageFileName(int compiledWebPageHash, SlideShowImageSize imageSize, Guid compiledWebPageGuid, bool bypassCaches)
        {
            string fileName = ImageUtil.GetCompiledImageDirectory("WebPage") + compiledWebPageGuid.ToString() + ".jpg";
            if (!bypassCaches && MiscUtil.TTLFileExists(fileName, CompiledWebPage.webPageImageTTL))
            {
                return fileName;
            }
            else
            {
                CompiledWebPage compiledWebPage = CompiledWebPage.LoadForImage(compiledWebPageHash, compiledWebPageGuid, DateTime.Now, bypassCaches);

                Bitmap bitmap = compiledWebPage.GenerateImage(fileName, imageSize);

                return fileName;
            }
        }

        //
        // GenerateImage
        //
        // This will generate an image suitable for rendering channel level "header" type data
        //
        private Bitmap GenerateImage(string fileName, SlideShowImageSize imageSize)
        {            
            Bitmap webPageBitmap = new Bitmap(ImageUtil.GetCompiledImageDirectory("WebPageFetchCache") + this.compiledWebPageGuid + ".jpg");

            webPageBitmap.Save(fileName, ImageFormat.Jpeg);

            return webPageBitmap;
        }
    }
}
