using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Drawing;
using System.Xml;

using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using Msn.Framework;

namespace Msn.PhotoMix.SlideShow
{       
    public class CompiledTextFeed
    {
        // Url to the Rss Feed
        private string textFeedUrl;        

        // Url to the channel image in the feed
        private string logoImageUrl;

        // Title of the feed
        private string title;

        // Flag to render an ad
        private bool renderAd = false;

        // Hash of the feed (hash of the url + logo + title)
        private int compiledTextFeedHash;

        // Date that the feed was referenced in a compile
        private DateTime compiledDate;

        // Date that the image was generated
        private DateTime imageGeneratedDate;

        // Unique id for this particular feed
        private Guid compiledTextFeedGuid;

        public CompiledTextFeed()
        {            
        }

        public string TextFeedUrl
        {
            get { return this.textFeedUrl; }
        }

        public string LogoImageUrl
        {
            get { return this.logoImageUrl; }
        }

        public string Title
        {
            get { return this.title; }
        }

        public bool RenderAd
        {
            get { return this.renderAd; }
        }

        public int CompiledTextFeedHash
        {
            get { return this.compiledTextFeedHash; }
        }

        public Guid CompiledTextFeedGuid
        {
            get { return this.compiledTextFeedGuid; }
        }

        //
        // Load
        //
        // Loads the object from the database, and if it doesn't exist, creates it
        static public CompiledTextFeed Load(string textFeedUrl, string logoImageUrl, string title, bool renderAd, int compiledTextFeedHash, Guid compiledTextFeedGuid)
        {
            CompiledTextFeed compiledTextFeed = new CompiledTextFeed();
            compiledTextFeed.textFeedUrl = textFeedUrl;
            if (textFeedUrl != null)
            {
                compiledTextFeed.compiledTextFeedHash = ((string)(textFeedUrl.ToLower() + renderAd.ToString())).GetHashCode();
                compiledTextFeed.compiledTextFeedGuid = Guid.NewGuid();
            }
            else
            {
                compiledTextFeed.compiledTextFeedHash = compiledTextFeedHash;
                compiledTextFeed.compiledTextFeedGuid = compiledTextFeedGuid;
            }
            compiledTextFeed.compiledDate = DateTime.Now;
            compiledTextFeed.imageGeneratedDate = DateTime.Now;

            using (PhotoMixQuery query = new PhotoMixQuery("SelectCompiledTextFeed"))
            {
                query.Parameters.Add("@CompiledTextFeedHash", SqlDbType.Int).Value = compiledTextFeed.compiledTextFeedHash;
                query.Parameters.Add("@CompiledTextFeedGuid", SqlDbType.UniqueIdentifier).Value = compiledTextFeed.compiledTextFeedGuid;
                if (textFeedUrl != null)
                {
                    query.Parameters.Add("@TextFeedUrl", SqlDbType.VarChar).Value = textFeedUrl;
                    query.Parameters.Add("@LogoImageUrl", SqlDbType.VarChar).Value = String.IsNullOrEmpty(logoImageUrl) ? (Object)DBNull.Value : (Object)logoImageUrl;
                    if (renderAd)
                        query.Parameters.Add("@Data", SqlDbType.Text).Value = "<RAD>true</RAD>";
                    query.Parameters.Add("@Title", SqlDbType.VarChar).Value = String.IsNullOrEmpty(title) ? (Object)DBNull.Value : (Object)title;
                    query.Parameters.Add("@CompiledDate", SqlDbType.DateTime).Value = compiledTextFeed.compiledDate;
                }
                else
                    query.Parameters.Add("@ImageGeneratedDate", SqlDbType.DateTime).Value = DateTime.Now;

                if (query.Reader.Read())
                {
                    compiledTextFeed.logoImageUrl = query.Reader.IsDBNull(0) ? null : query.Reader.GetString(0);
                    compiledTextFeed.title = query.Reader.IsDBNull(1) ? null : query.Reader.GetString(1);
                    if (!query.Reader.IsDBNull(2))
                    {
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.LoadXml(query.Reader.GetString(2));
                        try { compiledTextFeed.renderAd = FormUtil.GetBoolean(xmlDoc.SelectSingleNode("RAD").InnerText); }
                        catch { }
                    }
                    compiledTextFeed.compiledDate = query.Reader.IsDBNull(3) ? DateTime.MinValue : query.Reader.GetDateTime(3);
                    compiledTextFeed.imageGeneratedDate = query.Reader.IsDBNull(4) ? DateTime.MinValue : query.Reader.GetDateTime(4);
                    compiledTextFeed.compiledTextFeedGuid = query.Reader.GetGuid(5);
                }
            }

            return compiledTextFeed;
        }
        

        //
        // LoadForCompile
        //
        // Given data unique to a text rss feed, will either create a new compiled reference
        // or update an existing compile reference
        //
        static public CompiledTextFeed LoadForCompile(string textFeedUrl, string logoImageUrl, string title, bool renderAd)
        {
            return Load(textFeedUrl, logoImageUrl, title, renderAd, 0, Guid.Empty);
        }

        //
        // LoadForImage
        //
        // Given a text feed hash (for webstore lookup) and unique id (unique to the feed),
        // looks up and loads an instance of this class for image generation.
        //
        static public CompiledTextFeed LoadForImage(int compiledTextFeedHash, Guid compiledTextFeedGuid)
        {
            return Load(null, null, null, false, compiledTextFeedHash, compiledTextFeedGuid);
        }

        public static CompiledTextFeed GetCompiledTextFeed(int compiledTextFeedHash, Guid compiledTextFeedGuid)
        {
            CompiledTextFeed compiledTextFeed = CompiledTextFeed.LoadForImage(compiledTextFeedHash, compiledTextFeedGuid);

            return compiledTextFeed;
        }
        
    }
}
