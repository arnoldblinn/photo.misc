using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;


using System.Drawing;
using System.Web;
using System.Xml;

using Msn.PhotoMix.Passport;
using Msn.Framework;

namespace Msn.PhotoMix.SlideShow
{    
    public struct FeedItem
    {
        public string title;
        public string description;
        public DateTime pubDate;
        public string imageUrl;
        public int imageWidth;
        public int imageHeight;
    }

    public enum FeedType
    {
        Rss = 0,
        Facebook = 1,
        Space = 2,
        Flickr = 3,
        SmugMug = 4,
        WebPage = 5
    }

    public class Feed
    {
        // Type of the feed
        private FeedType feedType = FeedType.Rss;

        // Data for the type of feed
        private string data = null;        
        
        // Title and logo corresponding to this feed
        private string title = null;
        private string logoUrl = null;

        // Items returned by the feed 
        private List<FeedItem> items = null;

        // Date the items were fetched, and date they go invalid
        private DateTime itemsDate;
        private DateTime itemsInvalidDate;

        // Flag indicating if the object is "dirty" (e.g. it fetched data since we last loaded)
        private bool isDirty = false;

        // Length of time that a fetch of an RSS feed is considered valid
        static int defaultFetchItemsTTL = Convert.ToInt32(Config.GetSetting("DefaultFetchItemsTTL"));

        public Feed(FeedType feedType, string data)
        {
            this.feedType = feedType;
            this.data = data;
        }
        
        public string Data
        {
            get { return this.data; }            
        }
        
        public string Title
        {
            get { return this.title; }
        }

        public string LogoUrl
        {
            get { return this.logoUrl; }
        }

        public DateTime ItemsInvalidDate
        {
            get { return this.itemsInvalidDate; }
        }

        public void AddUrlChannelItems(Channel channel, List<ListItem> listItems, DateTime dateContext, bool bypassCaches)
        {
            List<FeedItem> feedItems = this.GetFeedItems(dateContext, bypassCaches);
            this.Save();

            foreach (FeedItem feedItem in feedItems)
            {
                UrlChannelItem urlChannelItem = new UrlChannelItem(feedItem.pubDate, feedItem.imageUrl, feedItem.imageWidth, feedItem.imageHeight);
                urlChannelItem.Channel = channel;
                urlChannelItem.ExpDate = this.ItemsInvalidDate;
                listItems.Add((ChannelItem)urlChannelItem);
            }
        }

        public void AddUrlChannelItems(Channel channel, List<ListItem> listItems, DateTime dateContext, bool bypassCaches, int width, int height)
        {
            List<FeedItem> feedItems = this.GetFeedItems(dateContext, bypassCaches);
            this.Save();

            foreach (FeedItem feedItem in feedItems)
            {
                if (feedItem.imageWidth >= width && feedItem.imageHeight >= height)
                {
                    UrlChannelItem urlChannelItem = new UrlChannelItem(feedItem.pubDate, feedItem.imageUrl, feedItem.imageWidth, feedItem.imageHeight);
                    urlChannelItem.Channel = channel;
                    urlChannelItem.ExpDate = this.ItemsInvalidDate;
                    listItems.Add((ChannelItem)urlChannelItem);
                }
            }
        }

        private void AddRssFeed(List<FeedItem> items, string rssFeedUrl, RssChannelPriority priority, DateTime dateContext)
        {            
            try
            {           
                // Load the xml document
                XmlDocumentEx xmlDocument = new XmlDocumentEx();
                xmlDocument.Load(rssFeedUrl);
                xmlDocument.LoadNamespaces();

                // Determine if an atom or rss feed
                bool atom = false;
                XmlNode xmlNode = xmlDocument.SelectSingleNode("rss");
                if (xmlNode == null)
                {
                    xmlNode = xmlDocument.SelectSingleNode("dfltns:feed", xmlDocument.NamespaceManager);
                    if (xmlNode == null)
                        return;

                    atom = true;
                }

                if (atom)
                {
                    this.title = xmlDocument.SelectSingleNode("dfltns:feed/dfltns:title", xmlDocument.NamespaceManager).InnerText;

                    try
                    {
                        this.logoUrl = xmlDocument.SelectSingleNode("dfltns:feed/dfltnslogo", xmlDocument.NamespaceManager).InnerText;
                    }
                    catch { }

                    AddRssItems(xmlDocument, items, dateContext, atom, priority);
                }
                else
                {
                    this.title = xmlDocument.SelectSingleNode("rss/channel/title").InnerText;

                    try
                    {
                        this.logoUrl = xmlDocument.SelectSingleNode("rss/channel/image/url").InnerText;
                    }
                    catch { }

                    AddRssItems(xmlDocument, items, dateContext, atom, priority);                    
                }
            }
            catch (Exception)
            {
            }
        }

        private void AddAlbumId(List<FeedItem> items, string spaceName, string albumId, DateTime dateContext)
        {
            string albumFeed = "http://" + spaceName + ".spaces.live.com/photos/" + albumId + "/feed.rss";

            AddRssFeed(items, albumFeed, RssChannelPriority.Image, dateContext);
        }

        private void AddFacebookFeed(List<FeedItem> items, string sessionKey, DateTime dateContext)
        {
            try
            {         
                // Get infinite session key
                ArrayList args = new ArrayList();
                args.Add("v=1.0");
                args.Add("api_key=" + MiscUtil.GetFacebookApiKey());
                args.Add("session_key=" + sessionKey);
                args.Add("method=photos.get");
                string subjectId = sessionKey.Split('-')[1];
                args.Add("subj_id=" + subjectId);
                XmlDocumentEx xmlDoc = MiscUtil.CallFacebook(args);
                xmlDoc.LoadNamespaces();

                foreach (XmlNode node in xmlDoc.SelectNodes("//dfltns:photo", xmlDoc.NamespaceManager))
                {
                    XmlNode source = node.SelectSingleNode("./dfltns:src_big", xmlDoc.NamespaceManager);
                    if (source == null)
                        source = node.SelectSingleNode("./dfltns:src", xmlDoc.NamespaceManager);
                    if (source == null)
                        source = node.SelectSingleNode("./dfltns:src_small", xmlDoc.NamespaceManager);
                    if (source != null)
                    {
                        string imageUrl = source.InnerText;
                        uint created = Convert.ToUInt32(node.SelectSingleNode("./dfltns:created", xmlDoc.NamespaceManager).InnerText);
                        DateTime pubDate = new System.DateTime(1970, 1, 1).AddSeconds(created);

                        int width = -1;
                        int height = -1;

                        FeedItem feedItem = new FeedItem();                        
                        feedItem.pubDate = pubDate;
                        feedItem.imageUrl = imageUrl;
                        feedItem.imageWidth = width;
                        feedItem.imageHeight = height;
                        items.Add(feedItem);                                    
                    }
                }                
            }
            catch (Exception)
            {
            }
        }

        public void AddWebPageItems(List<FeedItem> items, string url, DateTime dateContext)
        {
            List<string> imageSources = WebPageBitmap.LoadDocumentImages(url);
            if (imageSources != null)
            {
                foreach (string imageSource in imageSources)
                {
                    Bitmap bitmap = ImageUtil.LoadImageFromUrl(imageSource);
                    if (bitmap != null && bitmap.Height >= 640 && bitmap.Width >= 480)
                    {
                        FeedItem feedItem = new FeedItem();
                        feedItem.pubDate = dateContext;
                        feedItem.imageUrl = imageSource;
                        feedItem.imageWidth = bitmap.Width;
                        feedItem.imageHeight = bitmap.Height;
                        items.Add(feedItem);                                                            
                    }
                }
            }
        }

        public List<FeedItem> GetFeedItems(DateTime dateContext, bool bypassCaches)
        {
            if (this.items == null || 
                bypassCaches ||
                this.itemsInvalidDate > dateContext)                
            {
                List<FeedItem> items = new List<FeedItem>();

                if (this.feedType == FeedType.Rss)
                {
                    string[] fields = this.data.Split('|');

                    string rssFeedUrl = fields[0];
                    RssChannelPriority priority = (RssChannelPriority)Convert.ToInt32(fields[1]);

                    AddRssFeed(items, rssFeedUrl, priority, dateContext);


                }
                else if (this.feedType == FeedType.Facebook)
                {
                    string sessionKey = this.data;

                    AddFacebookFeed(items, sessionKey, dateContext);
                }
                else if (this.feedType == FeedType.WebPage)
                {
                    string url = this.data;                    

                    AddWebPageItems(items, url, dateContext);                    
                }
                else if (this.feedType == FeedType.SmugMug)
                {
                    string[] fields = this.data.Split('|');

                    string userName = fields[0];
                    SmugMugFeedType smugMugFeedType = (SmugMugFeedType)Convert.ToInt32(fields[1]);
                    string rssFeedUrl;

                    if (smugMugFeedType == SmugMugFeedType.Popular)
                        rssFeedUrl = "http://api.smugmug.com/hack/feed.mg?Type=nicknamePopular&Data=" + userName + "&format=rss";
                    else // default is most recent
                        rssFeedUrl = "http://api.smugmug.com/hack/feed.mg?Type=nicknameRecent&Data=" + userName + "&format=rss";

                    AddRssFeed(items, rssFeedUrl, RssChannelPriority.Image, dateContext);
                }
                else if (this.feedType == FeedType.Flickr)
                {
                    string userName = this.data;

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load("http://api.flickr.com/services/rest/?method=flickr.people.findByUsername&username=" + userName + "&api_key=9c870e813907c4a318a0941e70d2fa7a");

                    string rssFeedUrl = "http://api.flickr.com/services/feeds/photos_public.gne?id=" + xmlDoc.SelectSingleNode("rsp/user/@nsid").Value + "&lang=en-us&format=rss_200";

                    AddRssFeed(items, rssFeedUrl, RssChannelPriority.Image, dateContext);
                }
                else if (this.feedType == FeedType.Space)
                {
                    string[] fields = this.data.Split('|');

                    string spaceName = fields[0];
                    string albumId = fields[1];

                    if (!String.IsNullOrEmpty(albumId))
                    {
                        AddAlbumId(items, spaceName, albumId, dateContext);
                    }
                    else
                    {
                        Dictionary<string, string> albums = SpaceChannel.GetAlbums(spaceName);

                        foreach (string key in albums.Keys)
                        {
                            string albumId2 = albums[key];

                            AddAlbumId(items, spaceName, albumId2, dateContext);
                        }
                    }
                }

                this.isDirty = true;

                this.itemsDate = dateContext;
                this.itemsInvalidDate = dateContext.AddMinutes(defaultFetchItemsTTL);

                this.items = items;                                    
            }

            return this.items;

        }

        private static void ProcessItem(XmlNode node, List<FeedItem> items, DateTime dateContext, bool atom, RssChannelPriority priority)
        {
            XmlDocumentEx xmlDocument = (XmlDocumentEx)node.OwnerDocument;            

            // Get the item title
            string itemTitle = null;            
            if (atom)
                itemTitle = node.SelectSingleNode("dfltns:title", xmlDocument.NamespaceManager).InnerText;
            else
                itemTitle = node.SelectSingleNode("title").InnerText;            

            // Get the item description
            string itemDescription = null;
            if (atom)
            {
                try
                {
                    itemDescription = node.SelectSingleNode("dfltns:content[@type=\"html\"]", xmlDocument.NamespaceManager).InnerText;
                }
                catch { }

                if (itemDescription == null)
                {
                    try
                    {
                        itemDescription = node.SelectSingleNode("dfltns:subtitle", xmlDocument.NamespaceManager).InnerText;
                    }
                    catch { }
                }
            }
            else
            {
                if (xmlDocument.NamespaceManager.HasNamespace("content"))
                {
                    try
                    {
                        itemDescription = node.SelectSingleNode(".//content:encoded", xmlDocument.NamespaceManager).InnerText;
                    }
                    catch { }
                }

                if (itemDescription == null)
                {
                    try
                    {
                        itemDescription = node.SelectSingleNode(".//description").InnerText;
                    }
                    catch { }
                }
            }            

            // Get the item published date (if there is one)
            DateTime pubDate = dateContext;
            try
            {
                string pubDateString;
                if (atom)
                    pubDateString = node.SelectSingleNode("dfltns:published", xmlDocument.NamespaceManager).InnerText;
                else
                    pubDateString = node.SelectSingleNode("pubDate").InnerText;
                if (!String.IsNullOrEmpty(pubDateString))
                    pubDate = Convert.ToDateTime(pubDateString);
            }
            catch { }

            // Get the image width, height of the item
            XmlNodeList contentNodes = null;
            XmlNode content = null;
            string imageUrl = null;
            int width = -1, height = -1;
            if (atom)
            {
                content = node.SelectSingleNode(".//dfltns:link[@rel=\"enclosure\"][@type=\"image/jpeg\"]", xmlDocument.NamespaceManager);

                if (content != null)
                {
                    imageUrl = content.Attributes["href"].Value;
                }

            }
            else
            {
                // See if there is a media tag
                if (xmlDocument.NamespaceManager.HasNamespace("media"))
                {
                    contentNodes = node.SelectNodes(".//media:content[@type=\"image/jpeg\"]", xmlDocument.NamespaceManager);
                }



                // Try the enclosure variant
                if (contentNodes == null)
                    contentNodes = node.SelectNodes(".//enclosure[@type=\"image/jpeg\"]", xmlDocument.NamespaceManager);

                if (contentNodes != null)
                {                                        
                    int maxFileSize = -1;
                        
                    foreach (XmlNode contentNode in contentNodes)
                    {
                        int fileSize = 0;
                        try
                        {
                            fileSize = Convert.ToInt32(contentNode.Attributes["fileSize"].Value);
                        }
                        catch { }
                        if (fileSize > maxFileSize)
                        {
                            content = contentNode;
                        }
                    }
                
                    imageUrl = content.Attributes["url"].Value;
                    if (!String.IsNullOrEmpty(imageUrl))
                    {
                        try
                        {
                            string widthString = content.Attributes["width"].Value;
                            if (!String.IsNullOrEmpty(widthString))
                                width = Convert.ToInt32(widthString);
                        }
                        catch { }

                        try
                        {
                            string heightString = content.Attributes["height"].Value;
                            if (!String.IsNullOrEmpty(heightString))
                                height = Convert.ToInt32(heightString);
                        }
                        catch { }
                    }
                }
            }

            // Only save the data we need based on the priority
            if (priority == RssChannelPriority.Image && imageUrl != null)
            {
                itemTitle = null;
                itemDescription = null;
            }
            else if (priority == RssChannelPriority.Text && itemDescription != null)
            {
                imageUrl = null;
                width = -1;
                height = -1;
            }

            FeedItem feedItem = new FeedItem();
            feedItem.title = itemTitle;
            feedItem.description = itemDescription;
            feedItem.pubDate = pubDate;
            feedItem.imageUrl = imageUrl;
            feedItem.imageWidth = width;
            feedItem.imageHeight = height;
            items.Add(feedItem);            
        }

        private static void AddRssItems(XmlDocumentEx xmlDocument, List<FeedItem> items, DateTime dateContext, bool atom, RssChannelPriority priority)
        {
            XmlNodeList xmlNodes;
            if (atom)
                xmlNodes = xmlDocument.SelectNodes("dfltns:feed/dfltns:entry", xmlDocument.NamespaceManager);
            else
                xmlNodes = xmlDocument.SelectNodes("rss/channel/item");

            foreach (XmlNode node in xmlNodes)
            {
                ProcessItem(node, items, dateContext, atom, priority);
            }
        }        

        static private string ItemsToString(List<FeedItem> items)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<Items>");
            foreach (FeedItem feedItem in items)
            {
                sb.Append("<Item>");

                if (!String.IsNullOrEmpty(feedItem.title))
                    sb.Append("<Title>" + HttpUtility.HtmlEncode(feedItem.title) + "</Title>");
                if (!String.IsNullOrEmpty(feedItem.description))
                    sb.Append("<Desc>" + HttpUtility.HtmlEncode(feedItem.description) + "</Desc>");
                if (feedItem.pubDate != DateTime.MinValue)
                    sb.Append("<Pub>" + feedItem.pubDate.ToString() + "</Pub>");
                if (!String.IsNullOrEmpty(feedItem.imageUrl))
                    sb.Append("<Image>" + HttpUtility.HtmlEncode(feedItem.imageUrl) + "</Image>");
                if (feedItem.imageWidth != -1)
                    sb.Append("<ImageW>" + feedItem.imageWidth.ToString() + "</ImageW>");
                if (feedItem.imageHeight != -1)
                    sb.Append("<ImageH>" + feedItem.imageHeight.ToString() + "</ImageH>");

                sb.Append("</Item>");
            }
            sb.Append("</Items>");

            return sb.ToString();
        }

        private static List<FeedItem> ItemsFromString(string data)
        {
            List<FeedItem> items = new List<FeedItem>();

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(data);
            XmlNodeList itemNodes = xmlDocument.SelectNodes("Items/Item");

            foreach (XmlNode itemNode in itemNodes)
            {
                FeedItem feedItem = new FeedItem();

                try
                {
                    feedItem.title = itemNode.SelectSingleNode("Title").InnerText;
                }
                catch { }
                try
                {
                    feedItem.description = itemNode.SelectSingleNode("Desc").InnerText;
                }
                catch { }
                try
                {
                    feedItem.pubDate = Convert.ToDateTime(itemNode.SelectSingleNode("Pub").InnerText);
                }
                catch { feedItem.pubDate = DateTime.MinValue; }
                try
                {
                    feedItem.imageUrl = itemNode.SelectSingleNode("Image").InnerText;
                }
                catch { }
                try
                {
                    feedItem.imageWidth = Convert.ToInt32(itemNode.SelectSingleNode("ImageW").InnerText);
                }
                catch { }
                try
                {
                    feedItem.imageHeight = Convert.ToInt32(itemNode.SelectSingleNode("ImageH").InnerText);
                }
                catch { }

                items.Add(feedItem);
            }

            return items;
        }

        public static Feed LoadRssFeed(string rssUrl, RssChannelPriority priority)
        {
            return Load(FeedType.Rss, rssUrl + "|" + ((int)priority).ToString());
        }

        public static Feed LoadSmugMugFeed(string userName, SmugMugFeedType smugMugFeedType)
        {
            return Load(FeedType.SmugMug, userName + "|" + ((int)smugMugFeedType).ToString());
        }

        public static Feed LoadWebPage(string url)
        {
            return Load(FeedType.WebPage, url);
        }

        public static Feed LoadFlickrFeed(string userName)
        {
            return Load(FeedType.Flickr, userName);
        }

        public static Feed LoadSpaceFeed(string spaceName, string albumId)
        {
            return Load(FeedType.Space, spaceName + "|" + albumId);
        }

        public static Feed LoadFacebookFeed(string sessionKey)
        {
            return Load(FeedType.Facebook, sessionKey);
        }

        public static Feed Load(FeedType type, string data)
        {
            Feed feed = new Feed(type, data);

            string sql = "select Title, LogoUrl, ItemsData, ItemsDataDate, ItemsDataInvalidDate " +
                        "from Feeds " +
                        "where Type = @Type and Data = @Data and DataHash = @DataHash " +
                        "update Feeds " +
                        "set ReferencedDate = @ReferencedDate " +
                        "where Type = @Type and Data = @Data and DataHash = @DataHash ";

            using (PhotoMixQuery query = new PhotoMixQuery(sql, CommandType.Text))
            {
                query.Parameters.Add("@Type", SqlDbType.Int).Value = (int)type;
                query.Parameters.Add("@Data", SqlDbType.NVarChar).Value = data;
                query.Parameters.Add("@DataHash", SqlDbType.Int).Value = data.GetHashCode();                
                query.Parameters.Add("@ReferencedDate", SqlDbType.DateTime).Value = DateTime.Now;

                if (query.Reader.Read())
                {
                    feed.title = query.Reader.GetString(0);
                    feed.logoUrl = query.Reader.IsDBNull(1) ? null : query.Reader.GetString(1);
                    feed.itemsInvalidDate = query.Reader.IsDBNull(4) ? DateTime.MinValue : query.Reader.GetDateTime(4);
                    if (feed.itemsInvalidDate < DateTime.Now)
                    {
                        feed.itemsDate = query.Reader.IsDBNull(3) ? DateTime.MinValue : query.Reader.GetDateTime(3);
                        feed.items = query.Reader.IsDBNull(2) ? null : ItemsFromString(query.Reader.GetString(2));
                    }
                }
            }

            return feed;
        }

        public void Save()
        {
            if (!this.isDirty)
                return;

            string sql = "if not exists (select Data from Feeds where Type = @Type and Data = @Data and DataHash = @DataHash) " +
                "    insert into Feeds (" +
                "       Type, Data, DataHash, Title, LogoUrl, ItemsData, ItemsDataDate, ItemsDataInvalidDate, ReferencedDate " +
                "    )" +
                "    values (" +
                "       @Type, @Data, @DataHash, @Title, @LogoUrl, @ItemsData, @ItemsDataDate, @ItemsDataInvalidDate, @ReferencedDate " +
                "    )" +
                "else" +
		        "    update Feeds" +
		        "    set " +
                "       Title = @Title, LogoUrl = @LogoUrl, ItemsData = @ItemsData, ItemsDataDate = @ItemsDataDate, ItemsDataInvalidDate = @ItemsDataInvalidDate, ReferencedDate = @ReferencedDate" +
                "    where Type = @Type and Data = @Data and DataHash = @DataHash ";

            using (PhotoMixQuery query = new PhotoMixQuery(sql, CommandType.Text))
            {
                query.Parameters.Add("@Type", SqlDbType.Int).Value = (int)this.feedType;
                query.Parameters.Add("@Data", SqlDbType.NVarChar).Value = this.data;
                query.Parameters.Add("@DataHash", SqlDbType.Int).Value = this.data.GetHashCode();
                if (this.title != null)
                    query.Parameters.Add("@Title", SqlDbType.NVarChar).Value = this.title;
                else
                    query.Parameters.Add("@Title", SqlDbType.NVarChar).Value = DBNull.Value;
                if (this.logoUrl != null)
                    query.Parameters.Add("@LogoUrl", SqlDbType.NVarChar).Value = this.logoUrl;
                else
                    query.Parameters.Add("@LogoUrl", SqlDbType.NVarChar).Value = DBNull.Value;
                query.Parameters.Add("@ItemsDataDate", SqlDbType.DateTime).Value = this.itemsDate;
                query.Parameters.Add("@ItemsDataInvalidDate", SqlDbType.DateTime).Value = this.itemsInvalidDate;
                query.Parameters.Add("@ItemsData", SqlDbType.Text).Value = ItemsToString(this.items);
                query.Parameters.Add("@ReferencedDate", SqlDbType.DateTime).Value = DateTime.Now;


                query.Execute();
            }
        }
    }
}
