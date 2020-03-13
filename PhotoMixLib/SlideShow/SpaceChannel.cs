using System;
using System.Collections.Generic;
using System.Text;

using System.Drawing;
using System.Web;
using System.Xml;

using Msn.PhotoMix.Passport;

using Msn.Framework;

namespace Msn.PhotoMix.SlideShow
{
    public class SpaceChannel : Channel
    {
        private string spaceName = null;
        private string albumId = null;

        public SpaceChannel(Puid puid, Guid slideShowGuid, Guid channelGuid)
            : base(puid, slideShowGuid, channelGuid, ChannelType.Space)
        {
        }

        public SpaceChannel(Puid puid, Guid slideShowGuid)
            : base(puid, slideShowGuid, ChannelType.Space)
        {
        }   

        public string SpaceName
        {
            get { return this.spaceName; }
            set 
            {
                if (this.spaceName != value)
                {
                    this.spaceName = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }        

        public string AlbumId
        {
            get { return this.albumId; }
            set
            {
                if (this.albumId != value)
                {
                    this.albumId = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }


        public override void SaveDataToString(StringBuilder sb)
        {
            base.SaveDataToString(sb);

            if (!String.IsNullOrEmpty(this.spaceName))
                sb.Append("<Name>" + HttpUtility.HtmlEncode(this.spaceName) + "</Name>");            
            if (!String.IsNullOrEmpty(this.albumId))
                sb.Append("<AlbumId>" + HttpUtility.HtmlEncode(this.albumId) + "</AlbumId>");

        }

        public override void LoadDataFromXmlNode(XmlNode node)
        {
            base.LoadDataFromXmlNode(node);

            try { this.spaceName = node.SelectSingleNode("Name").InnerText; }
            catch { }            
            try { this.albumId = node.SelectSingleNode("AlbumId").InnerText; }
            catch { }
        }

        public override void LoadDataFromQueryString(HttpRequest request)
        {
            base.LoadDataFromQueryString(request);

            if (!String.IsNullOrEmpty(request.QueryString["Name"]))
                this.spaceName = request.QueryString["Name"];
            if (!String.IsNullOrEmpty(request.QueryString["AlbumId"]))
                this.albumId = request.QueryString["AlbumId"];

        }

        static public Dictionary<string, string> GetAlbums(string spaceName)
        {
            Dictionary<string, string> albums = new Dictionary<string, string>();

            try
            {
                // Get the space rss document
                XmlDocumentEx rssXmlDocument = new XmlDocumentEx();
                rssXmlDocument.Load("http://" + spaceName + ".spaces.live.com/photos/feed.rss");
                rssXmlDocument.LoadNamespaces();

                // Get the album xml documents that are part of this list
                List<XmlDocumentEx> albumXmlDocuments = new List<XmlDocumentEx>();
                XmlNodeList xmlNodes = rssXmlDocument.SelectNodes("rss/channel/item");

                foreach (XmlNode xmlNode in xmlNodes)
                {
                    // See if this is an album
                    XmlNode type = xmlNode.SelectSingleNode("live:type", rssXmlDocument.NamespaceManager);
                    if (type.InnerText == "photoalbum")
                    {
                        string title = xmlNode.SelectSingleNode("title").InnerText;
                        title = title.Substring(title.IndexOf(": ") + 2);
                        string albumGuid = xmlNode.SelectSingleNode("guid").InnerText;

                        albums.Add(title, albumGuid);
                    }
                }
            }
            catch (Exception)
            {
            }

            return albums;
        }                

        public override List<ListItem> InitItems(List<Guid> guidContext, DateTime dateContext, bool bypassCaches, System.Collections.Hashtable compileState)
        {
            List<ListItem> items = new List<ListItem>();

            Feed feed = Feed.LoadSpaceFeed(this.spaceName, this.albumId);

            feed.AddUrlChannelItems(this, items, dateContext, bypassCaches);
            
            return items;                    
        }
    }
}
