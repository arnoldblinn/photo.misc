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
    public enum RssChannelPriority
    {
        Text = 0,
        Image = 1,

        Max = 1
    }

    public class RssChannel : Channel
    {
        // Url that was fetched for this feed
        protected string rssFeedUrl = null;

        // Flag to tell how to handle images and text
        private RssChannelPriority rssChannelPriority = RssChannelPriority.Text;

        // Settings for text items within the channel
        private bool renderChannelImage = true;
        private bool renderChannelTitle = true;
        private bool renderItemTitles = true;
        private bool renderItemDescriptions = true;
        private bool renderAd = false;

        // Title and image data of the channel in the returned Rss feed
        protected string channelTitle = null;
        protected string channelImageUrl = null;        

        public RssChannel(Puid puid, Guid slideShowGuid, Guid channelGuid, ChannelType type)
            : base(puid, slideShowGuid, channelGuid, type)
        {
        }

        public RssChannel(Puid puid, Guid slideShowGuid, Guid channelGuid)
            : base(puid, slideShowGuid, channelGuid, ChannelType.Rss)
        {
        }

        public RssChannel(Puid puid, Guid slideShowGuid, ChannelType type)
            : base(puid, slideShowGuid, type)
        {
        }

        public RssChannel(Puid puid, Guid slideShowGuid)
            : base(puid, slideShowGuid, ChannelType.Rss)
        {
        }

        public override void SaveDataToString(StringBuilder sb)
        {
            base.SaveDataToString(sb);

            sb.Append("<Url>" + HttpUtility.HtmlEncode(this.rssFeedUrl) + "</Url>");
            sb.Append("<Pri>" + (int) this.rssChannelPriority + "</Pri>");            
            if (!this.renderChannelImage)
                sb.Append("<RCI>" + this.renderChannelImage.ToString() + "</RCI>");
            if (!this.renderChannelTitle)
                sb.Append("<RCT>" + this.renderChannelTitle.ToString() + "</RCT>");
            if (!this.renderItemTitles)
                sb.Append("<RIT>" + this.renderItemTitles.ToString() + "</RIT>");
            if (!this.renderItemDescriptions)
                sb.Append("<RID>" + this.renderItemDescriptions.ToString() + "</RID>");
            if (this.renderAd)
                sb.Append("<RAD>" + this.renderAd.ToString() + "</RAD>");

            // Data cached from the feed fetch
            if (!String.IsNullOrEmpty(this.channelTitle))
                sb.Append("<Title>" + HttpUtility.HtmlEncode(this.channelTitle) + "</Title>");
            if (!String.IsNullOrEmpty(this.channelImageUrl))
                sb.Append("<ImageUrl>" + HttpUtility.HtmlEncode(this.channelImageUrl) + "</ImageUrl>");            
        }

        public override void LoadDataFromXmlNode(XmlNode node)
        {
            base.LoadDataFromXmlNode(node);

            try { this.rssFeedUrl = node.SelectSingleNode("Url").InnerText; }
            catch { }
            try { this.rssChannelPriority = (RssChannelPriority)FormUtil.GetNumber(node.SelectSingleNode("Pri").InnerText, 0, (int)RssChannelPriority.Max); }
            catch { }
            try { this.renderChannelImage = FormUtil.GetBoolean(node.SelectSingleNode("RCI").InnerText); }
            catch { }
            try { this.renderChannelTitle = FormUtil.GetBoolean(node.SelectSingleNode("RCT").InnerText); }
            catch { }
            try { this.renderItemTitles = FormUtil.GetBoolean(node.SelectSingleNode("RIT").InnerText); }
            catch { }
            try { this.renderItemDescriptions = FormUtil.GetBoolean(node.SelectSingleNode("RID").InnerText); }
            catch { }
            try { this.renderAd = FormUtil.GetBoolean(node.SelectSingleNode("RAD").InnerText); }
            catch { }
            
            // Data cached from the feed fetch
            try { this.channelTitle = node.SelectSingleNode("Title").InnerText; }
            catch { }
            try { this.channelImageUrl = node.SelectSingleNode("ImageUrl").InnerText; }
            catch { }                        
        }

        public override void LoadDataFromQueryString(HttpRequest request)
        {
            base.LoadDataFromQueryString(request);

            try { this.rssFeedUrl = request.QueryString["Url"]; }
            catch { }
            try { this.rssChannelPriority = (RssChannelPriority)FormUtil.GetNumber(request.QueryString["Pri"], 0, (int)RssChannelPriority.Max); }
            catch { }
            try { this.renderChannelImage = FormUtil.GetBoolean(request.QueryString["RCI"]); }
            catch { }
            try { this.renderChannelTitle = FormUtil.GetBoolean(request.QueryString["RCT"]); }
            catch { }
            try { this.renderItemTitles = FormUtil.GetBoolean(request.QueryString["RIT"]); }
            catch { }
            try { this.renderItemDescriptions = FormUtil.GetBoolean(request.QueryString["RID"]); }
            catch { }
            try { this.renderAd = FormUtil.GetBoolean(request.QueryString["RAD"]); }
            catch { }
            
        }

        public string RssFeedUrl
        {
            get { return this.rssFeedUrl; }
            set
            {
                if (this.rssFeedUrl != value)
                {
                    this.rssFeedUrl = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public RssChannelPriority RssChannelPriority
        {
            get { return this.rssChannelPriority; }
            set
            {
                if (this.rssChannelPriority != value)
                {
                    this.rssChannelPriority = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public bool RenderChannelImage
        {
            get { return this.renderChannelImage; }
            set
            {
                if (this.renderChannelImage != value)
                {
                    this.renderChannelImage = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public bool RenderChannelTitle
        {
            get { return this.renderChannelTitle; }
            set
            {
                if (this.renderChannelTitle != value)
                {
                    this.renderChannelTitle = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public bool RenderItemTitles
        {
            get { return this.renderItemTitles; }
            set
            {
                if (this.renderItemTitles != value)
                {
                    this.renderItemTitles = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public bool RenderItemDescriptions
        {
            get { return this.renderItemDescriptions; }
            set
            {
                if (this.renderItemDescriptions != value)
                {
                    this.renderItemDescriptions = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public bool RenderAd
        {
            get { return this.renderAd; }
            set
            {
                if (this.renderAd != value)
                {
                    this.renderAd = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }
        
        public string ChannelTitle
        {
            get { return this.channelTitle; }
        }

        public string ChannelImageUrl
        {
            get { return this.channelImageUrl; }
        }
        
        public override bool ItemsNeedCompileState
        {
            get
            {
                return true;
            }
        }        

        public override List<ListItem> InitItems(List<Guid> guidContext, DateTime dateContext, bool bypassCaches, System.Collections.Hashtable compileState)        
        { 	    
            List<ListItem> items = new List<ListItem>();

            Feed feed = Feed.LoadRssFeed(this.RssFeedUrl, this.rssChannelPriority);

            List<FeedItem> feedItems = feed.GetFeedItems(dateContext, bypassCaches);
            feed.Save();

            if (this.renderChannelTitle)
            {
                this.channelTitle = feed.Title;
            }

            if (this.renderChannelImage)
            {
                this.channelImageUrl = feed.LogoUrl;
            }
                    
            foreach (FeedItem feedItem in feedItems)
            {
                if ((this.rssChannelPriority == RssChannelPriority.Text && !String.IsNullOrEmpty(feedItem.description)) ||
                    (String.IsNullOrEmpty(feedItem.imageUrl))
                    )
                {
                    TextRssChannelItem textRssChannelItem = new TextRssChannelItem(feedItem.pubDate, feedItem.title, feedItem.description);
                    textRssChannelItem.Channel = this;
                    textRssChannelItem.ExpDate = feed.ItemsInvalidDate;
                    items.Add((ChannelItem)textRssChannelItem);
                }
                else
                {
                    UrlChannelItem urlChannelItem = new UrlChannelItem(feedItem.pubDate, feedItem.imageUrl, feedItem.imageWidth, feedItem.imageHeight);
                    urlChannelItem.Channel = this;
                    urlChannelItem.ExpDate = feed.ItemsInvalidDate;
                    items.Add((ChannelItem)urlChannelItem);
                }
            }

            return items;
        }               
    }
}
