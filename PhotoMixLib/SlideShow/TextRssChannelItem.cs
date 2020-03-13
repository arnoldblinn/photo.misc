using System;
using System.Collections.Generic;
using System.Text;

using System.Collections;

using System.Web;
using Msn.Framework;

namespace Msn.PhotoMix.SlideShow
{
    public class TextRssChannelItem : ChannelItem
    {

        // Title and description gotten for an RSS item
        private string title = null;
        private string description = null;

        public TextRssChannelItem()
            : base(ListItemType.TextRssChannelItem)
        {

        }

        public TextRssChannelItem(DateTime pubDate, string title, string description)
            : base(pubDate, ListItemType.TextRssChannelItem)
        {
            this.title = title;
            this.description = description;
        }

        public string Title
        {
            get { return this.title; }
        }

        public string Description
        {
            get { return this.description; }
        }

        public override void SaveDataToString(StringBuilder sb)
        {
            base.SaveDataToString(sb);

            if (!String.IsNullOrEmpty(this.title))
                sb.Append("<Title>" + HttpUtility.HtmlEncode(this.title) + "</Title>");
            if (!String.IsNullOrEmpty(this.description))
                sb.Append("<Description>" + HttpUtility.HtmlEncode(this.description) + "</Description>");
        }

        public override void LoadDataFromXmlNode(System.Xml.XmlNode node)
        {
            base.LoadDataFromXmlNode(node);

            try
            {
                this.title = node.SelectSingleNode("Title").InnerText;
            }
            catch { }
            try
            {
                this.description = node.SelectSingleNode("Description").InnerText;
            }
            catch { }
        }

        public override SlideShowItem CreateSlideShowItem(Hashtable compileCache, DateTime dateContext, bool bypassCaches)
        {
            RssChannel rssFeedChannel = (RssChannel)this.Channel;
            
            CompiledTextFeed compiledTextFeed;
            if (compileCache[rssFeedChannel.ChannelGuid] != null)
            {
                compiledTextFeed = (CompiledTextFeed)compileCache[rssFeedChannel.ChannelGuid];
            }
            else
            {
                compiledTextFeed = CompiledTextFeed.LoadForCompile(rssFeedChannel.RssFeedUrl, rssFeedChannel.ChannelImageUrl, rssFeedChannel.ChannelTitle, rssFeedChannel.RenderAd);
                compileCache[rssFeedChannel.ChannelGuid] = compiledTextFeed;
            }
            
            CompiledTextFeedItem compiledTextFeedItem;
            if (compileCache[rssFeedChannel.ChannelGuid + this.title + this.description] != null)
            {
                compiledTextFeedItem = (CompiledTextFeedItem)compileCache[rssFeedChannel.ChannelGuid + this.title + this.description];
            }
            else
            {
                compiledTextFeedItem = CompiledTextFeedItem.LoadForCompile(compiledTextFeed.CompiledTextFeedHash, compiledTextFeed.CompiledTextFeedGuid, this.title, this.description);
                compileCache[rssFeedChannel.ChannelGuid + this.title + this.description] = compiledTextFeedItem;
            }

            string url = Config.GetSetting("CompiledImageBaseUrl") + "textfeeditem.ashx?" +
                    "ch=" + compiledTextFeed.CompiledTextFeedHash + 
                    "&cid=" + compiledTextFeed.CompiledTextFeedGuid + 
                    "&tih=" + compiledTextFeedItem.CompiledTextFeedItemHash +
                    (bypassCaches ? "&bc=1" : "") +
                    "&<SIZE>";

            return new SlideShowItem(this.ExpDate, this.PubDate, this.Channel.Name, url, true);            
        }

    }
}
