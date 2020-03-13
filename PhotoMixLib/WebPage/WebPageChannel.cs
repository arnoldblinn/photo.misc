using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Drawing;

using System.Web;

using Msn.Framework;
using Msn.PhotoMix.Passport;

namespace Msn.PhotoMix.SlideShow
{
    public enum WebPageType
    {
        FullPage = 0,
        ImagesOnly = 1,

        Max = 1
    }

    public class WebPageChannel : Channel
    {
        private string url;
        private int width = 640;
        private int height = 480;
        private WebPageType webPageType = WebPageType.FullPage;

        static private int webPageChannelExpiresTTL = Convert.ToInt32(Config.GetSetting("WebPageChannelExpiresTTL"));

        public WebPageChannel(Puid puid, Guid slideShowGuid, Guid channelGuid)
            : base(puid, slideShowGuid, channelGuid, ChannelType.WebPage)
        {
        }

        public WebPageChannel(Puid puid, Guid slideShowGuid)
            : base(puid, slideShowGuid, ChannelType.WebPage)
        {
        }

        public string Url
        {
            get { return this.url; }
            set
            {
                if (this.url != value)
                {
                    this.url = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public WebPageType WebPageType
        {
            get { return this.webPageType; }
            set
            {
                if (this.webPageType != value)
                {
                    this.ChangedItemsNeedCompileState();
                    this.webPageType = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public override bool ItemsNeedCompileState
        {
            get
            {
                return (this.webPageType == WebPageType.FullPage);
            }
        }
        
        public override List<ListItem> InitItems(List<Guid> guidContext, DateTime dateContext, bool bypassCaches, Hashtable compileState)
        {
            List<ListItem> items = new List<ListItem>();

            if (this.webPageType == WebPageType.FullPage)
            {
                ChannelItem channelItem = new ChannelItem(dateContext);
                channelItem.Channel = this;
                items.Add(channelItem);
            }
            else
            {

                Feed feed = Feed.LoadWebPage(this.url);

                List<FeedItem> feedItems = feed.GetFeedItems(dateContext, bypassCaches);
                feed.AddUrlChannelItems(this, items, dateContext, bypassCaches, this.width, this.height);                
            }

            return items;
        }

        public override void SaveDataToString(StringBuilder sb)
        {
            base.SaveDataToString(sb);

            if (!String.IsNullOrEmpty(this.url))
                sb.Append("<Url>" + HttpUtility.HtmlEncode(this.url) + "</Url>");
            if (this.webPageType != WebPageType.FullPage)
                sb.Append("<WPT>" + (int)this.webPageType + "</WPT>");
            if (this.width != 640)
                sb.Append("<WPW>" + this.width + "</WPW>");
            if (this.height != 640)
                sb.Append("<WPH>" + this.height + "</WPH>");
        }


        public override void LoadDataFromXmlNode(System.Xml.XmlNode node)
        {
            base.LoadDataFromXmlNode(node);

            try { this.url = node.SelectSingleNode("Url").InnerText; }
            catch { }
            try { this.webPageType = (WebPageType) Convert.ToInt32(node.SelectSingleNode("WPT").InnerText); }
            catch { }
            try { this.width = Convert.ToInt32(node.SelectSingleNode("WPW").InnerText); }
            catch { }
            try { this.height = Convert.ToInt32(node.SelectSingleNode("WPH").InnerText); }
            catch { }
        }

        public override void LoadDataFromQueryString(HttpRequest request)
        {
            base.LoadDataFromQueryString(request);

            if (!String.IsNullOrEmpty(request.QueryString["Url"]))
                this.url = FormUtil.GetUrl(request.QueryString["Url"]);
            if (!String.IsNullOrEmpty(request.QueryString["WPT"]))
                this.webPageType = (WebPageType)FormUtil.GetNumber(request.QueryString["WPT"], 0, (int)WebPageType.Max);
            if (!String.IsNullOrEmpty(request.QueryString["WPW"]))
                this.width = FormUtil.GetNumber(request.QueryString["WPW"]);
            if (!String.IsNullOrEmpty(request.QueryString["WPH"]))
                this.height = FormUtil.GetNumber(request.QueryString["WPH"]);
        }

        public override SlideShowItem CreateSlideShowItem(Hashtable compileCache, DateTime dateContext, bool bypassCaches)
        {
            if (this.webPageType == WebPageType.ImagesOnly)
                return base.CreateSlideShowItem(compileCache, dateContext, bypassCaches);
            else
            {
                CompiledWebPage compiledWebPage;
                if (compileCache[this.ChannelGuid] != null)
                {
                    compiledWebPage = (CompiledWebPage)compileCache[this.ChannelGuid];
                }
                else
                {
                    compiledWebPage = CompiledWebPage.LoadForCompile(this.url, dateContext);
                    compileCache[this.ChannelGuid] = compiledWebPage;
                }

                string url = Config.GetSetting("CompiledImageBaseUrl") + "webpageitem.ashx?" +
                    "ch=" + compiledWebPage.CompiledWebPageHash + 
                    "&cid=" + compiledWebPage.CompiledWebPageGuid +
                    (bypassCaches ? "&bc=1" : "") +
                    "&ti=" + compiledWebPage.FetchDataDate.Ticks;

                return new SlideShowItem(compiledWebPage.FetchDataDate.Add(new TimeSpan(0, WebPageChannel.webPageChannelExpiresTTL, 0)), compiledWebPage.FetchDataDate, this.Name, url, 600, 400);
            }
        }
    }
}
