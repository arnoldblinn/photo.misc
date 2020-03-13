using System;
using System.Collections.Generic;
using System.Text;

using System.Collections;

using System.Web;

using Msn.Framework;
using Msn.PhotoMix.Util;

namespace Msn.PhotoMix.SlideShow
{
    public class SlideShowItem : ListItem
    {
        private string title;
        private string url;
        private bool urlSizeAware = false;
        private int width = -1;
        private int height = -1;

        public SlideShowItem()
            : base(ListItemType.SlideShowItem)
        {

        }

        public SlideShowItem(DateTime expDate, DateTime pubDate, string title, string url)
            : base(expDate, pubDate, ListItemType.SlideShowItem)
        {            
            this.title = title;
            this.url = url;            
        }

        public SlideShowItem(DateTime expDate, DateTime pubDate, string title, string url, bool urlSizeAware)
            : base(expDate, pubDate, ListItemType.SlideShowItem)
        {
            this.title = title;
            this.url = url;
            this.urlSizeAware = urlSizeAware;            
        }

        public SlideShowItem(DateTime expDate, DateTime pubDate, string title, string url, int width, int height)
            : base(expDate, pubDate, ListItemType.SlideShowItem)
        {
            this.title = title;
            this.url = url;
            this.width = width;
            this.height = height;
        }

        public override void SaveDataToString(StringBuilder sb)
        {
            base.SaveDataToString(sb);

            sb.Append("<Title>" + HttpUtility.HtmlEncode(this.title) + "</Title>");
            sb.Append("<Url>" + HttpUtility.HtmlEncode(this.url) + "</Url>");
            if (this.urlSizeAware)
                sb.Append("<UrlSA>" + HttpUtility.HtmlEncode(this.urlSizeAware.ToString()) + "</UrlSA>");            
            if (this.width != -1)
                sb.Append("<Width>" + this.width + "</Width>");
            if (this.height != -1)
                sb.Append("<Height>" + this.height + "</Height>");            
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
                this.url = node.SelectSingleNode("Url").InnerText;
            }
            catch { }
            try
            {
                this.urlSizeAware = Convert.ToBoolean(node.SelectSingleNode("UrlSA").InnerText);
            }
            catch { }            
            try
            {
                this.width = Convert.ToInt32(node.SelectSingleNode("Width").InnerText);
            }
            catch { }
            try
            {
                this.height = Convert.ToInt32(node.SelectSingleNode("Height").InnerText);
            }
            catch { }            
        }

        public override SlideShowItem CreateSlideShowItem(Hashtable compileCache, DateTime dateContext, bool bypassCaches)
        {
            return this;
        }        

        public void GenerateRss(StringBuilder sb, SlideShowRssMediaType rssType, SlideShowImageSize targetImageSize, bool insertSizes, bool debug)
        {
            string url = this.url;
            int width = this.width;
            int height = this.height;

            if (this.urlSizeAware)
            {
                url = url.Replace("<SIZE>", "s=" + ((int)targetImageSize).ToString());

                width = SlideShow.slideShowImageWidths[(int)targetImageSize];
                height = SlideShow.slideShowImageHeights[(int)targetImageSize];                
            }

            if (insertSizes)
            {
                if (width == -1)
                    width = SlideShow.slideShowImageWidths[(int)targetImageSize];
                if (height == -1)
                    height = SlideShow.slideShowImageHeights[(int)targetImageSize];
            }

            string linkUrl = "";
            if (rssType == SlideShowRssMediaType.LinkMedia)
            {
                linkUrl = url;
            }
            else
            {
                linkUrl = "http://" + Config.GetSetting("HostSite");
            }

            sb.Append("<item>");
            sb.Append("<title>" + HttpUtility.HtmlEncode(this.title) + "</title>");
            sb.Append("<link>" + HttpUtility.HtmlEncode(linkUrl) + "</link>");
            sb.Append("<category>" + HttpUtility.HtmlEncode(this.title) + "</category>");
            
            // Write out the description
            sb.Append("<description>");
            sb.Append("<![CDATA[");
            sb.Append("<img src=\"" + url + "\"><br/>" + HttpUtility.HtmlEncode(this.title));
            if (debug)
            {
                sb.Append("<br/>");
                sb.Append("Expiraton date " + this.ExpDate.ToString() + "<br/>");
                sb.Append("<a href=\"" + url + "\">" + HttpUtility.HtmlEncode(url) + "</a><br/>");
            }
            sb.Append("]]>");
            sb.Append("</description>");

            // Write out the publish date (if we have one)
            // Format of date in RSS feed: Tue, 22 Jan 2008 20:26:39 -05:00                
            if (this.PubDate != DateTime.MinValue) 
                sb.Append("<pubDate>" + TimeUtil.DateToRSSString(this.PubDate) + "</pubDate>");

            
            if (rssType != SlideShowRssMediaType.LinkMedia)
            {
                string widthText = "";
                string heightText = "";
                string tag = "";
                string durationText = "";

                if (rssType == SlideShowRssMediaType.EnclosureRss)
                {
                    tag = "enclosure";
                }
                else if (rssType == SlideShowRssMediaType.MediaRss)
                {
                    tag = "media:content";                    

                    if (width != -1)
                        widthText = "width=\"" + width + "\" ";

                    if (height != -1)
                        heightText = "height=\"" + height + "\" ";
                }

                sb.Append("<" + tag + " type=\"image/jpeg\" " + widthText + heightText + "url=\"" + HttpUtility.HtmlEncode(url) + "\" " + durationText + "/>");
            }

            sb.Append("</item>");
        } 
    }
}
