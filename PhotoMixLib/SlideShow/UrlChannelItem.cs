using System;
using System.Collections.Generic;
using System.Text;

using System.Collections;

using System.Web;

namespace Msn.PhotoMix.SlideShow
{
    public class UrlChannelItem : ChannelItem
    {
        // URL for image, width, and height gotten from an enclosure or media RSS item
        private string url = null;
        private int width = -1;
        private int height = -1;

        public UrlChannelItem()
            : base(ListItemType.UrlChannelItem)
        {

        }

        public UrlChannelItem(DateTime pubDate, string url, int width, int height)
            : base(pubDate, ListItemType.UrlChannelItem)
        {
            this.url = url;
            this.width = width;
            this.height = height;
        }

        public string Url
        {
            get { return this.url; }
        }

        public int Width
        {
            get { return this.width; }
        }

        public int Height
        {
            get { return this.height; }
        }


        public override void SaveDataToString(StringBuilder sb)
        {
            base.SaveDataToString(sb);

            if (!String.IsNullOrEmpty(this.url))
                sb.Append("<Url>" + HttpUtility.HtmlEncode(this.url) + "</Url>");
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
                this.url = node.SelectSingleNode("Url").InnerText;
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
            return new SlideShowItem(this.ExpDate, this.PubDate, this.Channel.Name, this.Url, this.Width, this.Height);
        }
    }
}
