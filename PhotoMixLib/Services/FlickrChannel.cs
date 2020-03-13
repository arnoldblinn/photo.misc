using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Web;
using System.Xml;

using Msn.PhotoMix.Passport;

namespace Msn.PhotoMix.SlideShow
{
    public class FlickrChannel : Channel
    {
        private string userName = null;

        public FlickrChannel(Puid puid, Guid slideShowGuid, Guid channelGuid)
            : base(puid, slideShowGuid, channelGuid, ChannelType.Flickr)
        {
        }

        public FlickrChannel(Puid puid, Guid slideShowGuid)
            : base(puid, slideShowGuid, ChannelType.Flickr)
        {
        }        

        public string UserName
        {
            get { return this.userName; }
            set
            {
                if (this.userName != value)
                {
                    this.userName = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public override void SaveDataToString(StringBuilder sb)
        {
            base.SaveDataToString(sb);

            if (!String.IsNullOrEmpty(this.userName))
                sb.Append("<UserName>" + this.userName.ToString() + "</UserName>");
        }

        public override void LoadDataFromXmlNode(XmlNode node)
        {
            base.LoadDataFromXmlNode(node);

            try { this.userName = node.SelectSingleNode("UserName").InnerText; }
            catch { }
        }

        public override List<ListItem> InitItems(List<Guid> guidContext, DateTime dateContext, bool bypassCaches, System.Collections.Hashtable compileState)
        {
            List<ListItem> items = new List<ListItem>();

            Feed feed = Feed.LoadFlickrFeed(this.userName);

            feed.AddUrlChannelItems(this, items, dateContext, bypassCaches);

            return items;
        }        

    }

}
