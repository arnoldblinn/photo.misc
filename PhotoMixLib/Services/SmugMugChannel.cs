using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Web;
using System.Xml;

using Msn.PhotoMix.Passport;

namespace Msn.PhotoMix.SlideShow
{
    public enum SmugMugFeedType
    {
        Unknown = 0,
        Recent = 1,
        Popular = 2
    }

    public class SmugMugChannel : Channel
    {
        private string userName = null;
        private SmugMugFeedType smugMugFeedType = SmugMugFeedType.Unknown;

        public SmugMugChannel(Puid puid, Guid slideShowGuid, Guid channelGuid)
            : base(puid, slideShowGuid, channelGuid, ChannelType.SmugMug)
        {
        }

        public SmugMugChannel(Puid puid, Guid slideShowGuid)
            : base(puid, slideShowGuid, ChannelType.SmugMug)
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

        public SmugMugFeedType SmugMugFeedType
        {
            get { return this.smugMugFeedType; }
            set
            {
                if (this.smugMugFeedType != value)
                {
                    this.smugMugFeedType = value;
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
            sb.Append("<SmugMugFeedType>" + this.smugMugFeedType.ToString() + "</SmugMugFeedType>");
        }

        public override void LoadDataFromXmlNode(XmlNode node)
        {
            base.LoadDataFromXmlNode(node);

            try { this.userName = node.SelectSingleNode("UserName").InnerText; }
            catch { }
            try { this.smugMugFeedType = (SmugMugFeedType)Convert.ToInt32(node.SelectSingleNode("SmugMugFeedType").InnerText); }
            catch { }
        }        

        public override List<ListItem> InitItems(List<Guid> guidContext, DateTime dateContext, bool bypassCaches, System.Collections.Hashtable compileState)
        {
            List<ListItem> items = new List<ListItem>();
            
            Feed feed = Feed.LoadSmugMugFeed(this.userName, this.smugMugFeedType);

            feed.AddUrlChannelItems(this, items, dateContext, bypassCaches);

            return items;            
        }
        
    }

}
