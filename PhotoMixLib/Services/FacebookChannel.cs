using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Web;
using System.Xml;

using Msn.Framework;
using Msn.PhotoMix.Passport;

namespace Msn.PhotoMix.SlideShow
{
    public class FacebookChannel : Channel
    {
        private string sessionKey = null;

        public FacebookChannel(Puid puid, Guid slideShowGuid, Guid channelGuid)
            : base(puid, slideShowGuid, channelGuid, ChannelType.Flickr)
        {
        }

        public FacebookChannel(Puid puid, Guid slideShowGuid)
            : base(puid, slideShowGuid, ChannelType.Flickr)
        {
        }        

        public string SessionKey
        {
            get { return this.sessionKey; }
            set
            {
                if (this.sessionKey != value)
                {
                    this.sessionKey = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public override void SaveDataToString(StringBuilder sb)
        {
            base.SaveDataToString(sb);

            if (!String.IsNullOrEmpty(this.sessionKey))
                sb.Append("<SessionKey>" + this.sessionKey.ToString() + "</SessionKey>");
        }

        public override void LoadDataFromXmlNode(XmlNode node)
        {
            base.LoadDataFromXmlNode(node);

            try { this.sessionKey = node.SelectSingleNode("SessionKey").InnerText; }
            catch { }
        }
        
       
        public override List<ListItem> InitItems(List<Guid> guidContext, DateTime dateContext, bool bypassCaches, Hashtable compileState)
        {
            List<ListItem> items = new List<ListItem>();

            Feed feed = Feed.LoadFacebookFeed(this.sessionKey);
            feed.AddUrlChannelItems(this, items, dateContext, bypassCaches);

            return items;                  
        }
    }

}
