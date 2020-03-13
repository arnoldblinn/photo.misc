using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Drawing;
using System.Web;

using System.Collections;

namespace Msn.PhotoMix.SlideShow
{        
    public class ChannelItem : ListItem
    {
        // Pointer back to the slide show
        private Channel channel;

        public ChannelItem() : base(ListItemType.ChannelItem)
        {
        }

        public ChannelItem(DateTime expDate, DateTime pubDate)
            : base(expDate, pubDate, ListItemType.ChannelItem)
        {
        }

        public ChannelItem(ListItemType type) : base(type)
        {        
        }

        public ChannelItem(DateTime pubDate, ListItemType type)
            : base(pubDate, type)
        {
        }

        public ChannelItem(DateTime pubDate)
            : base(pubDate, ListItemType.ChannelItem)
        {

        }

        public ChannelItem(DateTime expDate, DateTime pubDate, ListItemType type)
            : base(expDate, pubDate, type)
        {
        }

        public Channel Channel
        {
            get { return this.channel; }
            set { this.channel = value; }
        }

        public override SlideShowItem CreateSlideShowItem(Hashtable compileCache, DateTime dateContext, bool bypassCaches)
        {
            return this.Channel.CreateSlideShowItem(compileCache, dateContext, bypassCaches);
        }                                     
    }

}
