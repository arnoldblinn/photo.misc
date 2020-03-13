using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

using System.Web;

using Msn.Framework;
using Msn.PhotoMix.Passport;

namespace Msn.PhotoMix.SlideShow
{
    public class StaticChannel : Channel
    {
        private string imageUrl;

        private bool addTimeIndex;
        

        public StaticChannel(Puid puid, Guid slideShowGuid, Guid channelGuid)
            : base(puid, slideShowGuid, channelGuid, ChannelType.Static)
        {
        }

        public StaticChannel(Puid puid, Guid slideShowGuid)
            : base(puid, slideShowGuid, ChannelType.Static)
        {
        }

        public bool AddTimeIndex
        {
            get { return this.addTimeIndex; }
            set
            {
                if (this.addTimeIndex != value)
                {
                    this.addTimeIndex = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public string ImageUrl
        {
            get { return this.imageUrl; }
            set 
            {
                if (this.imageUrl != value)
                {
                    this.imageUrl = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public override bool IsFixedCount
        {
            get { return true; }
        }        

        public override List<ListItem> InitItems(List<Guid> guidContext, DateTime dateContext, bool bypassCaches, Hashtable compileState)
        {
            List<ListItem> listItems = new List<ListItem>();

            string url = this.imageUrl;
            if (this.addTimeIndex)
            {
                url = MiscUtil.AppendQueryStringParameter(url, MiscUtil.TimeIndex());
            }

            UrlChannelItem urlChannelItem = new UrlChannelItem(dateContext, url, -1, -1);
            urlChannelItem.Channel = this;

            listItems.Add(urlChannelItem);

            return listItems;
        }

        public override void SaveDataToString(StringBuilder sb)
        {
            base.SaveDataToString(sb);

            if (!String.IsNullOrEmpty(this.imageUrl))
                sb.Append("<Url>" + HttpUtility.HtmlEncode(this.imageUrl) + "</Url>");
            if (this.addTimeIndex)
                sb.Append("<TI>True</TI>");
        }

        public override void LoadDataFromXmlNode(System.Xml.XmlNode node)
        {
            base.LoadDataFromXmlNode(node);

            try { this.imageUrl = node.SelectSingleNode("Url").InnerText; }
            catch { }
            try { this.addTimeIndex = (node.SelectSingleNode("TI").InnerText == "True"); }
            catch { } 
        }

        public override void LoadDataFromQueryString(HttpRequest request)
        {
            base.LoadDataFromQueryString(request);

            if (!String.IsNullOrEmpty(request.QueryString["Url"]))
                this.imageUrl = FormUtil.GetUrl(request.QueryString["Url"]);
            if (!String.IsNullOrEmpty(request.QueryString["TI"]))
                this.addTimeIndex = (request.QueryString["TI"] == "1");
        }
    }
}
