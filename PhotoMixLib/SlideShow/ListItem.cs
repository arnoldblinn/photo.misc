using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Drawing;
using System.Web;

using System.Collections;

using Msn.Framework;

namespace Msn.PhotoMix.SlideShow
{
    public enum ListItemType
    {
        Unknown = 0,
        ListItem = 0,
        SlideShowItem = 1,
        ChannelItem = 2,
        TextRssChannelItem = 3,
        UrlChannelItem = 4
    }

    public class ListItem : IComparable
    {        
        // Type of the item
        private ListItemType type;

        // Date that the item was published
        private DateTime pubDate = DateTime.MinValue;

        // Date that the item "expires"
        private DateTime expDate = DateTime.MaxValue;

        int IComparable.CompareTo(object obj)
        {
            ListItem listItem = (ListItem)obj;

            // We want newer ones first, so the compare is backwards
            return DateTime.Compare(listItem.pubDate, this.pubDate);
        }

        public ListItem(ListItemType type)
        {
            this.type = type;
        }

        public ListItem(DateTime pubDate, ListItemType type)
        {
            this.type = type;
            this.pubDate = pubDate;
        }

        public ListItem(DateTime expDate, DateTime pubDate, ListItemType type)
        {
            this.expDate = expDate;
            this.type = type;
            this.pubDate = pubDate;
        }

        public ListItemType Type
        {
            get { return this.type; }
        }

        public DateTime PubDate
        {
            get { return this.pubDate; }
            set { this.pubDate = value; }
        }

        public DateTime ExpDate
        {
            get { return this.expDate; }
            set { this.expDate = value; }
        }

        public virtual void SaveDataToString(StringBuilder sb)
        {
            if (this.pubDate != DateTime.MinValue)
                sb.Append("<Date>" + HttpUtility.HtmlEncode(this.pubDate.ToString()) + "</Date>");
            if (this.expDate != DateTime.MaxValue)
                sb.Append("<ExpDate>" + HttpUtility.HtmlEncode(this.expDate.ToString()) + "</ExpDate>");
        }

        public virtual SlideShowItem CreateSlideShowItem(Hashtable compileCache, DateTime dateContext, bool bypassCaches)
        {
            return null;
        }

        public void SaveToString(StringBuilder sb)
        {
            sb.Append("<Item>");
            sb.Append("<Type>" + (int)(this.type) + "</Type>");
            this.SaveDataToString(sb);
            sb.Append("</Item>");
        }

        public virtual void LoadDataFromXmlNode(XmlNode node)
        {
            try
            {
                this.pubDate = FormUtil.GetDateTime(node.SelectSingleNode("Date").InnerText);
            }
            catch { }
            try
            {
                this.expDate = FormUtil.GetDateTime(node.SelectSingleNode("ExpDate").InnerText);
            }
            catch { }
        }        

        public static ListItem LoadFromXmlNode(XmlNode node)
        {
            ListItemType type = (ListItemType)Convert.ToInt32(node.SelectSingleNode("Type").InnerText);
            ListItem listItem;

            if (type == ListItemType.TextRssChannelItem)
                listItem = (ListItem)(new TextRssChannelItem());
            else if (type == ListItemType.UrlChannelItem)
                listItem = (ListItem)(new UrlChannelItem());
            else if (type == ListItemType.SlideShowItem)
                listItem = (ListItem)(new SlideShowItem());                        
            else if (type == ListItemType.ChannelItem)
                listItem = (ListItem)(new ChannelItem());
            else
                listItem = new ListItem(ListItemType.ListItem);

            listItem.LoadDataFromXmlNode(node);

            return listItem;
        }

        //
        // Utility functions to marshal/unmarshall a list of items 
        //
        static public List<ListItem> RootLoadListFromString(string input)
        {
            List<ListItem> items = new List<ListItem>();

            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(input);
            XmlNodeList itemNodes = xmlDocument.SelectNodes("Items/Item");
            foreach (XmlNode itemNode in itemNodes)
            {
                ListItem listItem = ListItem.LoadFromXmlNode(itemNode);

                items.Add(listItem);
            }

            return items;
        }

        static public void RootSaveListToString(StringBuilder sb, List<ListItem> items)
        {
            sb.Append("<Items>");
            foreach (ListItem listItem in items)
            {
                listItem.SaveToString(sb);
            }
            sb.Append("</Items>");
        }

    }
}
