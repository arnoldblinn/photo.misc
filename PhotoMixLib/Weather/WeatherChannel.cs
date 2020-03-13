using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

using System.Collections;

using System.Web;

using Msn.Framework;
using Msn.PhotoMix.Passport;

namespace Msn.PhotoMix.SlideShow
{
    public class WeatherChannel : Channel
    {
        private string language = "en-US";
        private string location;

        static private int weatherChannelExpiresTTL = Convert.ToInt32(Config.GetSetting("WeatherChannelExpiresTTL"));

        public WeatherChannel(Puid puid, Guid slideShowGuid, Guid channelGuid)
            : base(puid, slideShowGuid, channelGuid, ChannelType.Weather)
        {
        }

        public WeatherChannel(Puid puid, Guid slideShowGuid)
            : base(puid, slideShowGuid, ChannelType.Weather)
        {
        }

        public string Language
        {
            get { return this.language; }
            set
            {
                if (this.language != value)
                {
                    this.language = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public string Location
        {
            get { return this.location; }
            set
            {
                if (this.location != value)
                {
                    this.location = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public override bool IsFixedCount
        {
            get { return true; }
        }

        public override bool ItemsNeedCompileState
        {
            get
            {
                return true;
            }
        }        

        public override List<ListItem> InitItems(List<Guid> guidContext, DateTime dateContext, bool bypassCaches, Hashtable compileState)
        {
            List<ListItem> listItems = new List<ListItem>();

            //WeatherChannelItem weatherChannelItem = new WeatherChannelItem(dateContext);
            //weatherChannelItem.Channel = this;
            ChannelItem channelItem = new ChannelItem(ListItemType.ChannelItem);
            channelItem.Channel = this;

            listItems.Add(channelItem);

            return listItems;
        }

        public override void SaveDataToString(StringBuilder sb)
        {
            base.SaveDataToString(sb);

            if (!String.IsNullOrEmpty(this.language))
                sb.Append("<Language>" + HttpUtility.HtmlEncode(this.language) + "</Language>");
            if (!String.IsNullOrEmpty(this.location))
                sb.Append("<Location>" + HttpUtility.HtmlEncode(this.location) + "</Location>");
        }

        public override void LoadDataFromXmlNode(System.Xml.XmlNode node)
        {
            base.LoadDataFromXmlNode(node);

            try { this.language = node.SelectSingleNode("Language").InnerText; }
            catch { }
            try { this.location = node.SelectSingleNode("Location").InnerText; }
            catch { }
        }

        public override void LoadDataFromQueryString(HttpRequest request)
        {
            base.LoadDataFromQueryString(request);

            if (!String.IsNullOrEmpty(request.QueryString["Language"]))
                this.language = FormUtil.GetString(request.QueryString["Language"], true, 0, 10);
            if (!String.IsNullOrEmpty(request.QueryString["Location"]))
                this.location = FormUtil.GetString(request.QueryString["Location"], true, 0, 10);

        }

        public override SlideShowItem CreateSlideShowItem(Hashtable compileCache, DateTime dateContext, bool bypassCaches)
        {           
            CompiledWeather compiledWeather;
            if (compileCache[this.ChannelGuid] != null)
            {
                compiledWeather = (CompiledWeather)compileCache[this.ChannelGuid];
            }
            else
            {
                compiledWeather = CompiledWeather.LoadForCompile(this.Language, this.Location, dateContext);
                compileCache[this.ChannelGuid] = compiledWeather;
            }

            string url = Config.GetSetting("CompiledImageBaseUrl") + "weatheritem.ashx" +
                "?lan=" + this.Language +
                "&loc=" + this.Location +
                "&<SIZE>" +
                (bypassCaches ? "&bc=1" : "") +
                "&ti=" + compiledWeather.FetchDataDate.Ticks;

            return new SlideShowItem(compiledWeather.FetchDataDate.Add(new TimeSpan(0, WeatherChannel.weatherChannelExpiresTTL, 0)), dateContext, this.Name, url, true);
        }
    }
}
