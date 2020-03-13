using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Drawing;
using System.IO;

using System.Web;

using Msn.Framework;
using Msn.PhotoMix.Passport;

namespace Msn.PhotoMix.SlideShow
{                       
    public class USTrafficChannel : Channel
    {
        // A Url to a traffic image is of the form root + w= + h= + params
        // e.g. http://api.tiles.virtualearth.net/api/GetMap.ashx?w=640&h=480&b=r&z=10&c=47.6410703222025,-122.28044433594&t=f

        // Query string root for the traffic image
        private static string trafficUrlRoot = "http://api.tiles.virtualearth.net/api/GetMap.ashx";

        private int trafficRegion = -1;

        // Time to live for traffic
        private static int trafficImageTTL = Convert.ToInt32(Config.GetSetting("USTrafficImageTTL"));
        static private int trafficChannelExpiresTTL = Convert.ToInt32(Config.GetSetting("USTrafficChannelExpiresTTL"));

        static public string [,] USTrafficRegionData = 
        {
            {"Atlanta", "b=r&z=10&c=33.8459,-84.32175&t=f"}, 
            {"Boston", "b=r&z=10&c=42.4063,-71.146&t=f"}, 
            {"Chicago", "b=r&z=10&c=41.85905,-87.99125&t=f"}, 
            {"Dallas/Ft. Worth", "b=r&z=10&c=32.65822,-97.01505&t=f"}, 
            {"Denver", "b=r&z=10&c=39.839,-105.026&t=f"}, 
            {"Detroit", "b=r&z=10&c=42.334,-83.514&t=f"}, 
            {"Houston", "b=r&z=10&c=29.735,-95.403&t=f"}, 
            {"Indianapolis", "b=r&z=10&c=39.84686,-86.13002&t=f"}, 
            {"Las Vegas", "b=r&z=10&c=36.1412,-115.0743&t=f"}, 
            {"Los Angeles", "b=r&z=10&c=33.91295,-118.1865&t=f"}, 
            {"Milwaukee", "b=r&z=10&c=43.0727,-88.2715&t=f"}, 
            {"Minneapolis", "b=r&z=10&c=44.9798,-93.1923&t=f"}, 
            {"New York", "b=r&z=10&c=40.7045,-73.5465&t=f"}, 
            {"Oklahoma City", "b=r&z=10&c=35.42055,-97.4973&t=f"}, 
            {"Philadelphia", "b=r&z=10&c=40.031,-75.124&t=f"}, 
            {"Phoenix", "b=r&z=10&c=33.485,-112.006&t=f"}, 
            {"Pittsburgh", "b=r&z=10&c=40.512,-79.975&t=f"}, 
            {"Portland", "b=r&z=10&c=45.52375,-122.6416&t=f"}, 
            {"Providence", "b=r&z=10&c=41.71535,-71.3938&t=f"}, 
            {"Sacramento", "b=r&z=10&c=38.2765,-120.801&t=f"}, 
            {"Salt Lake City", "b=r&z=10&c=40.75,-111.919&t=f"}, 
            {"San Antonio", "b=r&z=10&c=29.467,-98.32899&t=f"}, 
            {"San Diego", "b=r&z=10&c=32.912,-116.9165&t=f"}, 
            {"San Francisco", "b=r&z=10&c=37.89454,-122.3204&t=f"}, 
            {"Seattle", "b=r&z=10&c=47.5705,-122.499&t=f"}, 
            {"St Louis", "b=r&z=10&c=38.616,-90.348&t=f"}, 
            {"Tampa", "b=r&z=10&c=27.8692,-82.05925&t=f"}, 
            {"Toronto", "b=r&z=10&c=43.53894,-79.24803&t=f"}, 
            {"Washington DC", "b=r&z=10&c=38.8813,-77.0149&t=f"}
        };
        
        

        public USTrafficChannel(Puid puid, Guid slideShowGuid, Guid channelGuid)
            : base(puid, slideShowGuid, channelGuid, ChannelType.USTraffic)
        {
        }

        public USTrafficChannel(Puid puid, Guid slideShowGuid)
            : base(puid, slideShowGuid, ChannelType.USTraffic)
        {
        }
        
        public override bool IsFixedCount
        {
            get { return true; }
        }        
        
        public int TrafficRegion
        {
            get { return this.trafficRegion; }
            set
            {
                if (this.trafficRegion != value)
                {
                    this.trafficRegion = value;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public override List<ListItem> InitItems(List<Guid> guidContext, DateTime dateContext, bool bypassCaches, Hashtable compileState)
        {
            List<ListItem> listItems = new List<ListItem>();

            
            ChannelItem channelItem = new ChannelItem(ListItemType.ChannelItem);
            channelItem.Channel = this;
            listItems.Add(channelItem);

            return listItems;            
        }

        public override void SaveDataToString(StringBuilder sb)
        {
            base.SaveDataToString(sb);
            
            sb.Append("<Region>" + ((int)this.trafficRegion).ToString() + "</Region>");            
        }

        public override void LoadDataFromXmlNode(System.Xml.XmlNode node)
        {
            base.LoadDataFromXmlNode(node);

            try { this.trafficRegion = Convert.ToInt32(node.SelectSingleNode("Region").InnerText); }
            catch { }
        }

        public override void LoadDataFromQueryString(HttpRequest request)
        {
            base.LoadDataFromQueryString(request);

            if (!String.IsNullOrEmpty(request.QueryString["Region"]))
                this.trafficRegion = Convert.ToInt32(request.QueryString["Region"]);
        }

        public static string GetImageFileName(int trafficRegion, SlideShowImageSize imageSize, bool bypassCaches)
        {
            string fileName = ImageUtil.GetCompiledImageDirectory("USTraffic") + trafficRegion.ToString() + "_" + ((int)imageSize).ToString() + ".jpg";
            if (!bypassCaches && MiscUtil.TTLFileExists(fileName, USTrafficChannel.trafficImageTTL))
            {
                return fileName;
            }
            else
            {
                if (trafficRegion == -1)
                    trafficRegion = 0;

                // Max width and height that maps can return is is 800....
                int width = SlideShow.slideShowImageWidths[(int)imageSize];
                width = width > 800 ? 800 : width;
                int height = SlideShow.slideShowImageHeights[(int)imageSize];
                height = height > 800 ? 800 : height;
                
                string url = USTrafficChannel.trafficUrlRoot + 
                    "?w=" + width + 
                    "&h=" + height + 
                    (bypassCaches ? "&bc=1" : "") +
                    "&" + USTrafficChannel.USTrafficRegionData[(int)trafficRegion, 1];

                Bitmap bitmap = ImageUtil.LoadImageFromUrl(url);

                ImageUtil.SaveJpeg(fileName, bitmap, 100);

                return fileName;
            }
        }

        public override SlideShowItem CreateSlideShowItem(Hashtable compileCache, DateTime dateContext, bool bypassCaches)
        {            
            string url = Config.GetSetting("CompiledImageBaseUrl") + "ustraffic.ashx" +
                   "?tr=" + ((int)this.trafficRegion).ToString() + 
                   "&<SIZE>" +
                   (bypassCaches ? "&bc=1" : "") +
                   "&" + MiscUtil.TimeIndex();              

            return new SlideShowItem(dateContext.Add(new TimeSpan(0, USTrafficChannel.trafficChannelExpiresTTL, 0)), dateContext, this.Name, url, true);
        }
    }
}
