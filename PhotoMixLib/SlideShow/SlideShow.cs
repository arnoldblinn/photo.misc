using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Web;


using System.Collections;
using System.Collections.Specialized;

using Msn.Framework;
using Msn.PhotoMix.Passport;
using Msn.PhotoMix.Util;

namespace Msn.PhotoMix.SlideShow
{    
    public enum SlideShowImageSize
    {
        Size640x480 = 0,
        Size800x600 = 1,
        Size1600x900 = 2,
        Size1024x768 = 3,
        SizeMax = 3,

        SizeUnknown = Int32.MaxValue
    }
    
    public enum SlideShowRssMediaType
    {
        MediaRss = 0,
        EnclosureRss = 1,
        LinkMedia = 2,
        Max = 3,

        Unknown = Int32.MaxValue

    }

    public struct SlideShowInfo
    {
        public Guid slideShowGuid;
        public string name;
    }

    public class SlideShow : ICompile
    {
        public static int[] slideShowImageWidths = { 640, 800, 1600, 1024 };
        public static int[] slideShowImageHeights = { 480, 600, 900, 768 };

        // Puid owner of this slideshow
        private Puid puid;

        // Guid of the slide show
        private Guid guid;       
        
        // Informational values
        private string name;
        private string pin;
        private string description;

        private PMTimeZone pmTimeZone = PMTimeZone.GMT;

        // Default size and media type
        private SlideShowImageSize defaultImageSize = SlideShowImageSize.Size640x480;
        private SlideShowRssMediaType defaultRssMediaType = SlideShowRssMediaType.EnclosureRss;

        // Friendly name from our table
        private string friendlyName;
        private bool friendlyNameVerified;

        // Flag indicating if there is state data to be saved
        private bool dataDirty = false;
        
        // Compilation settings and structures
        private bool compileAll = true;
        private int compileMaxCount = 5;
        private int compileMinCount = 1;
        private CompileOrder compileOrder = CompileOrder.Listed;
        private CompileInclude compileInclude = CompileInclude.All;
        private int compileAge = 10;

        // Compilation state data
        private List<ListItem> compiledItemsCache = null;
        private DateTime compiledItemsCacheDate = DateTime.MinValue;
        private DateTime compiledItemsCacheInvalidDate = DateTime.MaxValue;
        private bool compiledItemsCacheDirty = false;
        
        // Channels in the slide show
        private List<Channel> channels;

        private DateTime creationDate;

        //
        // Constructors
        //        
        public SlideShow()
        {
        }        

        public static SlideShow InitNew(Puid puid, string name, string pin, string description, PMTimeZone timeZone, SlideShowImageSize defaultImageSize, SlideShowRssMediaType defaultRssMediaType)
        {
            SlideShow slideShow = new SlideShow();
            
            slideShow.guid = Guid.NewGuid();
            slideShow.puid = puid;
            slideShow.name = name;
            slideShow.pin = pin;
            slideShow.description = description;
            slideShow.pmTimeZone = timeZone;
            slideShow.channels = new List<Channel>();
            slideShow.dataDirty = true;
            slideShow.compiledItemsCacheDirty = false;
            slideShow.defaultImageSize = defaultImageSize;
            slideShow.defaultRssMediaType = defaultRssMediaType;
            slideShow.creationDate = DateTime.Now;

            return slideShow;
        }

        public static SlideShow LoadFromReader(SqlDataReader reader)
        {
            SlideShow slideShow = new SlideShow();

            slideShow.puid = new Puid(reader.GetInt32(1), reader.GetInt32(0));
            slideShow.guid = reader.GetGuid(2);
            slideShow.name = reader.IsDBNull(3) ? null : reader.GetString(3);
            slideShow.friendlyName = reader.IsDBNull(4) ? null : reader.GetString(4);
            slideShow.friendlyNameVerified = false;
            slideShow.description = reader.IsDBNull(5) ? null : reader.GetString(5);
            slideShow.pmTimeZone = (PMTimeZone)reader.GetInt32(6);
            slideShow.pin = reader.IsDBNull(7) ? null : reader.GetString(7);
            slideShow.defaultImageSize = (SlideShowImageSize)reader.GetInt32(8);
            slideShow.defaultRssMediaType = (SlideShowRssMediaType)reader.GetInt32(9);

            slideShow.compileAll = reader.GetBoolean(10);
            slideShow.compileMaxCount = reader.GetInt32(11);
            slideShow.compileMinCount = reader.GetInt32(12);
            slideShow.compileOrder = (CompileOrder)reader.GetInt32(13);
            slideShow.compileInclude = (CompileInclude)reader.GetInt32(14);
            slideShow.compileAge = reader.GetInt32(15);

            if (!reader.IsDBNull(16))
            {
                DateTime compiledItemsCacheDate = reader.GetDateTime(17);
                DateTime compiledItemsCacheInvalidDate = reader.GetDateTime(18);
                if (DateTime.Now < compiledItemsCacheInvalidDate)
                {
                    slideShow.compiledItemsCache = ListItem.RootLoadListFromString(reader.GetString(16));
                    slideShow.compiledItemsCacheDate = compiledItemsCacheDate;
                    slideShow.compiledItemsCacheInvalidDate = compiledItemsCacheInvalidDate;
                    slideShow.compiledItemsCacheDirty = false;
                }
                else
                {
                    slideShow.compiledItemsCacheDirty = true;
                }

                
            }

            slideShow.creationDate = reader.GetDateTime(19);

            return slideShow;
        }        

        public static List<SlideShow> LoadSlideShowsFromDb(Puid puid)
        {
            List<SlideShow> slideShows = new List<SlideShow>();
            

            string sql = "" +
                    "select PuidHigh, PuidLow, SlideShowGuid, Name, FriendlyName, Description, TimeZone, Pin, DefaultImageSize, DefaultRssMediaType, " +
                    "   CompileAll, CompileMaxCount, CompileMinCount, CompileOrder, CompileInclude, CompileAge, " +
                    "   CompiledItemsCache, CompiledItemsCacheDate, CompiledItemsCacheInvalidDate, CreationDate " +
                    "from SlideShows " +
                    "where " +
                    "   PuidHash = @PuidHash and PuidHigh = @PuidHigh and PuidLow = @PuidLow";

            using (PhotoMixQuery query = new PhotoMixQuery(sql, CommandType.Text))
            {                
                query.Parameters.Add("@PuidHigh", SqlDbType.Int).Value = puid.PuidHigh;
                query.Parameters.Add("@PuidLow", SqlDbType.Int).Value = puid.PuidLow;
                query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puid.GetHashCode();

                while (query.Reader.Read())
                {
                    slideShows.Add(SlideShow.LoadFromReader(query.Reader));
                }
            }

            return slideShows;
        }

        

        public static List<SlideShowInfo> GetSlideShowInfoList(Puid puid)
        {
            List<SlideShowInfo> result = new List<SlideShowInfo>();

            string sql = "" +
                    "select SlideShowGuid, Name " +
                    "   from SlideShows " +
                    "where " +
                    "   PuidHash = @PuidHash and PuidLow = @PuidLow and PuidHigh = @PuidHigh";
            using (PhotoMixQuery query = new PhotoMixQuery(sql, CommandType.Text))
            {
                query.Parameters.Add("@PuidHigh", SqlDbType.Int).Value = puid.PuidHigh;
                query.Parameters.Add("@PuidLow", SqlDbType.Int).Value = puid.PuidLow;
                query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puid.GetHashCode();

                while (query.Reader.Read())
                {
                    SlideShowInfo slideShowInfo = new SlideShowInfo();

                    slideShowInfo.slideShowGuid = query.Reader.GetGuid(0);
                    slideShowInfo.name = query.Reader.GetString(1);

                    result.Add(slideShowInfo);
                }

                return result;
            }			
        }
       
        public static SlideShow LoadFromDb(Puid puid, Guid guid)
        {
            return LoadFromDb(puid, 0, guid);
        }

        public static SlideShow LoadFromDb(int puidHash, Guid guid)
        {
            return LoadFromDb(null, puidHash, guid);
        }

        public static SlideShow CloneToPuid(string id, Puid newPuid)
        {
            Puid oldPuid;
            Guid oldGuid;
            Guid newGuid;

            SlideShow slideShow = LoadFromDb(id);

            oldPuid = slideShow.puid;
            oldGuid = slideShow.guid;
            newGuid = Guid.NewGuid();

            // Clone the slide show

            slideShow.puid = newPuid;
            slideShow.guid = newGuid;
            slideShow.dataDirty = true;
            slideShow.ClearCompiledItemsCache();
            slideShow.SaveToDb();

            // Clone the channels            
            List<Channel> channels = Channel.CloneChannelsFromDbToPuid(oldPuid, oldGuid, newPuid, newGuid);

            return slideShow;
        }

        public static SlideShow LoadFromDb(string id)
        {
            int puidHash;

            Guid slideShowGuid = SlideShow.LookupId(id, out puidHash);

            return LoadFromDb(puidHash, slideShowGuid);
        }

        private static SlideShow LoadFromDb(Puid puid, int puidHash, Guid guid)
        {                       
            string sql;
            if (puid == null)
            {
                sql = "" +
                    "select PuidHigh, PuidLow, SlideShowGuid, Name, FriendlyName, Description, TimeZone, Pin, DefaultImageSize, DefaultRssMediaType, " +
                    "   CompileAll, CompileMaxCount, CompileMinCount, CompileOrder, CompileInclude, CompileAge, " +
                    "   CompiledItemsCache, CompiledItemsCacheDate, CompiledItemsCacheInvalidDate, CreationDate " +
                    "from SlideShows " +
                    "where " +
                    "   PuidHash = @PuidHash and SlideShowGuid = @SlideShowGuid";
            }
            else
            {
                sql = "" +
                    "select PuidHigh, PuidLow, SlideShowGuid, Name, FriendlyName, Description, TimeZone, Pin, DefaultImageSize, DefaultRssMediaType, " +
                    "   CompileAll, CompileMaxCount, CompileMinCount, CompileOrder, CompileInclude, CompileAge, " +
                    "   CompiledItemsCache, CompiledItemsCacheDate, CompiledItemsCacheInvalidDate, CreationDate " +
                    "from SlideShows " +
                    "where " +
                    "   PuidHash = @PuidHash and PuidHigh = @PuidHigh and PuidLow = @PuidLow and SlideShowGuid = @SlideShowGuid";
            }

            using (PhotoMixQuery query = new PhotoMixQuery(sql, CommandType.Text))
            {

                if (puid != null)
                {
                    query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puid.GetHashCode();
                    query.Parameters.Add("@PuidHigh", SqlDbType.Int).Value = puid.PuidHigh;
                    query.Parameters.Add("@PuidLow", SqlDbType.Int).Value = puid.PuidLow;
                }
                else
                {
                    query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puidHash;
                }
                query.Parameters.Add("@SlideShowGuid", SqlDbType.UniqueIdentifier).Value = guid;


                if (query.Reader.Read())
                    return SlideShow.LoadFromReader(query.Reader);
                else
                    return null;
            }
        }

        //
        // Function to get the URL to the RSS feed.  This has a no knowledge of
        // the web site, but is the most convenient place to put this
        //
        public static SlideShow LoadFromFriendlyName(string friendlyName)
        {
            Guid slideShowGuid;
            int slideShowHash;

            slideShowGuid = FriendlyName.LookupFriendlyName(friendlyName, out slideShowHash);

            return SlideShow.LoadFromDb(slideShowHash, slideShowGuid);
        }

        static public Guid LookupId(string id, out int slideShowHash)
        {
            string[] values = id.Split('!');

            Guid slideShowGuid = Base32.Base32DecodeGuid(values[0]);
            slideShowHash = Convert.ToInt32(values[1], 16);

            return slideShowGuid;
        }

        static public string GenerateId(Guid slideShowGuid, int hash)
        {
            return Base32.Base32EncodeGuid(slideShowGuid) + "!" + hash.ToString("X");
        }

        public string GenerateId()
        {
            return SlideShow.GenerateId(this.guid, this.puid.GetHashCode());
        }

        public static SlideShow LoadFromId(string id)
        {
            int slideShowHash;
            Guid slideShowGuid = LookupId(id, out slideShowHash);

            return SlideShow.LoadFromDb(slideShowHash, slideShowGuid);
        }

        public static SlideShow LoadFromRssFeedQueryString(NameValueCollection queryString)
        {
            SlideShow slideShow = null;

            try
            {
                string id = queryString["id"];
                if (!String.IsNullOrEmpty(id))
                {
                    int slideShowHash;
                    Guid slideShowGuid = LookupId(id, out slideShowHash);

                    slideShow = SlideShow.LoadFromDb(slideShowHash, slideShowGuid);
                }                    
                else
                {
                    string friendlyName = queryString["fn"];

                    if (!String.IsNullOrEmpty(friendlyName))
                    {
                        slideShow = SlideShow.LoadFromFriendlyName(friendlyName);
                    }
                    else
                    {
                        Guid slideShowGuid = new Guid(queryString["sid"]);
                        int slideShowHash = Convert.ToInt32(queryString["h"]);

                        slideShow = SlideShow.LoadFromDb(slideShowHash, slideShowGuid);
                    }
                }
            }
            catch
            {

            }

            if (slideShow != null)
            {
                if (!String.IsNullOrEmpty(slideShow.Pin))
                {
                    if (queryString["pin"] != slideShow.Pin)
                        return null;
                }
            }

            return slideShow;
        }
        
        public string GetRssFeedUrl(bool useFriendlyName, string host, HttpRequest request)
        {
            string friendlyName = null;
            if (useFriendlyName)
            {
                friendlyName = this.GetFriendlyName();
            }

            string pin = this.Pin;

            string feedUrl = "";
            if (String.IsNullOrEmpty(friendlyName))
            {
                if (!String.IsNullOrEmpty(host))
                    feedUrl = "http://" + host;
                feedUrl += "/genrss/genrss.ashx?";
                feedUrl += "id=" + SlideShow.GenerateId(this.Guid, this.Puid.GetHashCode());
                if (!String.IsNullOrEmpty(this.Pin))
                    feedUrl += "&pin=" + this.Pin;
            }
            else
            {
                if (Config.GetSetting("DNSFriendlyName") == "1")
                {
                    feedUrl = friendlyName + "." + host;
                    if (!String.IsNullOrEmpty(this.Pin))
                        feedUrl = this.Pin + "." + feedUrl;
                    feedUrl = "http://" + feedUrl;
                }
                else
                {
                    feedUrl = "http://" + host + "/genrss/genrss.ashx?fn=" + friendlyName;
                    if (!String.IsNullOrEmpty(this.Pin))
                        feedUrl += "&pin=" + this.Pin;
                }
            }

            bool bypassCaches = (request.QueryString["bc"] == "1");
            if (bypassCaches)
                bypassCaches = (Config.GetSetting("AllowBypassCaches") == "1");
            if (bypassCaches)
                feedUrl = MiscUtil.AppendQueryStringParameter(feedUrl, "bc=1");

            return feedUrl;
        }

        //
        // Property getting/setting of values
        //
        public string Name
        {
            get 
            { 
                return this.name; 
            }
            set 
            {
                if (this.name != value)
                {
                    this.name = value;
                    this.dataDirty = true;
                }
            }
        }

        public SlideShowImageSize DefaultImageSize
        {
            get
            {
                return this.defaultImageSize;
            }
            set
            {
                if (this.defaultImageSize != value)
                {
                    this.defaultImageSize = value;
                    this.dataDirty = true;
                }
            }
        }

        public SlideShowRssMediaType DefaultRssMediaType
        {
            get
            {
                return this.defaultRssMediaType;
            }
            set
            {
                if (this.defaultRssMediaType != value)
                {
                    this.defaultRssMediaType = value;
                    this.dataDirty = true;
                }
            }
        }

        public bool SetFriendlyName(string newFriendlyName)
        {
            string oldFriendlyName = this.GetFriendlyName();

            // Update the friendly name in the database
            FriendlyName.UpdateFriendlyName(newFriendlyName, oldFriendlyName, this.Guid, puid.GetHashCode());

            // Store the friendly name
            this.friendlyName = newFriendlyName;

            // Don't assume that the friendly name is verified. Next time someone queries
            // it we will verify.  Because the "update" took place across several partions this
            // is important
            this.friendlyNameVerified = false;

            return true;
        }

        static public bool SwapFriendlyName(string newFriendlyName, SlideShow newSlideShow, SlideShow oldSlideShow)
        {
            string oldFriendlyName = newSlideShow.GetFriendlyName();

            FriendlyName.UpdateFriendlyName(newFriendlyName, oldFriendlyName, newSlideShow.Guid, newSlideShow.puid.GetHashCode(), oldSlideShow.Guid);

            newSlideShow.friendlyName = newFriendlyName;
            oldSlideShow.friendlyName = null;

            newSlideShow.friendlyNameVerified = false;
            oldSlideShow.friendlyNameVerified = false;

            return true;
        }

        public string GetFriendlyName()
        {
            if (this.friendlyName == null)
                return null;

            if (!this.friendlyNameVerified)
            {
                Guid fnGuid;
                int fnHash = 0;
                fnGuid = Msn.PhotoMix.SlideShow.FriendlyName.LookupFriendlyName(this.friendlyName, out fnHash);
                if (this.Guid != fnGuid || fnHash != puid.GetHashCode())
                {
                    this.friendlyName = null;
                }

                this.friendlyNameVerified = true;
            }

            return this.friendlyName;
        }

        

        public Puid Puid
        {
            get { return this.puid; }         
        }        

        public Guid Guid
        {
            get { return this.guid; }
        }

        public string Pin
        {
            get 
            { 
                return this.pin; 
            }
            set
            {
                if (value == "")
                    value = null;

                if (this.pin != value)
                {
                    this.dataDirty = true;
                    this.pin = value;
                }
            }
        }
        
        public string Description
        {
            get 
            { 
                return this.description; 
            }
            set 
            {
                if (value == "")
                    value = null;

                if (this.description != value)
                {
                    this.dataDirty = true;
                    this.description = value;
                }
            }
        }

        public PMTimeZone PMTimeZone
        {
            get { return this.pmTimeZone; }
            set
            {
                if (this.pmTimeZone != value)
                {
                    this.pmTimeZone = value;
                    this.dataDirty = true;
                    this.ClearCompiledItemsCache();
                }
            }
        }

        public bool IsFixedCount
        {
            get
            {
                return false;
            }
        }

        public bool CompileAll
        {
            get { return this.compileAll; }
            set
            {
                if (this.compileAll != value)
                {
                    this.compileAll = value;
                    this.dataDirty = true;
                    this.ClearCompiledItemsCache();
                }
            }
        }

        public int CompileMaxCount
        {
            get
            {
                return this.compileMaxCount;
            }
            set
            {
                if (this.compileMaxCount != value)
                {
                    this.compileMaxCount = value;
                    this.dataDirty = true;
                    this.ClearCompiledItemsCache();
                }
            }
        }

        public int CompileMinCount
        {
            get
            {
                return this.compileMinCount;
            }
            set
            {
                if (this.compileMinCount != value)
                {
                    this.compileMinCount = value;
                    this.dataDirty = true;
                    this.ClearCompiledItemsCache();
                }
            }
        }

        public CompileOrder CompileOrder
        {
            get
            {
                return this.compileOrder;
            }
            set
            {
                if (this.compileOrder != value)
                {
                    this.compileOrder = value;
                    this.dataDirty = true;
                    this.ClearCompiledItemsCache();
                }
            }
        }

        public CompileInclude CompileInclude
        {
            get
            {
                return this.compileInclude;
            }
            set
            {
                if (this.compileInclude != value)
                {
                    this.compileInclude = value;
                    this.dataDirty = true;
                    this.ClearCompiledItemsCache();
                }
            }
        }

        public int CompileAge
        {
            get
            {
                return this.compileAge;
            }
            set
            {
                if (this.compileAge != value)
                {
                    this.compileAge = value;
                    this.dataDirty = true;
                    this.ClearCompiledItemsCache();
                }
            }
        }


        public List<ListItem> CompiledItemsCache
        {
            get { return this.compiledItemsCache; }
        }

        public DateTime CompiledItemsCacheDate
        {
            get { return this.compiledItemsCacheDate; }
        }

        public DateTime CompiledItemsCacheInvalidDate
        {
            get { return this.compiledItemsCacheInvalidDate; }
        }

        public bool CompiledItemsCacheDirty
        {
            get { return this.compiledItemsCacheDirty; }
        }

        //
        // Methods to deal with the channels
        //
        public List<Channel> GetChannels()
        {
            if (this.channels == null)
                this.channels = Channel.LoadChannelsFromDb(this.puid, this.guid);

            return this.channels;
        }

        public void DeleteChannel(Channel targetChannel)
        {
            targetChannel.DeleteFromDb();
            channels.Remove(targetChannel);

            this.ClearCompiledItemsCache();

            this.SaveCompiledItemsCacheToDb();
        }

        private Channel FindChannel(Guid channelGuid)
        {
            List<Channel> channels = this.GetChannels();

            foreach (Channel channel in channels)
            {
                if (channel.ChannelGuid == channelGuid)
                {
                    return channel;
                }
            }

            return null;
        }

        public void DeleteChannel(Guid channelGuid)
        {            
            // Find the channel to delete
            Channel targetChannel = FindChannel(channelGuid);
            
            if (targetChannel != null)
                DeleteChannel(targetChannel);
        }

        public void AddChannel(Channel channel)
        {
            // Get the channels (this will make sure they are loaded in memory)
            List<Channel> channels = this.GetChannels();

            // Update the display order of the channel so that it is at the end
            // of the channel list
            channel.InitDisplayOrder(channels.Count);

            // Save the channel
            channel.SaveToDb();

            // Add the channel
            channels.Add(channel);
            this.ClearCompiledItemsCache();
            
            // Save the slide show
            this.SaveCompiledItemsCacheToDb();
        }

        public void UpdateDisplayOrderUp(Guid channelGuid)
        {
            int currentPosition = 0;            

            // Find the channel
            foreach (Channel channel in this.GetChannels())
            {
                if (channel.ChannelGuid == channelGuid)
                {
                    if (currentPosition > 0)
                    {
                        UpdateDisplayOrder(channelGuid, currentPosition - 1);
                        break;
                    }
                }
                currentPosition++;
            }
        }

        public void UpdateDisplayOrderDown(Guid channelGuid)
        {
            int currentPosition = 0;

            // Find the channel
            foreach (Channel channel in this.GetChannels())
            {
                if (channel.ChannelGuid == channelGuid)
                {
                    if (currentPosition < this.channels.Count - 1)
                    {
                        UpdateDisplayOrder(channelGuid, currentPosition + 1);
                        break;
                    }
                }
                currentPosition++;
            }
        }

        public void UpdateDisplayOrder(Guid channelGuid, int newPosition)
        {
            Channel targetChannel = null;

            int index = 0;

            foreach (Channel channel in this.GetChannels())
            {

                // Skip the position we are inserting into
                if (index == newPosition)
                    index++;

                // Skip ourself.  We'll update this position when we are done
                if (channel.ChannelGuid == channelGuid)
                {
                    targetChannel = channel;
                    continue;
                }

                // If the current index is less than the position we are inserting into
                channel.UpdateDisplayOrder(index);

                index++;
            }

            // If the new position is greater than the last index, than we must be wanting to add it to the end
            if (newPosition > index)
                newPosition = index;

            // If the target channel is null return
            if (targetChannel == null)
                return;
            
            // Update the target channel to the new position
            targetChannel.UpdateDisplayOrder(newPosition);

            // Move the channel to the new position in our list too
            this.channels.Remove(targetChannel);
            this.channels.Insert(newPosition, targetChannel);

            // Clear the compile if we are order dependent
            if (this.CompileOrder == CompileOrder.Listed)
            {
                this.ClearCompiledItemsCache();
                this.SaveCompiledItemsCacheToDb();
            }
        }        

        //
        // LogRss
        //
        public void LogRss(string trackId)
        {
            string sql = "" +
                "if not exists (select SlideShowGuid from SlideShowViews where PuidHash = @PuidHash and SlideShowGuid = @SlideShowGuid and TrackId = @TrackId)" +
                "    insert into SlideShowViews (" +
                "       PuidHash, SlideShowGuid, Count, TrackId, LastAccessDate " +
                "    )" +
                "    values (" +
                "       @PuidHash, @SlideShowGuid, 1, @TrackId, GetDate() " +
                "    )" +
                "else" +
                "    update SlideShowViews" +
                "    set " +
                "       Count = Count + 1, LastAccessDate = GetDate() " +
                "    where " +
                "       PuidHash = @PuidHash and SlideShowGuid = @SlideShowGuid and TrackId = @TrackId " +
                "update SlideShows " +
                "   set " +
                "       Count = Count + 1, LastAccessDate = GetDate() " +
                "    where " +
                "       PuidHash = @PuidHash and SlideShowGuid = @SlideShowGuid";

            using (PhotoMixQuery query = new PhotoMixQuery(sql, CommandType.Text))
            {
                query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = this.puid.GetHashCode();
                query.Parameters.Add("@SlideShowGuid", SqlDbType.UniqueIdentifier).Value = this.guid;
                query.Parameters.Add("@TrackId", SqlDbType.NVarChar).Value = (trackId == null ? "" : trackId);
                
                query.Execute();
            }
        }
       
        //
        // GenerateRss
        //
        // Generates Rss for the compiled items in the slideshow. This generates the Rss header
        // for our output feed, gets the compiled items, and asks each item to in turn generate
        // its output Rss.
        //        
        public void GenerateRss(StringBuilder sb)
        {
            GenerateRss(sb, SlideShowRssMediaType.MediaRss, SlideShowImageSize.SizeUnknown, false, false, false, Guid.Empty);
        }

        public void GenerateRss(StringBuilder sb, SlideShowRssMediaType rssType, SlideShowImageSize targetImageSize, bool insertSizes, bool bypassCaches, bool debug)
        {
            GenerateRss(sb, rssType, targetImageSize, insertSizes, bypassCaches, debug, Guid.Empty);
        }

        public void GenerateRss(StringBuilder sb, SlideShowRssMediaType rssType, SlideShowImageSize targetImageSize, bool insertSizes, bool bypassCaches, bool debug, Guid channelGuid)
        {
            // Calculate the date for a consistent snapshot
            DateTime dateContext = DateTime.Now;

            // Get our compiled items (this will compile as necessary).  Must be done first
            // because dates are referenced in calcuating the TTL below.
            List<ListItem> listItems;
            string name;
            string type;
            string guidString;
            DateTime creationDate;
            DateTime compiledDate;
            DateTime compiledInvalidDate;
            string description;
            if (channelGuid == Guid.Empty)
            {
                // Get the compiled items (and settings) from the slide show
                listItems = this.GetCompiledItems(null, dateContext, bypassCaches, null);
                name = this.name;
                type = "SlideShow";
                guidString = this.guid.ToString();
                creationDate = this.creationDate;
                compiledDate = this.compiledItemsCacheDate;
                compiledInvalidDate = this.compiledItemsCacheInvalidDate;
                description = this.description;
            }
            else
            {
                // Get the compiled items (and settings) from the channel
                Channel channel = this.FindChannel(channelGuid);
                name = channel.Name;
                type = "Channel " + channel.Type.ToString();
                guidString = channelGuid.ToString();
                creationDate = channel.CreationDate;
                compiledDate = channel.GetCompiledItemsDate();
                compiledInvalidDate = channel.GetCompiledItemsInvalidDate();                
                description = null;
                listItems = channel.GetCompiledItems(null, dateContext, bypassCaches, null);
            }

            // Use the default media type if none was passed in
            if (rssType == SlideShowRssMediaType.Unknown)
            {
                rssType = this.defaultRssMediaType;
            }

            // Use the default size if none was passed in
            if (targetImageSize == SlideShowImageSize.SizeUnknown)
            {
                targetImageSize = this.defaultImageSize;
            }

            // Generate the header
            sb.Append("<?xml version=\"1.0\" ?>");           
            string xmlns = "";
            if (rssType != SlideShowRssMediaType.EnclosureRss)
                xmlns = "xmlns:media=\"http://search.yahoo.com/mrss/\""; 
            sb.Append("<rss version=\"2.0\" " + xmlns + ">");
            sb.Append("<channel>");
            
            // Calculate the time to live for the feed
            if (compiledInvalidDate != DateTime.MaxValue)
            {
                TimeSpan ts = compiledInvalidDate.Subtract(dateContext);
                int ttl = (int)ts.TotalMinutes;
                sb.Append("<ttl>" + ttl.ToString() + "</ttl>");
            }

            string hostSite = "http://" + Config.GetSetting("HostSite");                       

            // Render the title, link for the feed
            sb.Append("<title>" + HttpUtility.HtmlEncode(name) + "</title>");
            sb.Append("<link>" + hostSite + "</link>");
            sb.Append("<generator>" + hostSite + "</generator>");
            
            // Write out the last build date            
            sb.Append("<lastBuildDate>" + TimeUtil.DateToRSSString(compiledDate) + "</lastBuildDate>");

            // Write out the publish date
            sb.Append("<pubDate>" + TimeUtil.DateToRSSString(dateContext) + "</pubDate>");

            // Write out the description header
            sb.Append("<description>");

            // Display our description if we have one
            if (!String.IsNullOrEmpty(description))
                sb.Append(HttpUtility.HtmlEncode(description));

            // For the debug version, add a bunch of stuff to the description
            if (debug)
            {
                sb.Append(HttpUtility.HtmlEncode("<br/>"));
                sb.Append(HttpUtility.HtmlEncode("RSS Feed for " + type + "<br/>"));
                sb.Append(HttpUtility.HtmlEncode("Guid " + guidString + "<br/>"));

                sb.Append(HttpUtility.HtmlEncode("Generated on " + dateContext.ToString() + "<br/>"));
                sb.Append(HttpUtility.HtmlEncode("Feed created on " + creationDate.ToString() + "<br/>"));
                sb.Append(HttpUtility.HtmlEncode("Compiled on " + compiledDate.ToString() + "<br/>"));
                sb.Append(HttpUtility.HtmlEncode("Compile invalid on " + compiledInvalidDate.ToString() + "<br/>"));
            }

            // Close out the description
            sb.Append("</description>");

            if (listItems.Count == 0)
            {
                // There are no items, so output one item pointing to the blank image
                SlideShowItem slideShowItem = new SlideShowItem(
                        DateTime.MaxValue,
                        dateContext,
                        "Displays",
                        Config.GetSetting("CompiledImageBaseUrl") + "BlankFeedItem.ashx?<SIZE>" + (bypassCaches ? "&bc=1" : ""),
                        true);

                slideShowItem.GenerateRss(sb, rssType, targetImageSize, insertSizes, debug);
            }
            else
            {
                // Compiled items are a list of class CompiledItem.  Ask each one to output its rss
                foreach (ListItem listItem in listItems)
                {
                    SlideShowItem slideShowItem = (SlideShowItem)listItem;
                    slideShowItem.GenerateRss(sb, rssType, targetImageSize, insertSizes, debug);
                }
            }

            // Generate the footer
            sb.Append("</channel>");
            sb.Append("</rss>");
        }      

        //
        // Compile
        //
        // Compiles the slide show into the final RSS feed.  
        //
        // GuidContext is a list of Guids of slideshows currently in the compile stack.  It is used to prevent 
        // circular references during the compilation process. This could happen if a slide show contains a 
        // reference channel pointing to itself, or through several layers.
        //
        // DateContext is the date of compile used for all default image and item generation pubDate
        // when no date is present
        //
        // CompileState is a hashtable that can be used by steps in the compile to store
        // stateful data that could potentially be re-used when subsequent items in a channel
        // might be referenced a second time.
        //
        // The typical compile will start by passing in null for these values.        
        //
        public void Compile()
        {            
            Compile(null, DateTime.MinValue, false, null);
        }

        public void Compile(List<Guid> guidContext, DateTime dateContext, bool bypassCaches, Hashtable compileState)
        {
            // The date time that this compilation is valid until
            DateTime compiledItemsCacheInvalidDate = DateTime.MaxValue;

            // These are the working list of target items
            List<ListItem> workingItems = new List<ListItem>();

            // This is the final list of compiled items
            //List<ListItem> compiledItemsCache = new List<ListItem>();
            
            // Create the guid context if ncessary
            if (guidContext == null)         
                guidContext = new List<Guid>();

            // Get the current date if we don't have a valid date
            if (dateContext == DateTime.MinValue)
                dateContext = DateTime.Now;

            // Create a compile state if none was passed in
            if (compileState == null)
                compileState = new Hashtable();

            // Skip the compile if the guid context contains this slideshow.  It must be a circular reference
            if (!guidContext.Contains(this.guid))
            {
                // Add our own guid for compile of our channels
                guidContext.Add(this.guid);
                
                // Compile the channels
                foreach (Channel channel in this.GetChannels())
                {
                    // Include the channel in the compile if it is active
                    if (channel.IsActive(dateContext, this.pmTimeZone))
                    {
                        // Get the compiled items from the channel
                        List<ListItem> channelItems = channel.GetCompiledItems(guidContext, dateContext, bypassCaches, compileState);

                        // Copy the items into the slide show list
                        workingItems.AddRange(channelItems);

                        // The slide show compile is only as valid as the channel compile
                        DateTime channelCompiledItemsCacheInvalidDate = channel.GetCompiledItemsInvalidDate();
                        if ((channelCompiledItemsCacheInvalidDate < compiledItemsCacheInvalidDate) &&
                            (channelCompiledItemsCacheInvalidDate >= channel.GetCompiledItemsDate()))  //Safety in case the channel hasn't updated it's cache invalid date
                            compiledItemsCacheInvalidDate = channelCompiledItemsCacheInvalidDate;
                    }

                    // The slide show should be recompiled when a channel becomes active or inactive
                    DateTime nextActiveDateChange = channel.GetActiveChangeDate(dateContext, this.pmTimeZone);
                    if (nextActiveDateChange < compiledItemsCacheInvalidDate)
                        compiledItemsCacheInvalidDate = nextActiveDateChange;
                }

                // Now that we have all the items in the working list, apply the slide show compile rules
                CompileUtil.PruneList((List<ListItem>)workingItems, this.compileAll, this.compileMaxCount, this.compileMinCount, this.compileOrder, this.compileInclude, this.compileAge, dateContext, this.creationDate);
               
                // Remove our guid from the context
                guidContext.Remove(this.guid);
            }            

            // Store the compiled items
            this.compiledItemsCache = workingItems;
            this.compiledItemsCacheDate = dateContext;
            this.compiledItemsCacheInvalidDate = compiledItemsCacheInvalidDate;
            this.compiledItemsCacheDirty = true;            
        }

        public void ClearCompiledItemsCache()
        {
            if (this.compiledItemsCache == null)
                return;

            this.compiledItemsCache = null;
            this.compiledItemsCacheDirty = true;
        }

        public List<ListItem> GetCompiledItems(List<Guid> guidContext, DateTime dateContext, bool bypassCaches, Hashtable compileState)
        {
            if (this.compiledItemsCache == null || 
                dateContext > this.compiledItemsCacheInvalidDate ||
                bypassCaches)
            {
                this.Compile(guidContext, dateContext, bypassCaches, compileState);
            }

            return this.compiledItemsCache;
        }

        public void SaveCompiledItemsCacheToDb()
        {
            if (!this.compiledItemsCacheDirty)
                return;            

            string sql = "" +
                "update SlideShows " +
                "set CompiledItemsCache = @CompiledItemsCache, CompiledItemsCacheDate = @CompiledItemsCacheDate, CompiledItemsCacheInvalidDate = @CompiledItemsCacheInvalidDate, ModifiedDate = GetDate() " +
                "where PuidHash = @PuidHash and PuidHigh = @PuidHigh and PuidLow = @PuidLow and SlideShowGuid = @SlideShowGuid";

            using (PhotoMixQuery query = new PhotoMixQuery(sql, CommandType.Text))
            {
                query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puid.GetHashCode();
                query.Parameters.Add("@PuidHigh", SqlDbType.Int).Value = puid.PuidHigh;
                query.Parameters.Add("@PuidLow", SqlDbType.Int).Value = puid.PuidLow;
                query.Parameters.Add("@SlideShowGuid", SqlDbType.UniqueIdentifier).Value = this.guid;

                if (this.compiledItemsCache == null)
                {
                    query.Parameters.Add("@CompiledItemsCache", SqlDbType.Text).Value = DBNull.Value;
                    query.Parameters.Add("@CompiledItemsCacheDate", SqlDbType.DateTime).Value = DBNull.Value;
                    query.Parameters.Add("@CompiledItemsCacheInvalidDate", SqlDbType.DateTime).Value = DBNull.Value;
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    ListItem.RootSaveListToString(sb, this.compiledItemsCache);

                    query.Parameters.Add("@CompiledItemsCache", SqlDbType.Text).Value = sb.ToString();
                    query.Parameters.Add("@CompiledItemsCacheDate", SqlDbType.DateTime).Value = this.compiledItemsCacheDate;
                    query.Parameters.Add("@CompiledItemsCacheInvalidDate", SqlDbType.DateTime).Value = this.compiledItemsCacheInvalidDate;
                }

                query.Execute();
            }

            this.compiledItemsCacheDirty = false;
        }

        public void DeleteFromDb()
        {
            // If we have a friendly name, delete it
            if (this.friendlyName != null)
            {
                Msn.PhotoMix.SlideShow.FriendlyName.DeleteFriendlyName(this.friendlyName, this.Guid, this.puid.GetHashCode());
            }

            // Now delete the slideshow and channels
            string sql = "" +
                "delete from Channels " +
                "where PuidHash = @PuidHash and PuidHigh = @PuidHigh and PuidLow = @PuidLow and SlideShowGuid = @SlideShowGuid " +
                "delete from SlideShows " +
                "where PuidHash = @PuidHash and PuidHigh = @PuidHigh and PuidLow = @PuidLow and SlideShowGuid = @SlideShowGuid ";

            using (PhotoMixQuery query = new PhotoMixQuery(sql, CommandType.Text))
            {
                query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puid.GetHashCode();
                query.Parameters.Add("@PuidHigh", SqlDbType.Int).Value = puid.PuidHigh;
                query.Parameters.Add("@PuidLow", SqlDbType.Int).Value = puid.PuidLow;
                query.Parameters.Add("@SlideShowGuid", SqlDbType.UniqueIdentifier).Value = this.guid;

                query.Execute();
            }
        }

        public void SaveChannelsToDb()
        {
            if (this.channels != null)
            {
                foreach (Channel channel in this.channels)
                {
                    channel.SaveToDb();
                }
            }
        }

        public void SaveToDb()
        {
            if (!this.compiledItemsCacheDirty && !this.dataDirty)
                return;

            if (this.compiledItemsCacheDirty && !this.dataDirty)
            {
                this.SaveCompiledItemsCacheToDb();

                return;
            }

            string sql = "" +
                "if not exists (select SlideShowGuid from SlideShows where PuidHigh = @PuidHigh and PuidLow = @PuidLow and SlideShowGuid = @SlideShowGuid)" +
                "    insert into SlideShows (" +
                "       PuidHash, PuidHigh, PuidLow, SlideShowGuid, Name, FriendlyName, Description, TimeZone, Pin, DefaultImageSize, DefaultRssMediaType, " +
                "       CompileAll, CompileMaxCount, CompileMinCount, CompileOrder, CompileInclude, CompileAge, " +
                "       CompiledItemsCache, CompiledItemsCacheDate, CompiledItemsCacheInvalidDate, CreationDate, ModifiedDate " +
                "    )" +
                "    values (" +
                "       @PuidHash, @PuidHigh, @PuidLow, @SlideShowGuid, @Name, @FriendlyName, @Description, @TimeZone, @Pin, @DefaultImageSize, @DefaultRssMediaType, " +
                "       @CompileAll, @CompileMaxCount, @CompileMinCount, @CompileOrder, @CompileInclude, @CompileAge, " +
                "       @CompiledItemsCache, @CompiledItemsCacheDate, @CompiledItemsCacheInvalidDate, GetDate(), GetDate() " +
                "    )" +
                "else" +
		        "    update SlideShows" +
		        "    set " +
                "       Name = @Name,  FriendlyName = @FriendlyName, Description = @Description, TimeZone = @TimeZone, Pin = @Pin, DefaultImageSize = @DefaultImageSize, DefaultRssMediaType = @DefaultRssMediaType, " +
                "       CompileAll = @CompileAll, CompileMaxCount = @CompileMaxCount, CompileMinCount = @CompileMinCount, CompileOrder = @CompileOrder, CompileInclude = @CompileInclude, CompileAge = @CompileAge, " +
                "       CompiledItemsCache = Case @CompiledItemsCacheDirty when 0 then CompiledItemsCache when 1 then @CompiledItemsCache end, " +
                "       CompiledItemsCacheDate  = Case @CompiledItemsCacheDirty when 0 then CompiledItemsCacheDate when 1 then @CompiledItemsCacheDate end, " +
                "       CompiledItemsCacheInvalidDate  = Case @CompiledItemsCacheDirty when 0 then CompiledItemsCacheInvalidDate when 1 then @CompiledItemsCacheInvalidDate end, " +
                "       ModifiedDate = GetDate() " +                
                "    where " +
                "       PuidHash = @PuidHash and PuidHigh = @PuidHigh and PuidLow = @PuidLow and SlideShowGuid = @SlideShowGuid";

            using (PhotoMixQuery query = new PhotoMixQuery(sql, CommandType.Text))
            {
                query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puid.GetHashCode();
                query.Parameters.Add("@PuidHigh", SqlDbType.Int).Value = puid.PuidHigh;
                query.Parameters.Add("@PuidLow", SqlDbType.Int).Value = puid.PuidLow;
                query.Parameters.Add("@SlideShowGuid", SqlDbType.UniqueIdentifier).Value = this.guid;                                    
                query.Parameters.Add("@Name", SqlDbType.NVarChar).Value = this.name;
                if (this.friendlyName == null)
                    query.Parameters.Add("@FriendlyName", SqlDbType.NVarChar).Value = DBNull.Value;
                else
                    query.Parameters.Add("@FriendlyName", SqlDbType.NVarChar).Value = this.friendlyName;

                if (this.description == null)
                    query.Parameters.Add("@Description", SqlDbType.NVarChar).Value = DBNull.Value;
                else
                    query.Parameters.Add("@Description", SqlDbType.NVarChar).Value = this.description;
                query.Parameters.Add("@TimeZone", SqlDbType.Int).Value = (int)this.pmTimeZone;
                if (this.pin == null)
                    query.Parameters.Add("@Pin", SqlDbType.NVarChar).Value = DBNull.Value;
                else
                    query.Parameters.Add("@Pin", SqlDbType.NVarChar).Value = this.pin;
                query.Parameters.Add("@DefaultImageSize", SqlDbType.Int).Value = (int)this.defaultImageSize;
                query.Parameters.Add("@DefaultRssMediaType", SqlDbType.Int).Value = (int)this.defaultRssMediaType;

                query.Parameters.Add("@CompileAll", SqlDbType.Bit).Value = this.CompileAll;
                query.Parameters.Add("@CompileMaxCount", SqlDbType.Int).Value = this.CompileMaxCount;
                query.Parameters.Add("@CompileMinCount", SqlDbType.Int).Value = this.CompileMinCount;
                query.Parameters.Add("@CompileOrder", SqlDbType.Int).Value = this.CompileOrder;
                query.Parameters.Add("@CompileInclude", SqlDbType.Int).Value = this.CompileInclude;
                query.Parameters.Add("@CompileAge", SqlDbType.Int).Value = this.CompileAge;
                query.Parameters.Add("@CompiledItemsCacheDirty", SqlDbType.Int).Value = (this.CompiledItemsCacheDirty ? 1 : 0);
                if (this.compiledItemsCache == null || this.CompiledItemsCacheDirty == false)
                {
                    query.Parameters.Add("@CompiledItemsCache", SqlDbType.Text).Value = DBNull.Value;
                    query.Parameters.Add("@CompiledItemsCacheDate", SqlDbType.DateTime).Value = DBNull.Value;
                    query.Parameters.Add("@CompiledItemsCacheInvalidDate", SqlDbType.DateTime).Value = DBNull.Value;
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    ListItem.RootSaveListToString(sb, this.compiledItemsCache);

                    query.Parameters.Add("@CompiledItemsCache", SqlDbType.Text).Value = sb.ToString();
                    query.Parameters.Add("@CompiledItemsCacheDate", SqlDbType.DateTime).Value = this.compiledItemsCacheDate;
                    query.Parameters.Add("@CompiledItemsCacheInvalidDate", SqlDbType.DateTime).Value = this.compiledItemsCacheInvalidDate;
                }


                query.Execute();
            }

            this.dataDirty = false;
            this.compiledItemsCacheDirty = false;
        }


    }
}
