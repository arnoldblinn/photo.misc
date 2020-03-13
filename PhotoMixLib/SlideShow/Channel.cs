using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

using System.Collections;

using System.Drawing;
using System.Web;
using System.Xml;

using Msn.Framework;
using Msn.PhotoMix.Passport;
using Msn.PhotoMix.Util;


namespace Msn.PhotoMix.SlideShow
{ 
    //
    // Channel
    //
    // A channel in a slide show is a source of photos for display.  These are the current
    // set of channel types.  When adding a type, you need to modify the code that loads them
    // from marshalled data in the database.
    //
    public enum ChannelType
    {
        Unknown = 0,
        Facebook = 1,
        Unused = 2,
        Space = 3,
        Reference = 4,
        Static = 5,
        Weather = 6,
        WebPage = 7,
        Rss = 8,
        USTraffic = 9,
        Flickr = 10,
        SmugMug = 11,
        FixedReference = 12,

        Count = 13,
        Max = 12
    }

    //
    // ActiveDay
    //
    // These flags control days of the week as a flag vector
    //
    [Flags]
    public enum ActiveDay : byte
    {
        None = 0,
        Mon = 1,
        Tue = 2,
        Wed = 4,
        Thu = 8,
        Fri = 16,
        Sat = 32,
        Sun = 64,
        All = Mon | Tue | Wed | Thu | Fri | Sat | Sun
    }        

    public class Channel : ICompile
    {        
        private Puid puid;

        // Guid for the slide show this channel is in, and the guid for the channel itself
        private Guid slideShowGuid = Guid.Empty;
        private Guid channelGuid = Guid.Empty;

        // Type of the channel
        private ChannelType type = ChannelType.Unknown;

        // Name of the channel
        private string name = null;

        // Order for rendering (controlled by the SlideShow)
        private int channelOrder = 0;

        // Creation date of the channel (also used for ordering in ties...older come first)
        private DateTime creationDate;
        
        // Compilation control settings (for channels that need to be compiled)		
        private bool compileAll = true;
        private int compileMaxCount = 5;
        private int compileMinCount = 1;
        private CompileOrder compileOrder = CompileOrder.Listed;
        private CompileInclude compileInclude = CompileInclude.All;
        private int compileAge = 10;

        // Flags controlling the time the channel is active for participation in a slideshow
        private bool active = true;        
        private bool activeLimitTime = false;        
        private int activeStartMinutes = 0;
        private int activeEndMinutes = 1440;        
        private bool activeLimitDays = false;
        private ActiveDay activeDays = ActiveDay.All;
        private bool activeExpire = false;
        private DateTime activeExpireDate = DateTime.MaxValue;
                
        // Populated items
        private List<ListItem> items = null;        
        
        // Cached compile of a channel
        private List<ListItem> compiledItemsCache = null;
        private DateTime compiledItemsCacheDate = DateTime.MinValue;
        private DateTime compiledItemsCacheInvalidDate = DateTime.MaxValue;
        private bool compiledItemsCacheDirty = false;

        // These flags are used to tell the channel that we've changed the results 
        // of "ItemsNeedCompileState" 
        // 
        // This is needed if a subclass overrides the virtual method and 
        // decides to start returning "false" instead of "true" for this
        // these values
        // 
        // Why?
        //
        // Because there are instances where the channel class will have saved
        // compiled data in the database when it no longer needs to.  We need to 
        // allow the channel to "blank" this data. These values will trigger this.
        private bool changedItemsNeedCompileState = false;

        // Flag indicating if the channel state is dirty (e.g. it needs to be saved)
        private bool dataDirty = false;

        //
        // Constructors
        //
        public Channel(Puid puid, Guid slideShowGuid, Guid channelGuid, ChannelType type)
        {
            if (channelGuid == Guid.Empty)
                channelGuid = Guid.NewGuid();

            this.puid = puid;
            this.slideShowGuid = slideShowGuid;            
            this.channelGuid = channelGuid;
            this.type = type;
            this.creationDate = DateTime.Now;
        }

        public Channel(Puid puid, Guid slideShowGuid, ChannelType type)
        {
            this.puid = puid;
            this.slideShowGuid = slideShowGuid;
            this.channelGuid = Guid.NewGuid();
            this.type = type;
            this.creationDate = DateTime.Now;
        }


        //
        // Functions to deal with the display order of this channel.
        //
        // Note that if there are other instances of channels loaded into memory, this call
        // may invalidate their state data, as calling these functions may update the order of
        // other channels in our slide show.
        // 
        // This function largely exists for the slideshow to update an individual channel 
        // channel order in the database
        //        
        public void UpdateDisplayOrder(int index)
        {
            if (this.channelOrder == index)
                return;

            string sql = "update Channels " +
                "set ChannelOrder = @ChannelOrder " +
                "where PuidHash = @PuidHash and PuidHigh = @PuidHigh and PuidLow = @PuidLow and ChannelGuid = @ChannelGuid";
            using (PhotoMixQuery query = new PhotoMixQuery(sql, CommandType.Text))
            {
                query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puid.GetHashCode();
                query.Parameters.Add("@PuidHigh", SqlDbType.Int).Value = puid.PuidHigh;
                query.Parameters.Add("@PuidLow", SqlDbType.Int).Value = puid.PuidLow;
                query.Parameters.Add("@ChannelOrder", SqlDbType.Int).Value = index;
                query.Parameters.Add("@ChannelGuid", SqlDbType.UniqueIdentifier).Value = this.channelGuid;

                query.Execute();
            }
        }

        public void InitDisplayOrder(int position)
        {
            this.channelOrder = position;
        }

        public int GetDisplayOrder()
        {
            return this.channelOrder;
        }
        
        //
        // Methods and properties that deal with the active state of the channel
        //      

        // Gets the next day of the week from an ActiveDay structure
        private ActiveDay GetNextDay(ActiveDay currentDay)
        {            
            if (currentDay == ActiveDay.Mon) return ActiveDay.Tue;
            else if (currentDay == ActiveDay.Tue) return ActiveDay.Wed;
            else if (currentDay == ActiveDay.Wed) return ActiveDay.Thu;
            else if (currentDay == ActiveDay.Thu) return ActiveDay.Fri;
            else if (currentDay == ActiveDay.Fri) return ActiveDay.Sat;
            else if (currentDay == ActiveDay.Sat) return ActiveDay.Sun;
            else return ActiveDay.Mon;            
        }

        // Gets the current day of the week from a DayOfWeek type
        private ActiveDay GetCurrentDay(DayOfWeek dayOfWeek)
        {
            if (dayOfWeek == DayOfWeek.Monday) return ActiveDay.Mon;
            else if (dayOfWeek == DayOfWeek.Tuesday) return ActiveDay.Tue;
            else if (dayOfWeek == DayOfWeek.Wednesday) return ActiveDay.Wed;
            else if (dayOfWeek == DayOfWeek.Thursday) return ActiveDay.Thu;
            else if (dayOfWeek == DayOfWeek.Friday) return ActiveDay.Fri;
            else if (dayOfWeek == DayOfWeek.Saturday) return ActiveDay.Sat;
            else return ActiveDay.Sun;
        }

        // Counts the number of days till the active state has changed from current day
        private int CountDays(ActiveDay currentDay)
        {
            // Comments will be by example. A "-" represents a bit unused.
            //
            // Assume this.ActiveDays = Mon|Wed -> -0000101 and currentDay = Thu -> -0001000            

            // Create a value that can be shifted left for compare.  This means we 
            // load active days in a 16 bit value and repeat the pattern for shifting.
            //
            // For our example, activeDaysWrap = -0000101|0000101-
            UInt16 activeDaysWrap = (UInt16)(((UInt16)this.activeDays << (UInt16)8) | ((UInt16)this.activeDays << (UInt16)1));

            // Now shift the current state the same way to see if the current day is on or off            
            bool currentState = ((activeDaysWrap & ((UInt16)currentDay << (UInt16)8))!= 0);

            // Now simply shift the bit pattern of active days wrap one left at a time till 
            // the state value changes
            for (int i = 1; i < 7; i++)
            {
                activeDaysWrap = (UInt16)(activeDaysWrap >> (UInt16)1);

                bool newState = ((activeDaysWrap & ((UInt16)currentDay << (UInt16)8)) != 0);

                // The state value changed.  Return the count
                if (newState != currentState)
                    return i;
            }

            // A value of 7 means the value never changed (e.g. all days were on or off)
            return 7;
        }

        //
        // GetActiveChangeDate
        //
        // Gets the date/time (rounded down to the minute) for when the active state
        // will change.  A value of DateTime.MaxValue indicates the state will never change
        //        
        public DateTime GetActiveChangeDate(DateTime currentDateTime, PMTimeZone timeZone)
        {
            // Adjust the current date time into the timezone of the feed
            currentDateTime = TimeUtil.AdjustDateTime(currentDateTime, timeZone);

            // Find the next active change time in minutes
            int nextActiveChange = this.GetActiveChangeTTL(currentDateTime, timeZone);

            DateTime newDateTime = DateTime.MaxValue;
            if (nextActiveChange < Int32.MaxValue)
                newDateTime = currentDateTime.Add(new TimeSpan(0, nextActiveChange, 0));

            // If the feed expires, account for this
            if (this.activeExpire && this.activeExpireDate < newDateTime && this.activeExpireDate > currentDateTime)
            {
                newDateTime = this.activeExpireDate;
            }

            if (newDateTime == DateTime.MaxValue)
                return newDateTime;

            return newDateTime.Subtract(new TimeSpan(0, 0, 0, newDateTime.Second, newDateTime.Millisecond));
        }
        
        // 
        // GetActiveChangeTTL
        //
        // Passes back in TTL form the number of minutes till the next active change. A value
        // of Int32.MaxValue indicates the state will never change.
        //
        private int GetActiveChangeTTL(DateTime currentDateTime, PMTimeZone pmTimeZone)
        {
            // If we aren't active, or if we aren't limiting by time or days, we don't
            // actually change active state
            if (this.active == false ||
                (!this.activeLimitTime && 
                 (!this.activeLimitDays || (this.activeLimitDays && this.activeDays == ActiveDay.None))
                )
               )               
               return Int32.MaxValue;                       
            
            ActiveDay currentDay = GetCurrentDay(currentDateTime.DayOfWeek);

            int days = 7;
            if (activeLimitDays)
            {
                days = CountDays(currentDay);
            }

            int startMinutes = 0;
            int endMinutes = 1440;
            int currentMinutes = currentDateTime.Hour * 60 + currentDateTime.Minute;
            if (activeLimitTime)
            {
                startMinutes = this.activeStartMinutes;
                endMinutes = this.activeEndMinutes;
            }

            // All the days are on, and it is active all day long
            if (days == 7 && (startMinutes == 0 && endMinutes == 1440))
            {
                return Int32.MaxValue;
            }
            // Some of the days are off and on, but it is on for all day on those days
            else if (days != 7 && (startMinutes == 0 && endMinutes == 1440))
            {
                return (days * 1440) - currentMinutes;
            }
            // All the days are on, but it is on and off for part of these days
            else if (days == 7 && !(startMinutes == 0 && endMinutes == 1440))
            {
                if (currentMinutes < startMinutes)
                {
                    return startMinutes - currentMinutes;
                }
                else if (currentMinutes < endMinutes)
                {
                    return endMinutes - currentMinutes;
                }
                else
                {
                    return (1440 - currentMinutes) + startMinutes;
                }
            }
            else
            {
                if (currentMinutes < startMinutes)
                {
                    // If the current day is an active day, then our TTL is the start time 
                    // of today
                    if ((currentDay & this.activeDays) != 0)
                        return startMinutes - currentMinutes;
                    // The next day we change we go from off to on, so start at that 
                    // time on that day
                    else
                        return (days * 1440) + startMinutes - currentMinutes;
                }
                else if (currentMinutes < endMinutes)
                {
                    // If the current day is an active day, than we are the end time of today
                    if ((currentDay & this.activeDays) != 0)
                        return endMinutes - currentMinutes;
                    // Run out today, plus the number of days till our start
                    else
                        return (1440 - currentMinutes) + (days * 1440) - (1440 - startMinutes);
                }
                else
                {
                    // If the current day is not active, than change at the start time on the delta
                    if ((currentDay & this.activeDays) == 0)
                        return (1440 - currentMinutes) + (days * 1440) - (1440 - startMinutes);
                    // Today is active.  If tomorrow is active, start at the start time tomorrow
                    else if (days > 1)
                        return (1440 - currentMinutes) + startMinutes;
                    // Today is active, tomorrow is inactive. We are off.
                    else
                    {
                        ActiveDay nextDay = GetNextDay(currentDay);

                        days = CountDays(nextDay);

                        // End of today, plus the inactive days
                        return (1440 - currentMinutes) + (1440 * days) + startMinutes;
                    }
                }
            }            
        }

        //
        // IsActive
        //
        // Will return a flag indicating if the channel should be considered
        // active given the input date and time
        //
        public bool IsActive(DateTime currentDateTime, PMTimeZone timeZone)
        {
            if (this.active == false ||
                (this.activeLimitDays && this.activeDays == ActiveDay.None) )
                return false;
     
            // Adjust the date time for the timezone
            currentDateTime = TimeUtil.AdjustDateTime(currentDateTime, timeZone);
            
            if (this.activeExpire &&
                currentDateTime > this.activeExpireDate)
                return false;

            if (this.activeLimitTime)
            {
                int start = this.activeStartMinutes;
                int end = this.activeEndMinutes;
                int current = currentDateTime.Hour * 60 + currentDateTime.Minute;

                if (start > current || current > end)
                    return false;
            }

            if (this.activeLimitDays)
            {
                DayOfWeek dayOfWeek = currentDateTime.DayOfWeek;

                if ((GetCurrentDay(dayOfWeek) & this.ActiveDays) == 0)
                    return false;
            }
                

            return true;
        }

        //
        // Properties that control the active state
        //
        public bool Active
        {
            get { return this.active; }
            set
            {
                if (this.active != value)
                {
                    this.dataDirty = true;
                    this.active = value;
                    this.ClearCompiledItemsCache();
                }
            }
        }        

        public bool ActiveLimitTime
        {
            get { return this.activeLimitTime; }
            set
            {
                if (this.activeLimitTime != value)
                {
                    this.dataDirty = true;
                    this.activeLimitTime = value;
                    this.ClearCompiledItemsCache();
                }
            }
        }
        
        public int ActiveStartMinutes
        {
            get { return this.activeStartMinutes; }
            set
            {
                if (this.activeStartMinutes != value)
                {
                    this.dataDirty = true;
                    this.activeStartMinutes = value;
                    this.ClearCompiledItemsCache();
                }
            }
        }

        public int ActiveEndMinutes
        {
            get { return this.activeEndMinutes; }
            set
            {
                if (this.activeEndMinutes != value)
                {
                    this.dataDirty = true;
                    this.activeEndMinutes = value;
                    this.ClearCompiledItemsCache();
                }
            }
        }

        public bool ActiveLimitDays
        {
            get { return this.activeLimitDays; }
            set
            {
                if (this.activeLimitDays != value)
                {
                    this.dataDirty = true;
                    this.activeLimitDays = value;
                    this.ClearCompiledItemsCache();
                }
            }
        }

        public ActiveDay ActiveDays
        {
            get { return this.activeDays; }
            set
            {
                if (this.activeDays != value)
                {
                    this.dataDirty = true;
                    this.activeDays = value;
                    this.ClearCompiledItemsCache();
                }
            }
        }

        public bool ActiveExpire
        {
            get { return this.activeExpire; }
            set
            {
                if (this.activeExpire != value)
                {
                    this.dataDirty = true;
                    this.activeExpire = value;
                    this.ClearCompiledItemsCache();
                }
            }
        }

        public DateTime ActiveExpireDate
        {
            get { return this.activeExpireDate; }
            set
            {
                if (this.activeExpireDate != value)
                {
                    this.dataDirty = true;
                    this.activeExpireDate = value;
                    this.ClearCompiledItemsCache();
                }
            }
        }

        //
        // General properties of the channel
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
                    this.dataDirty = true;
                    this.name = value;
                }
            }
        }     

        public Puid Puid
        {
            get { return this.puid; }
        }

        public Guid SlideShowGuid
        {
            get { return this.slideShowGuid; }
        }

        public Guid ChannelGuid
        {
            get { return this.channelGuid; }
        }

        public ChannelType Type
        {
            get { return this.type; }
        }

        public DateTime CreationDate
        {
            get { return this.creationDate; }
        }

        //
        // Compilation settings for when the channel is compiled into items
        //

        //
        // IsFixedCount
        //
        // When IsFixedCount is true, the channel does not display or look at the compile
        // rules. The count is considered fixed and no pruning etc. will be done.
        //
        public virtual bool IsFixedCount
        {
            get { return false; }
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

        //
        // CreateSlideShowItem
        //
        // When a channel item is referenced in a slide show compile, CreateSlideShowItem is called
        // to turn an item in a channel into an item in a slideshow.
        //
        // This method exists on the channel because the default implementation of
        // CreateSlideShowItem on ChannelItem will in turn call back to the channel
        // to create the slide show item.
        //
        // This is useful for channels where the all the rules for initialization of the items
        // and state data for the channel exist in the channel.
        //
        // In this case, why even create an item (other than a placeholder ChannelItem) in 
        // the InitItems call
        //        
        public virtual SlideShowItem CreateSlideShowItem(Hashtable compileCache, DateTime dateContext, bool bypassCaches)
        {
            return null;
        }

        
        // 
        // Functions to deal with the channel items
        //
        

        // 
        // FetchItemsTTL
        //
        // Returns how long a fetch from the external source should be considered valid. 
        //
        // Return value is minutes.
        //
        static int defaultFetchItemsTTL = Convert.ToInt32(Config.GetSetting("DefaultFetchItemsTTL"));
        public virtual TimeSpan FetchItemsTTL()
        {
            // Default value is six hours
            return new TimeSpan(0, defaultFetchItemsTTL, 0);
        }       
       
        //
        // InitItems
        //
        // Virtual method where a channel can initialize its items 
        //
        public virtual List<ListItem> InitItems(List<Guid> guidContext, DateTime dateContext, bool bypassCaches, Hashtable compileState)
        {
            return new List<ListItem>();
        }
        

        //
        // GetItems
        //
        // Gets the target items associated with this channel.  
        //
        public List<ListItem> GetItems(List<Guid> guidContext, DateTime dateContext, bool bypassCaches, Hashtable compileState)
        {
            if (this.items == null || bypassCaches)
            {                
                this.items = this.InitItems(guidContext, dateContext, bypassCaches, compileState);
            }

            return this.items;
        }
        
        //
        // ClearItems()
        //
        // Will clear the target items from this channel
        //
        public void ClearItems()
        {
            if (this.items == null)
                return;

            this.items = null;
            this.compiledItemsCache = null;
            this.compiledItemsCacheDirty = true;
        }        


        public void DataDirty()
        {
            this.dataDirty = true;
        }

        public void ClearCompiledItemsCache()
        {
            this.compiledItemsCache = null;
            this.compiledItemsCacheDirty = true;
        }
                       
        //
        // Functions to deal with the compiled channel items
        //
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

        public bool IsCompileDirty()
        {
            return this.compiledItemsCacheDirty;
        }

        public DateTime GetCompiledItemsInvalidDate()
        {
            return this.compiledItemsCacheInvalidDate;
        }

        public DateTime GetCompiledItemsDate()
        {
            return this.compiledItemsCacheDate;
        }

        static int defaultChannelCompileTTL = Convert.ToInt32(Config.GetSetting("DefaultChannelCompileTTL"));
        public virtual int CompileTTL()
        {
            return defaultChannelCompileTTL;
        }

        public void Compile()
        {
            this.Compile(null, DateTime.MinValue, false, null);
        }               
        
        public void Compile(List<Guid> guidContext, DateTime dateContext, bool bypassCaches, Hashtable compileState)
        {
            // Create the guid context if ncessary
            if (guidContext == null)
                guidContext = new List<Guid>();

            // Get the current date if we don't have a valid date
            if (dateContext == DateTime.MinValue)
                dateContext = DateTime.Now;

            // Create a compile state if none was passed in
            if (compileState == null)
                compileState = new Hashtable();

            // Create the list of compiled items
            List<ListItem> compiledItemsCache = new List<ListItem>();

            // Initially the compiled items cache is valid forever
            DateTime compiledItemsCacheInvalidDate = DateTime.MaxValue;

            // Get the target items to compile
            List<ListItem> items = this.GetItems(guidContext, dateContext, bypassCaches, compileState);    
                       
            // Add the items to our compiled list                             
            compiledItemsCache.AddRange(items);

            // Apply compile rules if the channel is not a fixed count channel
            if (!this.IsFixedCount)
            {
                CompileUtil.PruneList(compiledItemsCache, this.compileAll, this.compileMaxCount, this.compileMinCount, this.compileOrder, this.compileInclude, this.compileAge, dateContext, this.creationDate);
            }

            // Convert all the items into slide show items
            for (int i = 0; i < compiledItemsCache.Count; i++)
            {
                compiledItemsCache[i] = compiledItemsCache[i].CreateSlideShowItem(compileState, dateContext, bypassCaches);

                // Our compile is only as valid as the items in the compile
                if (compiledItemsCache[i].ExpDate < compiledItemsCacheInvalidDate)
                    compiledItemsCacheInvalidDate = compiledItemsCache[i].ExpDate;
            }                       

            // A compile must respect its own TTL
            DateTime tempDate = dateContext.Add(new TimeSpan(0, this.CompileTTL(), 0));
            if (tempDate < compiledItemsCacheInvalidDate)
            {
                compiledItemsCacheInvalidDate = tempDate;
            }            
            
            this.compiledItemsCacheDirty = true;
            this.compiledItemsCache = compiledItemsCache;
            this.compiledItemsCacheDate = dateContext;
            this.compiledItemsCacheInvalidDate = compiledItemsCacheInvalidDate;
        }

               
        //
        // Functions to save generic data associated with the channel
        //
        public virtual void SaveDataToString(StringBuilder sb)
        {
            if (!this.active)
                sb.Append("<C_A>False</C_A>");
            if (this.activeLimitTime)
                sb.Append("<C_ALT>True</C_ALT>");
            if (this.activeStartMinutes != 0)
                sb.Append("<C_ASM>" + this.activeStartMinutes + "</C_ASM>");
            if (this.activeEndMinutes != 1440)
                sb.Append("<C_AEM>" + this.activeEndMinutes + "</C_AEM>");
            if (this.activeLimitDays)
                sb.Append("<C_ALD>True</C_ALD>");
            if (this.activeDays != ActiveDay.All)
                sb.Append("<C_AD>" + (int)this.activeDays + "</C_AD>");
            if (this.activeExpire)
                sb.Append("<C_AE>True</C_AE>");
            if (this.activeExpireDate != DateTime.MaxValue)
                sb.Append("<C_AED>" + this.activeExpireDate + "</C_AED>");
        }

        public virtual void LoadDataFromXmlNode(XmlNode node)
        {
            try { this.active = FormUtil.GetBoolean(node.SelectSingleNode("C_A").InnerText); }
            catch { }
            try { this.activeLimitTime = FormUtil.GetBoolean(node.SelectSingleNode("C_ALT").InnerText); }
            catch { }
            try { this.activeStartMinutes = FormUtil.GetNumber(node.SelectSingleNode("C_ASM").InnerText, 0, 1440); }
            catch { }
            try { this.activeEndMinutes = FormUtil.GetNumber(node.SelectSingleNode("C_AEM").InnerText, 0, 1440); }
            catch { }
            try { this.activeLimitDays = FormUtil.GetBoolean(node.SelectSingleNode("C_ALD").InnerText); }
            catch { }
            try { this.activeDays = (ActiveDay)FormUtil.GetNumber(node.SelectSingleNode("C_AD").InnerText); }
            catch { }
            try { this.activeExpire = FormUtil.GetBoolean(node.SelectSingleNode("C_AE").InnerText); }
            catch { }
            try { this.activeExpireDate = FormUtil.GetDate(node.SelectSingleNode("C_AED").InnerText); }
            catch { }
        }

        public virtual void LoadDataFromQueryString(HttpRequest request)
        {
            try { this.active = FormUtil.GetBoolean(request.QueryString["C_A"]); }
            catch { }
            try { this.activeLimitTime = FormUtil.GetBoolean(request.QueryString["C_ALT"]); }
            catch { }
            try { this.activeStartMinutes = FormUtil.GetNumber(request.QueryString["C_ASM"], 0, 1440); }
            catch { }
            try { this.activeEndMinutes = FormUtil.GetNumber(request.QueryString["C_AEM"], 0, 1440); }
            catch { }
            try { this.activeLimitDays = FormUtil.GetBoolean(request.QueryString["C_ALD"]); }
            catch { }
            try { this.activeDays = (ActiveDay)FormUtil.GetNumber(request.QueryString["C_AD"]); }
            catch { }
            try { this.activeExpire = FormUtil.GetBoolean(request.QueryString["C_AE"]); }
            catch { }
            try { this.activeExpireDate = FormUtil.GetDate(request.QueryString["C_AED"]); }
            catch { }
        }

        private void RootLoadDataFromString(string input)
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.LoadXml(input);

            XmlNode node = xmlDocument.SelectSingleNode("Channel");

            this.LoadDataFromXmlNode(node);
        }

        private void RootSaveDataToString(StringBuilder sb)
        {
            sb.Append("<Channel>");
            this.SaveDataToString(sb);
            sb.Append("</Channel>");
        }

        public static Channel ChannelFromType(Puid puid, Guid slideShowGuid, Guid channelGuid, ChannelType type)
        {
            Channel channel;

            if (type == ChannelType.Rss)
                channel = (Channel)new RssChannel(puid, slideShowGuid, channelGuid);
            else if (type == ChannelType.Space)
                channel = (Channel)new SpaceChannel(puid, slideShowGuid, channelGuid);
            else if (type == ChannelType.Facebook)
                channel = (Channel)new FacebookChannel(puid, slideShowGuid, channelGuid);
            else if (type == ChannelType.Reference)
                channel = (Channel)new ReferenceChannel(puid, slideShowGuid, channelGuid);
            else if (type == ChannelType.Static)
                channel = (Channel)new StaticChannel(puid, slideShowGuid, channelGuid);
            else if (type == ChannelType.Weather)
                channel = (Channel)new WeatherChannel(puid, slideShowGuid, channelGuid);
            else if (type == ChannelType.WebPage)
                channel = (Channel)new WebPageChannel(puid, slideShowGuid, channelGuid);
            else if (type == ChannelType.USTraffic)
                channel = (Channel)new USTrafficChannel(puid, slideShowGuid, channelGuid);
            else if (type == ChannelType.Flickr)
                channel = (Channel)new FlickrChannel(puid, slideShowGuid, channelGuid);
            else if (type == ChannelType.SmugMug)
                channel = (Channel)new SmugMugChannel(puid, slideShowGuid, channelGuid);
            else if (type == ChannelType.FixedReference)
                channel = (Channel)new FixedReferenceChannel(puid, slideShowGuid, channelGuid);
            else
                channel = new Channel(puid, slideShowGuid, channelGuid, ChannelType.Unknown);

            channel.type = type;

            return channel;
        }

        //
        // Functions to load and save the channel to/from the database
        //
        public static Channel LoadFromDB(SqlDataReader reader)
        {
            Puid puid = new Puid(reader.GetInt32(1), reader.GetInt32(0));
            Guid slideShowGuid = reader.GetGuid(2);
            Guid channelGuid = reader.GetGuid(3);
            ChannelType type = (ChannelType)reader.GetInt32(4);

            Channel channel = ChannelFromType(puid, slideShowGuid, channelGuid, type);
            
            channel.channelOrder = reader.GetInt32(5);
            channel.creationDate = reader.GetDateTime(6);
            channel.name = reader.IsDBNull(7) ? null : reader.GetString(7);            

            channel.compileAll = reader.GetBoolean(8);
            channel.compileMaxCount = reader.GetInt32(9);
            channel.compileMinCount = reader.GetInt32(10);
            channel.compileOrder = (CompileOrder)reader.GetInt32(11);
            channel.compileInclude = (CompileInclude)reader.GetInt32(12);
            channel.compileAge = reader.GetInt32(13);

            if (!reader.IsDBNull(14))
                channel.RootLoadDataFromString(reader.GetString(14));

            return channel;
        }
        
        public static Channel LoadFromDb(Puid puid, Guid slideShowGuid, Guid channelGuid)
        {
            string sql = 
                "select " +
                "   PuidHigh, PuidLow, SlideShowGuid, ChannelGuid, Type, " +
                "   ChannelOrder, CreationDate, Name, " +
                "   CompileAll, CompileMaxCount, CompileMinCount, CompileOrder, CompileInclude, CompileAge, " +
                "   Data " +
                "from Channels " +
                "where PuidHash = @PuidHash and PuidHigh = @PuidHigh and PuidLow = @PuidLow and SlideShowGuid = @SlideShowGuid and ChannelGuid = @ChannelGuid";
            using (PhotoMixQuery query = new PhotoMixQuery(sql, CommandType.Text))
            {
                query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puid.GetHashCode();
                query.Parameters.Add("@PuidHigh", SqlDbType.Int).Value = puid.PuidHigh;
                query.Parameters.Add("@PuidLow", SqlDbType.Int).Value = puid.PuidLow;
                query.Parameters.Add("@SlideShowGuid", SqlDbType.UniqueIdentifier).Value = slideShowGuid;
                query.Parameters.Add("@ChannelGuid", SqlDbType.UniqueIdentifier).Value = channelGuid;

                query.Reader.Read();

                return LoadFromDB(query.Reader);
            }
        }

        public static List<Channel> CloneChannelsFromDbToPuid(Puid oldPuid, Guid oldSlideShowGuid, Puid newPuid, Guid newSlideShowGuid)
        {
            List<Channel> channels = LoadChannelsFromDb(oldPuid, oldSlideShowGuid);

            foreach (Channel channel in channels)
            {
                channel.channelGuid = Guid.NewGuid();
                channel.slideShowGuid = newSlideShowGuid;
                channel.puid = newPuid;
                channel.dataDirty = true;
                channel.ClearCompiledItemsCache();
                channel.ClearItems();
                channel.SaveToDb();
            }

            return channels;
        }

        public static List<Channel> LoadChannelsFromDb(Puid puid, Guid slideShowGuid)
        {
            List<Channel> channels = new List<Channel>();

            string sql = "" +
                "select " +
                "   PuidHigh, PuidLow, SlideShowGuid, ChannelGuid, Type, " +
                "   ChannelOrder, CreationDate, Name, " +
                "   CompileAll, CompileMaxCount, CompileMinCount, CompileOrder, CompileInclude, CompileAge, " +
                "   Data " +                
                "from Channels " +
                "where PuidHash = @PuidHash and PuidHigh = @PuidHigh and PuidLow = @PuidLow and SlideShowGuid = @SlideShowGuid " +
                "order by ChannelOrder, CreationDate asc";
            using (PhotoMixQuery query = new PhotoMixQuery(sql, CommandType.Text))
            {
                query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puid.GetHashCode();
                query.Parameters.Add("@PuidHigh", SqlDbType.Int).Value = puid.PuidHigh;
                query.Parameters.Add("@PuidLow", SqlDbType.Int).Value = puid.PuidLow;
                query.Parameters.Add("@SlideShowGuid", SqlDbType.UniqueIdentifier).Value = slideShowGuid;

                while (query.Reader.Read())
                {
                    Channel channel = LoadFromDB(query.Reader);

                    channels.Add(channel);
                }
            }

            return channels;
        }

        public void DeleteFromDb()
        {
            string sql = "" +
                "delete from Channels " +
                "where PuidHash = @PuidHash and PuidHigh = @PuidHigh and PuidLow = @PuidLow and SlideShowGuid = @SlideShowGuid and ChannelGuid = @ChannelGuid";

            using (PhotoMixQuery query = new PhotoMixQuery(sql, CommandType.Text))
            {
                query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puid.GetHashCode();
                query.Parameters.Add("@PuidHigh", SqlDbType.Int).Value = puid.PuidHigh;
                query.Parameters.Add("@PuidLow", SqlDbType.Int).Value = puid.PuidLow;
                query.Parameters.Add("@SlideShowGuid", SqlDbType.UniqueIdentifier).Value = this.SlideShowGuid;
                query.Parameters.Add("@ChannelGuid", SqlDbType.UniqueIdentifier).Value = this.ChannelGuid;

                query.Execute();
            }
        }

        //
        // ItemsNeedCompileState
        // 
        // This virtual function indicates that the items
        // in the channel create stateful data as part of the compilation
        // process.  This is an indicator to the channel that the items
        // in the channel compile should be cached.
        //
        public virtual bool ItemsNeedCompileState
        {
            get { return false; }
        }

        public void ChangedItemsNeedCompileState()
        {
            this.changedItemsNeedCompileState = true;
        }
        
        //
        // SaveCacheCompile
        //
        // Indicates if the channel should bother to save its cached compile data
        //
        private bool SaveCacheCompile()
        {
            bool requiresCompile = false;

            // See if the channel compile rules select a subset of items
            // as part of the compilation process
            if (!this.IsFixedCount &&
                !(this.compileAll && this.CompileOrder == CompileOrder.Listed))
                requiresCompile = true;

            // If we create stateful data in the compile, or we've trimmed to a subset
            // of items, we should cache our compile state.  Otherwise it is cheap enough
            // to re-do it so don't waste the storage in the DB
            if (this.ItemsNeedCompileState || requiresCompile)
                return true;

            return false;            
        }       

        public void SaveToDb()
        {
            if (!this.dataDirty && 
                ((!this.compiledItemsCacheDirty || !this.SaveCacheCompile()) && !this.changedItemsNeedCompileState)
                )
                return;            

            string sql = "" +
                "if not exists (select SlideShowGuid, ChannelGuid from Channels where PuidHigh = @PuidHigh and PuidLow = @PuidLow and SlideShowGuid = @SlideShowGuid and ChannelGuid = @ChannelGuid)" +                
                "    insert into Channels (" +
                "       PuidHash, PuidHigh, PuidLow, SlideShowGuid, ChannelGuid, ChannelOrder, Type, Name, " +
                "       CompileAll, CompileMaxCount, CompileMinCount, CompileOrder, CompileInclude, CompileAge, " +
                "       Data, " +
                "       SaveCacheCompile, CompiledItemsCache, CompiledItemsCacheDate, CompiledItemsCacheInvalidDate, " +
                "       CreationDate, ModifiedDate " +                
                "    )" +
                "    values (" +
                "       @PuidHash, @PuidHigh, @PuidLow, @SlideShowGuid, @ChannelGuid, @ChannelOrder, @Type, @Name, " +
                "       @CompileAll, @CompileMaxCount, @CompileMinCount, @CompileOrder, @CompileInclude, @CompileAge, " +
                "       @Data, " +               
                "       @SaveCacheCompile, @CompiledItemsCache, @CompiledItemsCacheDate, @CompiledItemsCacheInvalidDate, " +
                "       GetDate(), GetDate() " +                
                "    )" +
                "else" +
		        "    update Channels" +
		        "    set " + 
                "       ChannelOrder = @ChannelOrder, Type = @Type, Name = @Name, " +
                "       CompileAll = @CompileAll, CompileMaxCount = @CompileMaxCount, CompileMinCount = @CompileMinCount, CompileOrder = @CompileOrder, CompileInclude = @CompileInclude, CompileAge = @CompileAge, " +
                "       Data = @Data, " +                                
                "       SaveCacheCompile = @SaveCacheCompile, " +
                "       CompiledItemsCache = Case @CompiledItemsCacheDirty when 0 then CompiledItemsCache when 1 then @CompiledItemsCache end, " +
                "       CompiledItemsCacheDate = Case @CompiledItemsCacheDirty when 0 then CompiledItemsCacheDate when 1 then @CompiledItemsCacheDate end, " +
                "       CompiledItemsCacheInvalidDate = Case @CompiledItemsCacheDirty when 0 then CompiledItemsCacheInvalidDate when 1 then @CompiledItemsCacheInvalidDate end, " +
                "       ModifiedDate = GetDate() " +
                "    where PuidHash = @PuidHash and PuidHigh = @PuidHigh and PuidLow = @PuidLow and SlideShowGuid = @SlideShowGuid and ChannelGuid = @ChannelGuid";


            using (PhotoMixQuery query = new PhotoMixQuery(sql, CommandType.Text))
            {
                query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puid.GetHashCode();
                query.Parameters.Add("@PuidHigh", SqlDbType.Int).Value = puid.PuidHigh;
                query.Parameters.Add("@PuidLow", SqlDbType.Int).Value = puid.PuidLow;
                query.Parameters.Add("@SlideShowGuid", SqlDbType.UniqueIdentifier).Value = this.SlideShowGuid;
                query.Parameters.Add("@ChannelGuid", SqlDbType.UniqueIdentifier).Value = this.ChannelGuid;
                query.Parameters.Add("@ChannelOrder", SqlDbType.Int).Value = this.channelOrder;
                query.Parameters.Add("@Type", SqlDbType.Int).Value = (int)this.Type;
                if (this.Name == null)
                    query.Parameters.Add("@Name", SqlDbType.NVarChar).Value = DBNull.Value;
                else
                query.Parameters.Add("@Name", SqlDbType.NVarChar).Value = this.Name;                

                query.Parameters.Add("@CompileAll", SqlDbType.Bit).Value = this.CompileAll;
                query.Parameters.Add("@CompileMaxCount", SqlDbType.Int).Value = this.CompileMaxCount;
                query.Parameters.Add("@CompileMinCount", SqlDbType.Int).Value = this.CompileMinCount;
                query.Parameters.Add("@CompileOrder", SqlDbType.Int).Value = this.CompileOrder;
                query.Parameters.Add("@CompileInclude", SqlDbType.Int).Value = this.CompileInclude;
                query.Parameters.Add("@CompileAge", SqlDbType.Int).Value = this.CompileAge;

                StringBuilder sb = new StringBuilder();
                this.RootSaveDataToString(sb);            
                query.Parameters.Add("@Data", SqlDbType.NVarChar).Value = sb.ToString();
                
                query.Parameters.Add("@SaveCacheCompile", SqlDbType.Bit).Value = this.SaveCacheCompile();
                query.Parameters.Add("@CompiledItemsCacheDirty", SqlDbType.Int).Value = (this.compiledItemsCacheDirty ? 1 : 0);
                if (this.compiledItemsCache == null || this.compiledItemsCacheDirty == false || this.SaveCacheCompile() == false)
                {
                    query.Parameters.Add("@CompiledItemsCache", SqlDbType.Text).Value = DBNull.Value;
                    query.Parameters.Add("@CompiledItemsCacheDate", SqlDbType.DateTime).Value = DBNull.Value;
                    query.Parameters.Add("@CompiledItemsCacheInvalidDate", SqlDbType.DateTime).Value = DBNull.Value;
                }
                else
                {
                    sb = new StringBuilder();
                    ListItem.RootSaveListToString(sb, this.compiledItemsCache);
                    query.Parameters.Add("@CompiledItemsCache", SqlDbType.Text).Value = sb.ToString();
                    query.Parameters.Add("@CompiledItemsCacheDate", SqlDbType.DateTime).Value = this.compiledItemsCacheDate;
                    query.Parameters.Add("@CompiledItemsCacheInvalidDate", SqlDbType.DateTime).Value = this.compiledItemsCacheInvalidDate;
                    this.compiledItemsCacheDirty = false;
                }                
                
                query.Execute();
            }
        }

        static public int UserChannelCount(Puid puid)
        {
            int channelCount = 0;

            // to simplify calling code path... return 0 for a null user
            if (puid == Puid.Empty)
                return channelCount;

            using (PhotoMixQuery query = new PhotoMixQuery("UserChannelCount"))
            {
                query.Parameters.Add("@PuidHash", SqlDbType.Int).Value = puid.GetHashCode();
                query.Parameters.Add("@PuidLow", SqlDbType.Int).Value = puid.PuidLow;
                query.Parameters.Add("@PuidHigh", SqlDbType.Int).Value = puid.PuidHigh;
                if (query.Reader.Read())
                {
                    channelCount = query.Reader.GetInt32(0);
                }
            }

            return channelCount;
        }

        // *************************************
        //			Jobs and Processing
        // *************************************

        static public int TotalCountChannels
        {
            get
            {
                //$ Change to stored proc
                using (PhotoMixQuery query = new PhotoMixQuery("select count(*) as ChannelCount from SlideShowChannels", CommandType.Text))
                {
                    if (!query.Reader.Read())
                        throw new PhotoMixException(null, PhotoMixError.InternalError, "Failed to get channel count in PmxJobs loop.");
                    int channelCount = (int)query.Reader["ChannelCount"];
                    return channelCount;
                }
            }
        }

        // process a single channel by 'checking it out' with SelectNextProcessRequest and then calling Process on it
        static public bool ProcessOne()
        {
            return false; // didn't process one
        }

    }
}
