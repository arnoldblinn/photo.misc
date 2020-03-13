using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

using Msn.PhotoMix.Passport;
using Msn.Framework;
using Msn.PhotoMix.Util;

namespace Msn.PhotoMix.SlideShow
{
    public class ReferenceChannel : Channel
    {
        private string friendlyName;
        private string referenceId;
        
        private int referenceHash;
        private Guid referenceGuid;

        private string pin = null;        

        public ReferenceChannel(Puid puid, Guid slideShowGuid, Guid channelGuid, ChannelType channelType)
            : base(puid, slideShowGuid, channelGuid, channelType)
        {

        }

        public ReferenceChannel(Puid puid, Guid slideShowGuid, Guid channelGuid)
            : base(puid, slideShowGuid, channelGuid, ChannelType.Reference)
        {

        }


        public ReferenceChannel(Puid puid, Guid slideShowGuid, ChannelType channelType)
            : base(puid, slideShowGuid, channelType)
        {

        }

        public ReferenceChannel(Puid puid, Guid slideShowGuid)
            : base(puid, slideShowGuid, ChannelType.Reference)
        {

        }

        public override void SaveDataToString(StringBuilder sb)
        {
            base.SaveDataToString(sb);

            if (!String.IsNullOrEmpty(this.referenceId))
                sb.Append("<RefId>" + this.referenceId.ToString() + "</RefId>");
            if (!String.IsNullOrEmpty(this.friendlyName))
                sb.Append("<FriendlyN>" + this.friendlyName.ToString() + "</FriendlyN>");            
            if (!String.IsNullOrEmpty(this.pin))
                sb.Append("<Pin>" + this.pin.ToString() + "</Pin>");            
        }

        public override void LoadDataFromXmlNode(System.Xml.XmlNode node)
        {
            base.LoadDataFromXmlNode(node);

            try
            {
                this.referenceId = node.SelectSingleNode("RefId").InnerText;
                this.referenceGuid = SlideShow.LookupId(this.referenceId, out this.referenceHash);
            }
            catch { }
            try
            {
                this.friendlyName = node.SelectSingleNode("FriendlyN").InnerText;
            }
            catch { }
            try
            {
                this.pin = node.SelectSingleNode("Pin").InnerText;
            }
            catch { }            
        }

        public string FriendlyName
        {
            get { return this.friendlyName; }
            set
            {
                if (this.friendlyName != value)
                {
                    this.friendlyName = value;

                    if (!String.IsNullOrEmpty(this.friendlyName))
                    {
                        this.referenceId = null;
                        this.referenceGuid = Guid.Empty;
                        this.referenceHash = 0;
                    }

                    this.slideShow = null;
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public string ReferenceId
        {
            get { return this.referenceId; }
            set
            {
                if (this.referenceId != value)
                {
                    this.referenceId = value;
                    this.referenceGuid = SlideShow.LookupId(this.referenceId, out this.referenceHash);

                    if (!String.IsNullOrEmpty(this.referenceId))
                    {
                        this.friendlyName = null;
                    }
                    this.slideShow = null;
                    this.DataDirty();                    
                    this.ClearItems();
                }
            }
        }

        public string Pin
        {
            get { return this.pin; }
            set
            {
                if (this.pin != value)
                {
                    this.pin = value;
                    this.slideShow = null;                    
                    this.DataDirty();
                    this.ClearItems();
                }
            }
        }

        public Guid GetReferenceGuid()
        {
            if (!String.IsNullOrEmpty(this.referenceId))
            {
                return this.referenceGuid;
            }
            else
            {
                int slideShowHash;
                return Msn.PhotoMix.SlideShow.FriendlyName.LookupFriendlyName(this.friendlyName, out slideShowHash);
            }
        }                

        private SlideShow slideShow = null;
        public SlideShow GetSlideShow()
        {            
            if (this.slideShow == null)
            {
                try
                {
                    SlideShow slideShow;
                    if (this.friendlyName != null)
                    {
                        slideShow = SlideShow.LoadFromFriendlyName(this.friendlyName);
                    }
                    else
                    {
                        slideShow = SlideShow.LoadFromDb(this.referenceHash, this.referenceGuid);
                    }
                     

                    if (slideShow != null && !String.IsNullOrEmpty(slideShow.Pin))
                    {
                        if (this.pin != slideShow.Pin)
                            slideShow = null;
                    }

                    this.slideShow = slideShow;
                }
                catch (Exception)
                {
                }
            }

            return this.slideShow;
        }

        public override List<ListItem> InitItems(List<Guid> guidContext, DateTime dateContext, bool bypassCaches, Hashtable compileState)
        {
            List<ListItem> items;
            SlideShow slideShow = this.GetSlideShow();
            
            if (slideShow != null)
            {
                items = slideShow.GetCompiledItems(guidContext, dateContext, bypassCaches, compileState);

                slideShow.SaveChannelsToDb();
                slideShow.SaveToDb();
            }
            else
                items = new List<ListItem>();

            return items;
        }     
    }
}
