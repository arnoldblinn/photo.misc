using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

using Msn.Framework;
using Msn.PhotoMix.Passport;

namespace Msn.PhotoMix.SlideShow
{
    public class FixedReferenceChannel : ReferenceChannel
    {
        // Parameters only apply to a "fixedReference"
        private bool fixedReferenceHasCount;        
        private bool fixedReferenceHasImage;

        public FixedReferenceChannel(Puid puid, Guid slideShowGuid, Guid channelGuid)
            : base(puid, slideShowGuid, channelGuid, ChannelType.FixedReference)
        {

        }

        public FixedReferenceChannel(Puid puid, Guid slideShowGuid)
            : base(puid, slideShowGuid, ChannelType.FixedReference)
        {

        }

        public override void SaveDataToString(StringBuilder sb)
        {
            base.SaveDataToString(sb);

            sb.Append("<FRHC>" + this.fixedReferenceHasCount.ToString() + "</FRHC>");                
            sb.Append("<FRHI>" + this.fixedReferenceHasImage.ToString() + "</FRHI>");
        }

        public override void LoadDataFromXmlNode(System.Xml.XmlNode node)
        {
            base.LoadDataFromXmlNode(node);

            try { this.fixedReferenceHasCount = FormUtil.GetBoolean(node.SelectSingleNode("FRHC").InnerText); }
            catch { }                

            try { this.fixedReferenceHasImage = FormUtil.GetBoolean(node.SelectSingleNode("FRHI").InnerText); }
            catch { }            
        }        

        public override bool IsFixedCount
        {
            get
            {
                return this.fixedReferenceHasCount;
            }            
        }        

        public bool FixedReferenceHasCount
        {
            get
            {
                return this.fixedReferenceHasCount;
            }
            set
            {
                if (this.fixedReferenceHasCount != value)
                {
                    this.fixedReferenceHasCount = value;
                    this.DataDirty();
                    this.ClearCompiledItemsCache();
                }
            }
        }
        
        public bool FixedReferenceHasImage
        {
            get
            {
                return this.fixedReferenceHasImage;
            }
            set
            {
                if (this.fixedReferenceHasImage != value)
                {
                    this.fixedReferenceHasImage = value;
                    this.DataDirty();
                }
            }
        }
    }
}
