using System;

namespace Msn.PhotoMix.Passport
{
	public class Puid
	{
		public static Puid Empty = new Puid(0, 0);

		private UInt64 puid;


		public Puid(int puidLow, int puidHigh)
		{
			puid = ((UInt64) (((UInt32) (puidLow)) | ((UInt64) ((UInt32) (puidHigh))) << 32));
		}
        

		public Puid(string puid)
		{
			this.puid = Convert.ToUInt64(puid, 16);
		}

		public int PuidLow
		{
			get { return (int)(UInt32)puid; }
		}

		public int PuidHigh
		{
			get { return (int)((UInt32) (((UInt64) (puid) >> 32) & 0xFFFFFFFF)); }
		}

		public override bool Equals(object obj)
		{
			if (obj == null)
				return false;
			return obj.GetType() == this.GetType() && ((Puid)obj).puid == puid;
		}

		public override int GetHashCode()
		{
			return puid.GetHashCode();
		}
		
		public string ToHex()
		{
			return PuidHigh.ToString("X8") + PuidLow.ToString("X8");
			
		}

		public static bool operator ==(Puid puid1, Puid puid2) 
		{
			if (((object)puid1) != null && ((object)puid2) != null)
				return puid1.PuidHigh == puid2.PuidHigh && puid1.PuidLow == puid2.PuidLow;
			else
				return ((object)puid1) == ((object)puid2);
		}

		public static bool operator !=(Puid puid1, Puid puid2) 
		{
			if (((object)puid1) != null && ((object)puid2) != null)
				return !(puid1.PuidHigh == puid2.PuidHigh && puid1.PuidLow == puid2.PuidLow);
			else
				return ((object)puid1) != ((object)puid2);
		}

	}
}
