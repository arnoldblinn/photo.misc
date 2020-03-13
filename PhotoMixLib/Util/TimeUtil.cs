using System;
using System.Collections.Generic;
using System.Text;

namespace Msn.PhotoMix.Util
{
    public enum PMTimeZone
    {
        GMT = 0,
        EST = 1,
        EDT = 2,
        CST = 3,
        CDT = 4,
        MST = 5,
        MDT = 6,
        PST = 7,
        PDT = 8
    }

    class TimeUtil
    {        
        static public string DateToRSSString(DateTime dateTime)
        {
            return dateTime.ToString("ddd, %d MMM yyyy HH:mm:ss zzz");
        }

        static private int[] adjusts = { 0, -4, -4, -5, -5, -6, -6, -7, -7 };
        static private bool[] daylight = { false, false, true, false, true, false, true, false, true };
       
        static public DateTime AdjustDateTime(DateTime dateTime, PMTimeZone timeZone)
        {
            // Adjust the date time passed to GMT
            dateTime = dateTime.Subtract(TimeZone.CurrentTimeZone.GetUtcOffset(dateTime));

            // Adjust the date time to the target time zone
            dateTime = dateTime.Add(new TimeSpan(0, adjusts[(int)timeZone] * 60, 0));

            // See if it is daylight savings time, and if the timezone respects daylight, and adjust if not
            if (dateTime.IsDaylightSavingTime() && !daylight[(int)timeZone])
                dateTime = dateTime.Subtract(new TimeSpan(0, 60, 0));

            return dateTime;
        }  
    }


}
