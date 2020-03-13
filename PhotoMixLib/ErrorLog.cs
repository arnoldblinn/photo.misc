using System;
using System.Collections;
using System.Collections.Generic;
using System.Web.Services.Protocols;
using System.Data.SqlClient;
using System.Data;
using Msn.Framework;

namespace Msn.PhotoMix
{
	public class ErrorLog
	{
		static int EventIDGeneralError = 1000;

		static private Dictionary<string, DateTime> waitUntilTimes = new Dictionary<string, DateTime>();

		static ErrorLog()
		{
		}

		static public void Init()
		{
			SoapExceptionHandler.LogException += new LogExceptionHandler(WriteEntry);
		}

		static void WriteLogEntry(Exception e)
		{
			throw new Exception("The method or operation is not implemented.");
		}

		// Table based lookup for mapping exceptions to events.
		// The FilterTokens field of the EventLogFilters table contains a | delimited list of strings.
		// ie: string1|string2|string3
		// To match the filter we check to make sure each string is found in the call stack in order
		struct EventLogFilter
		{
			public string[] FilterTokens;
			public string Description;
			public int EventId;
		}

		static private List<EventLogFilter> filters;
		static private DateTime refreshFilters = DateTime.MinValue;

		static private List<EventLogFilter> Filters
		{
			get
			{
				if (filters != null && refreshFilters > DateTime.Now)
					return filters;
				List<EventLogFilter> filtersNew = new List<EventLogFilter>();
                using (PhotoMixQuery query = new PhotoMixQuery("select FilterTokens, Description, EventId from EventLogFilters", CommandType.Text))
				{
					while (query.Reader.Read())
					{
						EventLogFilter filter = new EventLogFilter();
						filter.FilterTokens = query.Reader.GetString(0).Split('|');
						filter.Description = query.Reader.GetString(1);
						filter.EventId = query.Reader.GetInt32(2);
						filtersNew.Add(filter);
					}
					refreshFilters = DateTime.Now.Add(TimeSpan.FromMinutes(5));
				}
				filters = filtersNew;
				return filters;
			}
		}

		private static bool FilterError(ref string error, ref int eventId)
		{
			// grab the current filter list so it doesn't refresh on us in another thread
			List<EventLogFilter> filters = Filters;
			foreach (EventLogFilter filter in filters)
			{
				// the error we are searching through
				string errorSearch = error;
				// intial value of our match flag
				bool matches = true;

				// find all tokens in order
				foreach (string filterToken in filter.FilterTokens)
				{
					int i = errorSearch.IndexOf(filterToken);
					if (i == -1)
					{
						matches = false;
						break;
					}
					// advance down the string we are searching
					errorSearch = errorSearch.Substring(i + filterToken.Length);
				}
				if (matches)
				{
					error = filter.Description + "\r\n" + error;
					eventId = filter.EventId;
					return true;
				}
			}

			return false;
		}


		// our public method to write an entry given an exeption

		public static void WriteEntry(Exception e)
		{
			try
			{
                // check to see if this is a PhotoMixException that we don't want to log
                PhotoMixException photoMixException = e as PhotoMixException;
                if (photoMixException != null && photoMixException.LogError == false)
					return;

				string error = e.ToString();

				// this is out simple way of throttling error logs
				if (waitUntilTimes.ContainsKey(error) && waitUntilTimes[error] > DateTime.Now)
					return;
				DateTime waitUntil = DateTime.Now.Add(TimeSpan.FromMinutes(5));
				waitUntilTimes[error] = waitUntil;

				int eventID = EventIDGeneralError;
				System.Diagnostics.EventLog.WriteEntry("PhotoMix", error, System.Diagnostics.EventLogEntryType.Error, (int)eventID);
				bool filtered = FilterError(ref error, ref eventID);
			}
			catch
			{
			}
		}
	}
}
