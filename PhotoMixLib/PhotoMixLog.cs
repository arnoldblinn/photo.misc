using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Web.Services.Protocols;
using Msn.Framework;

namespace Msn.PhotoMix
{
	/// <summary>
	/// Summary description for PhotoMixLog.
	/// </summary>
	public class PhotoMixLog
	{
		public enum EventId
		{
			ProcessSucceeded = 1,
			ProcessFailed = 2,
			ProcessReleasingDomain = 3,
			ProcessReservingNamespace = 4,
			ProcessEnablingHotmail = 5,
			ProcessDisablingHotmail = 6,
			ProcessMigrateOwnerSucceeded = 7,
			ProcessMigrateOwnerFailed = 8,
			ProcessMigrateOwnerEmail = 9,
			ProcessMigrateOwnerRecycle = 10,
			ProcessGrantNamespaceAdmin = 11,
			ProcessEnablingExchange = 12,
			ProcessDisablingExchange = 13,
			SignupBlockedEasiCount = 101,
			SignupBlockedWord = 102,
			SignupBlockedRoot = 103,
			SignupBlockedSubDomain = 104,
			SignupBlockedNamespace = 105,
			SignupBlockedIdn = 106,
			ImportDomainFailed = 201,
			ImportDomainActive = 202,
			ImportDomainSuspend = 203,
			ImportDomainCancel = 204,
		}

        static PhotoMixLog()
		{
		}

		private static string ErrorFromSqlException(SqlException e)
		{
			string error = string.Format("Sql Error Encountered:\r\nMessage: {0}\r\nClass: {1}\r\nNumber: {2}\r\nState: {3}\r\nServer: {4}\r\nSource: {5}\r\nProcedure: {6}\r\nLine Number: {7}\r\nHelp: {8}\r\nStack:\r\n{9}\r\n",
				e.Message, e.Class, e.Number, e.State, e.Server, e.Source, e.Procedure, e.LineNumber, e.HelpLink, e.StackTrace);
			return error;			
		}

		private static string ErrorFromSoapException(SoapException e)
		{
			string error = string.Format("Soap Service Error Encountered:\r\nMessage: {0}\r\nSource: {1}\r\nCode: {2}\r\nActor: {3}\r\nHelp: {4}\r\nDetail:\r\n{5}\r\nStack:\r\n{6}\r\n",
				e.Message, e.Source, e.Code, e.Actor, e.HelpLink, e.Detail.OuterXml, e.StackTrace);
			return error;			
		}

		static public void CreateLogEntry(string domainName, EventId eventId)
		{
			CreateLogEntry(domainName, DateTime.Now, eventId, string.Empty);
		}

		static public void CreateLogEntry(string domainName, EventId eventId, string eventText)
		{
			CreateLogEntry(domainName, DateTime.Now, eventId, eventText);
		}

		static public void CreateLogEntry(string domainName, EventId eventId, Exception e)
		{
			string error = e.ToString();
			Type errorType = e.GetType();

			//fix up error info with more details
			if (errorType == typeof(SoapException))
				error = ErrorFromSoapException((SoapException)e);
			else if (errorType == typeof(SqlException))
				error = ErrorFromSqlException((SqlException) e);
								
			CreateLogEntry(domainName, DateTime.Now, eventId, error);
		}


		static public void CreateLogEntry(string domainName, DateTime eventTime, 
			EventId eventId, string eventText)
		{
			using (PhotoMixQuery query = new PhotoMixQuery("CreateLogEntry"))
			{
				query.Parameters.Add("@DomainName", SqlDbType.VarChar, 256).Value = domainName;
				query.Parameters.Add("@EventTime", SqlDbType.DateTime).Value = eventTime;
				query.Parameters.Add("@EventId", SqlDbType.Int).Value = (int)eventId;
				query.Parameters.Add("@EventText", SqlDbType.VarChar, 2048).Value = (eventText.Length > 2048) ? eventText.Substring(0, 2048) : eventText;
				int rows = query.Execute();
			}
		}

		static public void LogWebRequests(string webRequestKey, string webRequestMethod)
		{
			using (PhotoMixQuery query = new PhotoMixQuery("LogWebRequests", CommandType.StoredProcedure))
			{
				query.Parameters.Add("@WebRequestKey", SqlDbType.VarChar, 100).Value = webRequestKey;
				query.Parameters.Add("@WebRequestMethod", SqlDbType.VarChar, 100).Value = (webRequestMethod.Length > 100) ? webRequestMethod.Substring(0, 100) : webRequestMethod;
				int rows = query.Execute();
			}
		}
    }
}
