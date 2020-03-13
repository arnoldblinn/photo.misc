using System;
using System.Collections.Generic;
using System.Text;

namespace Msn.PhotoMix
{
    public class PhotoMixErrorAttribute : Attribute
	{
		private bool logError;
		private string message;

		public bool LogError
		{
			get { return logError; }
			set { logError = value; }
		}

		public string Message
		{
			get { return message; }
			set { message = value; }
		}
	
	}
}
