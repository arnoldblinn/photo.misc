using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Web.Services.Protocols;
using Msn.Framework;
using System.Reflection;

namespace Msn.PhotoMix
{
	public class PhotoMixException : Exception, ISoapExceptionFormatter
	{
		private PhotoMixError photoMixError;
		private String partnerErrorCode;
		private String partnerErrorDetail;
		private String message;
		private String context;
		private bool logError = true;

		public PhotoMixException(string context, PhotoMixError photoMixError)
		{
			this.context = context;
			this.photoMixError = photoMixError;
			GetPhotoMixErrorAttributes();
		}

		public PhotoMixException(string context, PhotoMixError photoMixError, string message)
			:
			base(message)
		{
			this.context = context;
			this.message = message;
			this.photoMixError = photoMixError;
			GetPhotoMixErrorAttributes();
		}

		public PhotoMixException(string context, PhotoMixError photoMixError, String partnerErrorCode, String partnerErrorDetail)
		{
			this.context = context;
			this.photoMixError = photoMixError;
			this.partnerErrorCode = partnerErrorCode;
			this.partnerErrorDetail = partnerErrorDetail;
			GetPhotoMixErrorAttributes();
		}

		public PhotoMixException(string context, PhotoMixError photoMixError, string message, Exception innerException)
			:
			base(message, innerException)
		{
			this.context = context;
			this.photoMixError = photoMixError;
			this.message = message;
			GetPhotoMixErrorAttributes();
		}

	
		public void GetPhotoMixErrorAttributes()
		{
			FieldInfo fi = photoMixError.GetType().GetField(photoMixError.ToString());
			PhotoMixErrorAttribute[] attributes = (PhotoMixErrorAttribute[])fi.GetCustomAttributes(typeof(PhotoMixErrorAttribute), false);
			if (attributes.Length > 0)
			{
				if (this.message == null)
					message = attributes[0].Message;
				logError = attributes[0].LogError;
			}
		}

		public PhotoMixError PhotoMixError
		{
			get { return photoMixError; }
		}

		public bool LogError
		{
			get { return logError; }
		}

		public override string ToString()
		{
			string output = string.Format("PhotoMixException\r\nPhotoMixError: {0}\r\n", PhotoMixError);
			if (context != null)
				output += string.Format("Context: {0}\r\n", context);
			if (partnerErrorCode != null)
				output += string.Format("PartnerErrorCode: {0}\r\n", partnerErrorCode);
			if (partnerErrorDetail != null)
				output += string.Format("PartnerErrorDetail: {0}\r\n", partnerErrorDetail);
			if (message!= null)
				output += string.Format("Message: {0}\r\n", message);
			output += string.Format("Stack Trace:\r\n{0}", StackTrace);
			return output;
		}

		public SoapException ToSoapException(SoapMessage soapMessage, bool hidePrivateDetails)
		{
			// Build the detail element of the SOAP fault.
			System.Xml.XmlDocument document = new System.Xml.XmlDocument();
			System.Xml.XmlNode node = document.CreateNode(XmlNodeType.Element, SoapException.DetailElementName.Name, SoapException.DetailElementName.Namespace);

			// Create node for Error Code
			System.Xml.XmlNode code = document.CreateNode(XmlNodeType.Element, "ErrorCode", null);
			code.InnerText = ((int)PhotoMixError).ToString();
			node.AppendChild(code);

			// Create node for Error Enum
			System.Xml.XmlNode codeEnum = document.CreateNode(XmlNodeType.Element, "ErrorEnum", null);
			codeEnum.InnerText = PhotoMixError.ToString();
			node.AppendChild(codeEnum);


			// Create node for Partner Error Code
			if (!string.IsNullOrEmpty(partnerErrorCode))
			{
				System.Xml.XmlNode partnerErrorCodeNode = document.CreateNode(XmlNodeType.Element, "PartnerErrorCode", null);
				partnerErrorCodeNode.InnerText = partnerErrorCode;
				node.AppendChild(partnerErrorCodeNode);
			}

			// Create node for Partner Error Detail
			if (!string.IsNullOrEmpty(partnerErrorDetail) && !hidePrivateDetails)
			{
				System.Xml.XmlNode partnerErrorDetailNode = document.CreateNode(XmlNodeType.Element, "PartnerErrorDetail", null);
				partnerErrorDetailNode.InnerText = partnerErrorDetail;
				node.AppendChild(partnerErrorDetailNode);
			}

			return new SoapException(
				this.message == null ? PhotoMixError.ToString() : string.Format("{0}: {1}", PhotoMixError.ToString(), this.message),
				SoapException.ClientFaultCode,
				soapMessage == null ? "" : soapMessage.Url,
				node);
		}

	}
}
