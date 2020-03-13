using System;
using System.IO;
using System.Drawing;
using System.Xml;
using System.Collections;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Net;

using Msn.Framework;

//
// Contains one-off functions that have no other home
//

namespace Msn.PhotoMix
{
    public class MiscUtil
    {
        //
        // TimeIndex - Return a string that can be appended to an URL to make it unique so
        //             that it won't get cached
        //
        static public string TimeIndex()
        {
            return "ti=" + DateTime.Now.Ticks / 10000000;
        }

        //
        // Returns an advertising image that can be integrated with an information image
        //
        static public Bitmap GetAd()
        {
            string url = "http://" + Config.GetSetting("HostSite") + "/DemoAd.aspx";  //$ Temporary

            return ImageUtil.LoadImageFromUrl(url);
        }

        //
        // Returns whether a file exists based on a time-to-live parameter
        //
        static public bool TTLFileExists(string fileName, int ttl)
        {
            try
            {
                DateTime createDate = File.GetLastWriteTime(fileName);
                if (createDate.AddSeconds(ttl) >= DateTime.Now)
                    return true;
            }
            catch (Exception)
            {
            }

            return false;
        }

        //
        // Returns a string suitably escaped for javascript
        //
        static public string EscapeJavascriptString(string value)
        {
            return value.Replace("'", "\\'");
        }

        //
        // Appends a parameter to the end of a query string on an URL
        //
        static public string AppendQueryStringParameter(string url, string parameter)
        {
            if (url.Contains("?"))
                return url + "&" + parameter;
            else
                return url + "?" + parameter;
        }

        //
        // Get the Facebook API key
        //
        static public string GetFacebookApiKey()
        {
            return Config.GetSetting("FacebookApiKey");
        }

        //
        // Get the Facebook secret
        //
        static public string GetFacebookSecret()
        {
            return Config.GetSetting("FacebookSecret");
        }

        //
        // Calculate a signature for Facebook
        //
        static public string FacebookSignature(ArrayList args)
        {
            StringBuilder sb = new StringBuilder();
            args.Sort();

            foreach (string arg in args)
                sb.Append(arg);

            sb.Append(MiscUtil.GetFacebookSecret());

            StringBuilder retval = new StringBuilder();

            try
            {
                MD5 md = MD5.Create();
                byte[] hash = md.ComputeHash(Encoding.Default.GetBytes(sb.ToString().Trim()));

                foreach (byte b in hash)
                    retval.Append(String.Format("{0:x2}", b));
            }
            catch (ArgumentException e)
            {
                throw new Exception("MD5 encoding error", e);
            }

            return retval.ToString();
        }

        static public XmlDocumentEx CallFacebook(ArrayList args)
        {
            WebRequest request;
            string serverUrl = "https://api.facebook.com";

            try
            {
                request = WebRequest.Create(serverUrl + "/restserver.php");
            }
            catch (SecurityException e)
            {
                throw new Exception("Facebook connection exception", e);
            }

            request.ContentType = "application/x-www-form-urlencoded";
            request.Method = "POST";

            args.Add(String.Format("sig={0}", FacebookSignature(args)));

            StringBuilder sb = new StringBuilder();
            foreach (string arg in args)
            {
                if (sb.Length != 0)
                    sb.Append("&");
                sb.Append(arg);
            }

            byte[] buffer = Encoding.ASCII.GetBytes(sb.ToString());
            request.ContentLength = buffer.Length;

            Stream stream = request.GetRequestStream();
            try
            {
                stream.Write(buffer, 0, buffer.Length);
            }
            catch (IOException e)
            {
                throw new Exception("Facebook POST exception", e);
            }
            finally
            {
                stream.Close();
            }

            WebResponse response = request.GetResponse();
            StreamReader streamReader = new StreamReader(response.GetResponseStream());
            XmlDocumentEx xmlDoc = new XmlDocumentEx();

            try
            {
                xmlDoc.Load(streamReader);
            }
            catch (XmlException e)
            {
                throw new Exception("Facebook call exception", e);
            }

            return xmlDoc;
        }
    }
}