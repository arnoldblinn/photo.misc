using System;
using System.Text;
using System.IO;
using System.Net;
using System.Xml;
using System.Web;


namespace PhotoStuff
{
    public class WebUtil
    {        
        public CookieCollection cookies = new CookieCollection();
        
        public void FileWebData(string strURI, string strFile)
        {
            // Declare locals.
            HttpWebRequest request = null;
            HttpWebResponse response = null;
    
            // Create the web request
            request = (HttpWebRequest)WebRequest.Create(strURI);
            request.AllowAutoRedirect = true;
            request.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.1.4322)";
            
            // Add the existing cookies
            request.CookieContainer = new CookieContainer();
            request.CookieContainer.Add(cookies);
                        
            // Get the response
            response = (HttpWebResponse)request.GetResponse();    
            Stream objStream = response.GetResponseStream();
            BinaryReader objBinaryReader = new BinaryReader(objStream);            
            
            FileInfo fi1 = new FileInfo(strFile);
            if (fi1.Exists)
                fi1.Delete();
                
            FileStream fs = new FileStream(strFile, FileMode.CreateNew);    
            BinaryWriter objBinaryWriter = new BinaryWriter(fs);
            
            byte[] buf = new byte[256];
            int count = objBinaryReader.Read( buf, 0, 256 );
   
            while (count > 0) 
            {                
                objBinaryWriter.Write(buf, 0, count);
                count = objBinaryReader.Read(buf, 0, 256);
            }
            
            cookies.Add(response.Cookies);
            
            objBinaryWriter.Close();
            
        }
        
        public string PostWebData(string strURI, string strBody)
        {
            // Declare locals.
            HttpWebRequest request = null;
            HttpWebResponse response = null;
            ASCIIEncoding encoding=new ASCIIEncoding();
            byte[]  byte1=encoding.GetBytes(strBody);
    
            // Setup the request
            request = (HttpWebRequest)WebRequest.Create(strURI);                        
            request.Method = "POST";
            request.ContentLength=strBody.Length;        
            request.CookieContainer = new CookieContainer();
            request.CookieContainer.Add(cookies);
            request.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.1.4322)";                        
            
            Stream newStream = request.GetRequestStream();
            request.AllowAutoRedirect = true;
            newStream.Write(byte1,0,byte1.Length);
                    
            // Get the response
            response = (HttpWebResponse)request.GetResponse();
    
            string strResult = new StreamReader(response.GetResponseStream()).ReadToEnd();
            //Console.WriteLine(strResult);
            
            // Store the response cookies
            cookies.Add(response.Cookies);
            
            return strResult;
        }
        
        public string GetWebData(string strURI)
        {
            // Declare locals.
            HttpWebRequest request = null;
            HttpWebResponse response = null;
    
            request = (HttpWebRequest)WebRequest.Create(strURI);
            request.AllowAutoRedirect = true;
            request.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.1.4322)";
            request.CookieContainer = new CookieContainer();
            request.CookieContainer.Add(cookies);                        
            
            response = (HttpWebResponse)request.GetResponse();
    
            string strResult = new StreamReader(response.GetResponseStream()).ReadToEnd();
            //Console.WriteLine(strResult);
            
            cookies.Add(response.Cookies);
            
            return strResult;
        }
            
        public bool PPSignIn(string strName, string strPassword)
        {
            try
            {
                // Set up the login stuff.  Note that the uri technically changes per passport domain        
                string strFormData = "<LoginRequest><ClientInfo name=\"MSN6\" version=\"1.35\"/><User><SignInName>" + strName + "</SignInName><Password>" + strPassword + "</Password><SavePassword>false</SavePassword></User><DAOption>0</DAOption><TargetOption>0</TargetOption></LoginRequest>";        
                string uriString = "https://loginnet.passport.com/ppsecure/clientpost.srf?id=6528&RU=http%3A%2F%2Flogin.msn.com%2Fpassport%2F";        
    
                // Login to passport                
                string strResponse = PostWebData(uriString, strFormData);                
            
                //Console.WriteLine("\nResponse received was : " + strResponse);        
            
                // Parse out the response and redirect url
                StringReader reader = new StringReader(strResponse);
                XmlTextReader xml = new XmlTextReader(reader); 
                xml.WhitespaceHandling = WhitespaceHandling.None;
                            
                if (xml.Read() && xml.NodeType == XmlNodeType.Element && xml.Name == "LoginResponse" && xml.GetAttribute("Success") == "true")
                {
                    string strRedirect = "";
                    while (xml.Read())
                    {
                        if (xml.NodeType == XmlNodeType.Element && xml.Name == "Redirect")
                        {
                            xml.Read();
                            strRedirect = xml.Value;
                            break;
                        }            
                    }
                    
                    if (strRedirect != "")
                    {
                        strResponse = GetWebData(strRedirect);
                        //Console.WriteLine("\nResponse received was : " + strResponse);
                        
                        return true;
                    }
                }
                
                return false;
            }
            catch
            {
                return false;
            }
        }                
    }
}        