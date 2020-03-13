using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Collections;
using System.Net;
using System.Web;

namespace PhotoStuff
{           
    public class DavItem
    {
        public string strName;
        public string strDisplayName;
        public bool fIsCollection;
        public bool fIsHidden;
        public int iContentLength;
        public string strContentType;
        public DateTime dtCreated;
        public DateTime dtModified;
    }    
        
    public class DavUtil
    {    
        public static ArrayList DAVGetData(WebUtil objWebUtil, string strRoot)
        {
            ArrayList objResult = new ArrayList();
            
            if (strRoot == "" || strRoot == null)
                strRoot = "http://groups.msn.com";
                
            // Build the query.
    		string folderQuery = null;
    		folderQuery += "<?xml version='1.0' encoding='UTF-8' ?>";
    		folderQuery += "<a:propfind xmlns:a='DAV:' xmlns:b='urn:schemas-microsoft-com:datatypes'>";
    		folderQuery += "<a:prop>";
    		folderQuery += "<a:name/>";
    		folderQuery += "<a:parentname/>";
    		folderQuery += "<a:href/>";
    		folderQuery += "<a:ishidden/>";
    		folderQuery += "<a:isreadonly/>";
    		folderQuery += "<a:getcontenttype/>";
    		folderQuery += "<a:contentclass/>";
    		folderQuery += "<a:getcontentlanguage/>";
    		folderQuery += "<a:creationdate/>";
    		folderQuery += "<a:lastaccessed/>";
    		folderQuery += "<a:getlastmodified/>";
    		folderQuery += "<a:getcontentlength/>";
    		folderQuery += "<a:iscollection/>";
    		folderQuery += "<a:isstructureddocument/>";
    		folderQuery += "<a:defaultdocument/>";
    		folderQuery += "<a:displayname/>";
    		folderQuery += "<a:isroot/>";
    		folderQuery += "<a:resourcetype/>";
    		folderQuery += "</a:prop>";
    		folderQuery += "</a:propfind>";
    		
    		// Declare locals.
            HttpWebRequest request = null;
            HttpWebResponse response = null;        
            
            // We need to try a few times to hit the server
                        
            // Get the response
            int iTryCount = 2;
            while (iTryCount > 0 && response == null)
            {                
                try
                {
                    // Setup the request
                    request = (HttpWebRequest)WebRequest.Create(strRoot);                        
                    request.Method = "PROPFIND";       
                    request.CookieContainer = new CookieContainer();
                    request.CookieContainer.Add(objWebUtil.cookies);        
                    request.AllowAutoRedirect = false;    
                    request.UserAgent = "Microsoft Data Access Internet Publishing Provider DAV";
                    request.Headers.Add("PROPFIND", folderQuery);
                    request.Headers.Add("Depth", "1");
    		        request.Headers.Add("Translate", "f");            
    		        
    		        // Get the response            
                    response = (HttpWebResponse)request.GetResponse();
                }
                catch (WebException e)
                {
                    if (iTryCount == 1)
                        throw e;                    
                }   
                iTryCount--;             
            }
    
            string strResult = new StreamReader(response.GetResponseStream()).ReadToEnd();
            //Console.WriteLine(strResult);
            
            // Store the response cookies
            objWebUtil.cookies.Add(response.Cookies);
            
            // Get the xml dom from the string
            XmlDocument xmldoc = new XmlDocument();
            
            // For some reason, I can't get multiple namespaces to work in this parser.
            // I'm going to cheat and simply eliminate them from the xml before I parse.
            // Ugly, but works. 
           
            strResult = strResult.Replace("a:", "");
            strResult = strResult.Replace("b:", "");
            
            //Console.WriteLine(strResult);
            
            xmldoc.LoadXml(strResult);                     
    
            
            XmlNodeList nlist = xmldoc.SelectNodes("//response"); 
            
            for (int i = 0; i < nlist.Count; i++)
            {                        
                XmlNode n = nlist[i];
                
                string strName = HttpUtility.UrlDecode(n.SelectSingleNode("href").InnerText);                        
                
                if (strName.ToLower() != strRoot.ToLower())
                {
                    DavItem objDavItem = new DavItem();
    
                    objDavItem.strName = strName;                
                    objDavItem.strDisplayName = n.SelectSingleNode("propstat/prop/displayname").InnerText;
                    objDavItem.fIsCollection = (n.SelectSingleNode("propstat/prop/iscollection").InnerText == "1");
                    objDavItem.fIsHidden = (n.SelectSingleNode("propstat/prop/ishidden").InnerText == "1");
                    objDavItem.iContentLength = Convert.ToInt32(n.SelectSingleNode("propstat/prop/getcontentlength").InnerText);
                    objDavItem.strContentType = n.SelectSingleNode("propstat/prop/getcontenttype").InnerText;                
                    objDavItem.dtCreated = Convert.ToDateTime(n.SelectSingleNode("propstat/prop/creationdate").InnerText);
                    objDavItem.dtModified = Convert.ToDateTime(n.SelectSingleNode("propstat/prop/getlastmodified").InnerText);
                    
                    objResult.Add(objDavItem);                                
                }
    
            }
                       
            
            return objResult;
        }                    
    }
    
    enum Transition { None = 0, Dissolve = 1, Fade = 2, Wipe = 3};
    
    class ImageFile
    {
        public string strName = null;
        public DateTime dtModified = DateTime.MaxValue;
        public DateTime dtCreated = DateTime.MaxValue;
        public long lSize = 0;
        public string strPre = null;
        public string strPost = null;                
    }
    
    public class SizeComparer : IComparer  
    {
        int IComparer.Compare( object x, object y )  
        {                 
            long d = ((ImageFile)y).lSize - ((ImageFile)x).lSize;
            if (d < 0) return -1;
            else if (d == 0) return 0;
            else return 1;     
        }
    }        
    
    public class DateCreatedComparer : IComparer  
    {
        int IComparer.Compare( object x, object y )  
        {
            DateTime dtX = ((ImageFile)x).dtCreated;
            DateTime dtY = ((ImageFile)y).dtCreated;
            
            if (dtX < dtY) return -1;
            else if (dtX == dtY) return 0;
            else return 1;
        }
    }
    
    public class DateModifiedComparer : IComparer  
    {
        int IComparer.Compare( object x, object y )  
        {
            DateTime dtX = ((ImageFile)x).dtModified;
            DateTime dtY = ((ImageFile)y).dtModified;
            
            if (dtX < dtY) return -1;
            else if (dtX == dtY) return 0;
            else return 1;
        }
    }
        
    public class NameComparer : IComparer
    {
        int IComparer.Compare(object x, object y)
        {
            return (new CaseInsensitiveComparer()).Compare( ((ImageFile)x).strName, ((ImageFile)y).strName );
        }
    }     
    
    public class PhotoList    
    {
        Random m_objRandom = new Random(unchecked((int)DateTime.Now.Ticks));         
        
        private WebUtil m_objWebUtil;
        
        public PhotoList(WebUtil objWebUtil)
        {
            m_objWebUtil = objWebUtil;
        }
        
        private int GetDepth(string strDepth)
        {
            if (strDepth == "*" || strDepth == "" || strDepth == null)
                return 0;
            else
                return Convert.ToInt32(strDepth);
        }
        
        private int GetCount(string strCount)
        {
            if (strCount == "*" || strCount == "" || strCount == null)
                return Int32.MaxValue;
            else
                return Convert.ToInt32(strCount);
        }        
        
        private string GetDefaultAttribute(string strAttribute, string strDefaultAttribute)
        {
            if (strAttribute == null || strAttribute == "")
                return strDefaultAttribute;
            else
                return strAttribute;
        }
                        
        public ArrayList Compile(
            string strPhotoList,
            bool fFile
        )
        {
            XmlTextReader xml;
            // Load in the photo list into an xml document                        
            if (fFile)
            {
                xml = new XmlTextReader(strPhotoList);
            }
            else
            {
                StringReader reader = new StringReader(strPhotoList);
                xml = new XmlTextReader(reader); 
            }
            xml.WhitespaceHandling = WhitespaceHandling.None;
            
            // Read the first node and verify that it is a photolist                        
            if (!(xml.Read() && xml.NodeType == XmlNodeType.Element && xml.Name == "photolist"))
                throw new Exception("Invalid photolist: " + strPhotoList);
                
            // Read in the count
            int iCount = GetCount(xml.GetAttribute("count"));
            bool fRepeat = (xml.GetAttribute("repeat") == "1");
            
            // Read in any pre or post transition attribute
            string strPre = xml.GetAttribute("pre");
            string strPost = xml.GetAttribute("post");            
                        
            // Get the sort
            string strSort = xml.GetAttribute("sort");
            bool fDesc = (xml.GetAttribute("desc") == "1");
            
            // Point the reader at the first element in the list
            xml.Read();
            
            // The pointer should point to the end element now....
                            
            // Process the nodes                
            return ProcessNodes(xml, iCount, fRepeat, strSort, fDesc, strPre, strPost);
        }
        
        private void SortImageList(
            ArrayList objWorkingList,            
            string strSort
        )
        {                                                    
            // Deal with a sort as necessary                   
            if (strSort == "size" || strSort == "datec" || strSort == "datem" || strSort == "name")
            {    
                // Comparer
                IComparer myComparer;
                
                // See if we need to fill in size, date data and get the right compare function
                if (strSort == "datec" || strSort == "datem")
                {   
                    // Get the dates if we don't have them
                    for (int i = 0; i < objWorkingList.Count; i++)
                    {
                        ImageFile objImageFile = (ImageFile)objWorkingList[i];
                        if (objImageFile.dtCreated == DateTime.MaxValue)
                            objImageFile.dtCreated = File.GetCreationTime(objImageFile.strName);                            
                        if (objImageFile.dtModified == DateTime.MaxValue)
                            objImageFile.dtModified = File.GetLastWriteTime(objImageFile.strName);                   
                    
                    }                                         
                        
                    // Get the right compare function
                    if (strSort == "size")
                        myComparer = new SizeComparer();
                    else if (strSort == "datec")
                        myComparer = new DateCreatedComparer();
                    else
                        myComparer = new DateModifiedComparer();
                }
                else
                {
                    // All we need is the compare function
                    myComparer = new NameComparer();
                }
                
                // Apply the sort
                objWorkingList.Sort(myComparer);
            }
        }
        
        private ArrayList SelectFromList(
            ArrayList objWorkingList,
            int iCount,
            bool fRepeat,
            string strSort,
            bool fDesc
        )
        {
            
            ArrayList objResultList = new ArrayList();
            
            if (objWorkingList.Count == 0)
                return objWorkingList;
            
            // Adjust the requested count
            if (iCount == Int32.MaxValue)
                iCount = objWorkingList.Count;
            else if (iCount > objWorkingList.Count && !fRepeat)
                iCount = objWorkingList.Count;                
    
            // If random, start filling them in              
            if (strSort == "random")
            {                
                ArrayList objRecycleList = new ArrayList();
                                                    
                while (iCount > 0)
                {                        
                    int i = m_objRandom.Next(objWorkingList.Count);
                    
                    objResultList.Add(objWorkingList[i]);
                    
                    objRecycleList.Add(objWorkingList[i]);
                    
                    objWorkingList.RemoveAt(i);
                    
                    iCount--;
                    
                    if (objWorkingList.Count == 0 && iCount > 0)
                    {
                        ArrayList temp;
                        
                        temp = objRecycleList;
                        objRecycleList = objWorkingList;
                        objWorkingList = temp;
                    }                                                        
                }
            }
            // Pull them from the top of the list
            else if (fDesc)
            {
                int i = objWorkingList.Count - 1;
                while (iCount > 0)
                {
                    objResultList.Add(objWorkingList[i]);
                                                
                    iCount--;
                    
                    if (iCount > 0)
                    {
                        i--;
                        if (i < 0)
                            i = objWorkingList.Count - 1;
                    }
                }                    
            }
            // Pull them from the bottom of the list
            else 
            {
                int i = 0;
                while (iCount > 0)
                {
                    objResultList.Add(objWorkingList[i]);
                    
                    iCount--;
                    
                    if (iCount > 0)
                    {
                        i++;
                        if (i > objWorkingList.Count - 1)
                            i = 0;
                    }
                }
            }                    
            
            return objResultList;
        }
        
        
        public ArrayList AddDirectory(
            string strDirectory,
            int iCount,
            bool fRepeat, 
            string strSort,
            bool fDesc,
            bool fRecurse,
            string strRoot,
            string strPre,
            string strPost,
            string strDirMask,
            string strFileMask,
            int iDepth,
            int iCurrentDepth
            )
        {
            // Get the directory information
            DirectoryInfo di = new DirectoryInfo(strDirectory);
            
            // Read all the directories that match the mask
            DirectoryInfo[] dis = di.GetDirectories(strDirMask);
                                   
            // Get a working list for the result
            ArrayList objWorkingList = new ArrayList(); 

            // Process the files in this directory if there are no subdirectorys, or
            // the root files are request
            // Get the image files in the root (this directory) that match the mask
            if (strRoot == "heavy" || strRoot == "even" || dis.Length == 0)
            {       
                // Get all the files                
                FileInfo[] fis = di.GetFiles(strFileMask);                
                
                for (int i = 0; i < fis.Length; i++)
                {
                    ImageFile objImageFile = new ImageFile();
                    objImageFile.strName = strDirectory + fis[i].Name;
                    objImageFile.strPre = strPre;
                    objImageFile.strPost = strPost;
                    objImageFile.lSize = fis[i].Length;
                    
                    objWorkingList.Add(objImageFile);
                }

                // Evenly weighted files in a directory are filtered down to the same level
                // as those in subdirectories before being included
                if (strRoot == "even")
                {
                    SortImageList(objWorkingList, strSort);
                
                    objWorkingList = SelectFromList(objWorkingList, iCount, fRepeat, strSort, fDesc);
                }                                
            }
            
            // Get all the subdirectories that match the mask
            if (fRecurse)
            {                                
                for (int i = 0; i < dis.Length; i++)
                {
                    ArrayList objDirectoryFiles = AddDirectory(strDirectory + dis[i].Name + "\\", iCount, fRepeat, strSort, fDesc, fRecurse, strRoot, strPre, strPost, strDirMask, strFileMask, iDepth, iCurrentDepth + 1);
                
                    for (int j = 0; j < objDirectoryFiles.Count; j++)
                    {                
                        objWorkingList.Add(objDirectoryFiles[j]);            
                    }
                }            
            }                
            
            // Normalize and sort only those that are deeper than the requested depth
            if (iCurrentDepth >= iDepth)
            {
                SortImageList(objWorkingList, strSort);
            
                return SelectFromList(objWorkingList, iCount, fRepeat, strSort, fDesc);
            }
            else
            {
                return objWorkingList;
            }
        }                                            
        
        public ArrayList AddGroup(
            string strGroup,
            int iCount,
            bool fRepeat, 
            string strSort,
            bool fDesc,
            bool fRecurse,
            string strRoot,
            string strPre,
            string strPost,            
            int iDepth,
            int iCurrentDepth
            )
        {   
            ArrayList objDavData;
            
            // Read the group information
            objDavData = DavUtil.DAVGetData(m_objWebUtil, strGroup);
            
            int fileCount = 0;
            int directoryCount = 0;
            for (int i = 0; i < objDavData.Count; i++)
            {
                DavItem objDavItem = (DavItem)objDavData[i];
                
                if (objDavItem.fIsCollection)
                    directoryCount++;
                    
                if (objDavItem.strContentType == "image/jpeg")
                    fileCount++;
            }
            
            // Get a working list for the result
            ArrayList objWorkingList = new ArrayList();                        

            // Process the files in this directory if there are no subdirectorys, or
            // the root files are request
            if (fileCount > 0 && (strRoot == "heavy" || strRoot == "even" || directoryCount == 0))
            {                       
                
                for (int i = 0; i < objDavData.Count; i++)
                {
                    DavItem objDavItem = (DavItem)objDavData[i];
                    if (objDavItem.strContentType == "image/jpeg")
                    {
                        ImageFile objImageFile = new ImageFile();
                        objImageFile.strName = objDavItem.strName;
                        objImageFile.strPre = strPre;
                        objImageFile.strPost = strPost;
                        objImageFile.lSize = objDavItem.iContentLength;
                        objImageFile.dtCreated = objDavItem.dtCreated;
                        objImageFile.dtModified = objDavItem.dtModified;
                        
                        objWorkingList.Add(objImageFile);
                    }                                        
                }

                // Evenly weighted files in a directory are filtered down to the same level
                // as those in subdirectories before being included
                if (strRoot == "even")
                {
                    SortImageList(objWorkingList, strSort);
                
                    objWorkingList = SelectFromList(objWorkingList, iCount, fRepeat, strSort, fDesc);
                }                                
            }
            
            // Get all the subdirectories that match the mask
            if (fRecurse && directoryCount > 0)
            {                                
                for (int i = 0; i < objDavData.Count; i++)
                {
                    DavItem objDavItem = (DavItem)objDavData[i];
                    if (objDavItem.fIsCollection)
                    {
                        ArrayList objGroupFiles = AddGroup(objDavItem.strName, iCount, fRepeat, strSort, fDesc, fRecurse, strRoot, strPre, strPost, iDepth, iCurrentDepth + 1);
                
                        for (int j = 0; j < objGroupFiles.Count; j++)
                        {                
                            objWorkingList.Add(objGroupFiles[j]);            
                        }
                    }
                }            
            }                
            
            // Normalize and sort only those that are deeper than the requested depth
            if (iCurrentDepth >= iDepth)
            {
                SortImageList(objWorkingList, strSort);
            
                return SelectFromList(objWorkingList, iCount, fRepeat, strSort, fDesc);
            }
            else
            {
                return objWorkingList;
            }
                        
        }                                            
        
        private ArrayList ProcessNodes(
            XmlTextReader xml,
            int iCount,
            bool fRepeat,
            string strSort,
            bool fDesc,
            string strPre,
            string strPost)
        {            
            ArrayList objWorkingList = new ArrayList();
                
            while (xml.NodeType != XmlNodeType.EndElement)
            {
                if (xml.NodeType != XmlNodeType.Element)
                    throw new Exception("Malformed photolist");        
                    
                if (xml.Name == "imagefile")
                {
                    ImageFile objImageFile;
                    string strName;
                    //Datetime dtModified;
                    //Datetime dtCreated;
                    
                    // Read in any pre or post transition attribute
                    string strImageFilePre = GetDefaultAttribute(xml.GetAttribute("pre"), strPre);
                    string strImageFilePost = GetDefaultAttribute(xml.GetAttribute("post"), strPost);
                    
                    // Advance the parser to read the name
                    xml.Read();
                    strName = xml.Value;
                    
                    // Read in some file information
                                                                                                                        
                    // Create the object for the image file
                    objImageFile = new ImageFile();
                    objImageFile.strName = strName;
                    //objImageFile.dtCreated = File.GetCreationTime(strName);
                    //objImageFile.dtModified = File.GetLastWriteTime(strName);
                    //objImageFile.lSize = (new FileInfo(strName)).Length;
                    
                    objImageFile.strPre = strImageFilePre;
                    objImageFile.strPost = strImageFilePost;
                    
                    // Add to the list                
                    objWorkingList.Add(objImageFile);                                        
                    
                    // Advance the parser to read the end element
                    xml.Read();                    
                }
                else if (xml.Name == "photolist" || xml.Name == "photolistfile")
                {                                        
                    // Get the count, sorting attributes
                    string strPLPre = GetDefaultAttribute(xml.GetAttribute("pre"), strPre);
                    string strPLPost = GetDefaultAttribute(xml.GetAttribute("post"), strPost);
                    int iPLCount = GetCount(xml.GetAttribute("count"));
                    string strPLSort = xml.GetAttribute("sort");
                    bool fPLDesc = (xml.GetAttribute("desc") == "1");
                    bool fPLRepeat = (xml.GetAttribute("repeat") == "1");
                
                    // Get the xml for the embedded data
                    XmlTextReader objXmlEmbedded;
                    if (xml.Name == "photolist")
                    {
                        objXmlEmbedded = xml;
                    }
                    else
                    {
                        // Advance to the file name
                        xml.Read();           
                             
                        // Read in the items        
                        objXmlEmbedded = new XmlTextReader(xml.Value);
                        objXmlEmbedded.WhitespaceHandling = WhitespaceHandling.None;
                        
                        // Read the first node and verify that it is a photolist                        
                        if (!(objXmlEmbedded.Read() && objXmlEmbedded.NodeType == XmlNodeType.Element && objXmlEmbedded.Name == "photolist"))
                            throw new Exception("Invalid photolist");
                            
                        // Advance the parser to point at the end element
                        xml.Read();
                    }
                    
                    // Read to the first element in the list
                    objXmlEmbedded.Read();
                                        
                    // Get the items in the embedded photo list            
                    ArrayList objPhotoListItems = ProcessNodes(objXmlEmbedded, iPLCount, fPLRepeat, strPLSort, fPLDesc, strPLPre, strPLPost);
                    
                    
                    // Add them (in order) to the working list
                    for (int i = 0; i < objPhotoListItems.Count; i++)                        
                        objWorkingList.Add(objPhotoListItems[i]);
                }                                                                 
                else if (xml.Name == "directory")
                {                 
                    string strDirPre = GetDefaultAttribute(xml.GetAttribute("pre"), strPre);
                    string strDirPost = GetDefaultAttribute(xml.GetAttribute("pre"), strPost);                    
                    string strDirDirMask = GetDefaultAttribute(xml.GetAttribute("dirmask"), "*");
                    string strDirFileMask = GetDefaultAttribute(xml.GetAttribute("filemask"), "*.jpg");
                    bool fDirRecurse = (xml.GetAttribute("recurse") == "1");                    
                    int iDirCount = GetCount(xml.GetAttribute("count"));                                        
                    bool fDirRepeat = (xml.GetAttribute("repeat") == "1");                    
                    string strDirSort = GetDefaultAttribute(xml.GetAttribute("sort"), "listed");
                    bool fDirDesc = (xml.GetAttribute("desc") == "1");                                                                                
                    string strDirRoot = GetDefaultAttribute(xml.GetAttribute("root"), "even");
                    int iDirDepth = GetDepth(xml.GetAttribute("depth"));
                    
                    // Read to the directory name
                    xml.Read();                    
                    string strDirectory = xml.Value;                                        
                    
                    // Create a list for the items
                    ArrayList objDirImageList = AddDirectory(strDirectory, iDirCount, fDirRepeat, strDirSort, fDirDesc, fDirRecurse, strDirRoot, strDirPre, strDirPost, strDirDirMask, strDirFileMask, iDirDepth, 0);
                                        
                    // Add them (in order) to the working list
                    for (int i = 0; i < objDirImageList.Count; i++)
                        objWorkingList.Add(objDirImageList[i]);                    
                    
                    // Advance to the close tag
                    xml.Read();
                }
                else if (xml.Name == "group")
                {
                    string strGroupPre = GetDefaultAttribute(xml.GetAttribute("pre"), strPre);            
                    string strGroupPost = GetDefaultAttribute(xml.GetAttribute("pre"), strPost);                                        
                    bool fGroupRecurse = (xml.GetAttribute("recurse") == "1");                    
                    int iGroupCount = GetCount(xml.GetAttribute("count"));                                        
                    bool fGroupRepeat = (xml.GetAttribute("repeat") == "1");                    
                    string strGroupSort = GetDefaultAttribute(xml.GetAttribute("sort"), "listed");
                    bool fGroupDesc = (xml.GetAttribute("desc") == "1");                                                                                
                    string strGroupRoot = GetDefaultAttribute(xml.GetAttribute("root"), "even");
                    int iGroupDepth = GetDepth(xml.GetAttribute("depth"));
                    
                    // Read to the group name
                    xml.Read();                    
                    string strGroup = xml.Value;                                        
                    
                    // Create a list for the items
                    ArrayList objDirImageList = AddGroup(strGroup, iGroupCount, fGroupRepeat, strGroupSort, fGroupDesc, fGroupRecurse, strGroupRoot, strGroupPre, strGroupPost, iGroupDepth, 0);
                                        
                    // Add them (in order) to the working list
                    for (int i = 0; i < objDirImageList.Count; i++)
                        objWorkingList.Add(objDirImageList[i]);                    
                    
                    // Advance to the close tag
                    xml.Read();    
                }
                else
                {
                    throw new Exception("Unknown command: " + xml.Name);
                }
                
                // Advance the parser to the next element
                xml.Read();
            }                            
        
            // Sort the list
            SortImageList(objWorkingList, strSort);
            
            // Select from the list
            return SelectFromList(objWorkingList, iCount, fRepeat, strSort, fDesc);
        }                        

    }            
}