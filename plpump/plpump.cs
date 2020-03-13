using System;
using System.Net;
using System.Collections;
using System.IO;
using System.Text;
using System.Xml;

using PhotoStuff;

public class Test
{
    public static void Main(string[] args)
    {               

        bool fBadCommand = false;                
                        
        // Parse out the command line arguments
        string strPhotoListFile = null;        
        string strTransformFile = null;        
        string strOutputDirectory = null;
        
        string strUserName = null;
        string strPassword = null;
        
        bool fVerbose = false;
        bool fDelete = false;
        
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "-v")
                fVerbose = true;
            else if (args[i] == "-d")
                fDelete = true;
            else if (args[i].Length > 3)
            {
                string strCommand = args[i].Substring(0, 3).ToLower();
                string strValue = args[i].Substring(3, args[i].Length - 3);
                
                if (strCommand == "-u:")
                    strUserName = strValue;
                else if (strCommand == "-p:")
                    strPassword = strValue;
                else if (strCommand == "-f:")
                    strPhotoListFile = strValue;
                else if (strCommand == "-t:")
                    strTransformFile = strValue;
                else if (strCommand == "-o:")
                    strOutputDirectory = strValue;                
                else
                    fBadCommand = true;
            }
            else
            {
                fBadCommand = true;
            }
        }
        
        
        if (fBadCommand || strPhotoListFile == null || strTransformFile == null || strOutputDirectory == null)
        {
            Console.WriteLine("plpump -f:photolistfile -t:transformfile -o:outputdirectory [-u:username] [-p:password] [-v] [-d]");
            
            return;
        }

        if (fVerbose) Console.WriteLine("Processing Images...");
                
        ArrayList objResult;
        WebUtil objWebUtil = new WebUtil();
    
        // Sign into Passport if passed a user name and password
        if (strUserName != null && strPassword != null)
        {
            if (fVerbose) Console.WriteLine("Signing into Passport");
            if (!objWebUtil.PPSignIn(strUserName, strPassword))
            {
                if (fVerbose) Console.WriteLine("Can't sign in");
                return;
            }
            if (fVerbose) Console.WriteLine("Sign in successful");
        }                                                                    
        
        // Create the photo list and compile it
        if (fVerbose) Console.WriteLine("Building photolist....");
        PhotoList objPhotoList = new PhotoList(objWebUtil);                                      
        objResult = objPhotoList.Compile(strPhotoListFile, true);            
        if (fVerbose) Console.WriteLine("Photolist complete.");                
        
        // Delete the files in the target directory as necessary
        if (fDelete)
        {
            string strTemplate = strOutputDirectory + "*.jpg";
            if (fVerbose) Console.WriteLine("Deleting files matching " + strTemplate);
            string[] files = Directory.GetFiles(strOutputDirectory, "*.jpg");
            for (int i = 0; i < files.Length; i++)
            {
                if (fVerbose) Console.WriteLine("Deleting file " + files[i]);
                System.IO.File.Delete(files[i]);
            }
        }
        
        // Create a transform
        XmlDocument objXmlDocument = new XmlDocument();
        objXmlDocument.Load(strTransformFile);        
        
        for (int i = 0; i < objResult.Count; i++)
        {            
            ImageFile objImageFile = (ImageFile)objResult[i];
            string strName = objImageFile.strName;               
            string strFileName = DateTime.Now.ToString("yyMMddHHmmssfffffff");                                             
            string strOutput = strOutputDirectory + strFileName + ".jpg";
            
            if (fVerbose) Console.WriteLine("Processing Item " + i);
            
            if (strName.Length > 7 && strName.Substring(0, 7).ToLower() == "http://")
            {
                string strTempName = Path.GetTempFileName().Replace("tmp", "jpg");
                
                if (fVerbose) Console.WriteLine("Saving " + strName + " to temp file " + strTempName);                
                objWebUtil.FileWebData(strName, strTempName);
                
                if (fVerbose) Console.WriteLine("Transforming " + strTempName + " into " + strOutput);
                ImageUtil.SaveImage(strTempName, strOutput, objXmlDocument);
                
                if (fVerbose) Console.WriteLine("Deleting temp file " + strTempName);
                System.IO.File.Delete(strTempName);        
            }   
            else
            {
                if (fVerbose) Console.WriteLine("Transforming " + strName + " into " + strOutput);
                ImageUtil.SaveImage(strName, strOutput, objXmlDocument);
            }                                 
            
        }
        
        if (fVerbose)
        {
            Console.WriteLine("Press enter to continue");
            Console.ReadLine();
        }
    }    
}    