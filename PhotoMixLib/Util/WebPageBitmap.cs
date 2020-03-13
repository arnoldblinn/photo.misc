using System;
using System.IO;
using System.Drawing;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Xml;
using System.Xml.Xsl;
using mshtml;
using Msn.Framework;

namespace Msn.PhotoMix
{
    public delegate void WebPageBitmapCallback(int lineCount);

    // <summary>
    //      Renders a webpage using a WebBrowser control at the specified browser dimension
    // </summary>
    public class WebPageBitmap
    {
        private String url;
        private MemoryStream stream;
        private int height;
        private int width;
        bool renderImage = true;
        Bitmap capture = null;
        List<string> imageSrc = null;
        bool docLoaded = false;

        public WebPageBitmap(String url, MemoryStream stream, int width, int height, bool renderImage)
        {
            this.height = height;
            this.width = width;
            this.url = url;
            this.stream = stream;
            this.renderImage = renderImage;
        }

        public void Dispose()
        {
            if (capture != null)
                this.capture.Dispose();
        }

        //
        // Process the contents of a fully loaded WebBrowser control
        //
        private void WebBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            WebBrowser webBrowser = (WebBrowser)sender;

            try
            {
                webBrowser.ClientSize = new Size(this.width, this.height);
                webBrowser.ScrollBarsEnabled = false;
                if (this.renderImage)
                {
                    this.capture = new Bitmap(webBrowser.Bounds.Width, webBrowser.Bounds.Height);
                    webBrowser.BringToFront();
                    webBrowser.DrawToBitmap(capture, webBrowser.Bounds);
                }
                else
                {
                    this.imageSrc = new List<string>();
                    HtmlElementCollection images = webBrowser.Document.Images;
                    for (int i = 0; i < images.Count; i++)
                    {
                        string src = ((mshtml.HTMLImgClass)webBrowser.Document.Images[i].DomElement).src;
                        if (!this.imageSrc.Contains(src))
                            this.imageSrc.Add(src);
                    }
                }
            }
            catch (Exception)
            {
            }

            webBrowser.Dispose();
            docLoaded = true;
        }

        //
        //  Load a WebBrowser control with the requested contents (url or stream)
        //
        private void Load()
        {
            WebBrowser browser = null;

            try
            {
                browser = new WebBrowser();
                browser.ScrollBarsEnabled = false;
                browser.Size = new Size(width, height);
                browser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(WebBrowser_DocumentCompleted);

                if (this.url != null)
                    browser.Navigate(this.url);
                else
                {
                    stream.Seek(0, 0);
                    browser.DocumentStream = this.stream;
                }

                while (!docLoaded)
                {
                    Application.DoEvents();
                }
            }
            catch (Exception e)
            {
                //$ Tracing doesn't work correctly in a multithreaded environment
                /*
                TraceEx.EnableLogFileTraceOutput("trace.txt", Config.GetSetting("TraceDirectory"));
                TraceEx.WriteLine("Error = " + e.Message);
                TraceEx.CloseTraceOutput();
                */
                ErrorLog.WriteEntry(e);
                if (capture != null)
                {
                    this.capture.Dispose();
                    capture = null;
                }
                if (browser != null)
                    try
                    {
                        browser.Dispose();
                    }
                    catch (Exception)
                    {
                    }
            }

            return;
        }

        //
        // Run the WebBrowser control in a single apartment thread
        //
        static private void Invoke(WebPageBitmap wpb)
        {
            Thread thread = new Thread(wpb.Load);
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        public Bitmap Bitmap
        {
            get { return this.capture; }
        }

        public List<string> ImageSrc
        {
            get { return this.imageSrc; }
        }

        //
        // Create a bitmap from a url
        //
        static public Bitmap Fetch(String url, int width, int height)
        {
            WebPageBitmap wpb = new WebPageBitmap(url, null, width, height, true);

            Invoke(wpb);

            return wpb.Bitmap;
        }

        //
        // Create a bitmap from an html stream
        //
        static public Bitmap LoadStream(MemoryStream stream, int width, int height)
        {
            WebPageBitmap wpb = new WebPageBitmap(null, stream, width, height, true);

            Invoke(wpb);

            return wpb.Bitmap;
        }

        //
        // Fetch the list of image sources from the url
        //
        static public List<string> LoadDocumentImages(String url)
        {
            WebPageBitmap wpb = new WebPageBitmap(url, null, 0, 0, false);

            Invoke(wpb);

            return wpb.ImageSrc;
        }

        //
        // Build a bitmap from an xsl file and an xml document
        //
        static public Bitmap LoadXsl(string fileName, XmlDocument doc, int width, int height)
        {
            // Load the style sheet.
            XslCompiledTransform xslt = new XslCompiledTransform();
            xslt.Load(fileName);

            // Execute the transform and output the results to a stream.
            MemoryStream stream = new MemoryStream();
            XmlWriter results = XmlWriter.Create(stream);
            xslt.Transform(doc, results);
            results.Close();

            return WebPageBitmap.LoadStream(stream, width, height);
        }

    }    
}
