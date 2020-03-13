using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Xml;

using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using Msn.Framework;

namespace Msn.PhotoMix.SlideShow
{       
    public class CompiledWeather
    {
        // Language for this feed
        private string language;

        // Location of the user
        private string location;

        // Weather data from feed
        private XmlNode weatherData;

        // Date that the feed was referenced in a compile
        private DateTime compiledDate;

        // Date that the image was generated
        private DateTime imageGeneratedDate;

        // Date that the data for the image was last fetched
        private DateTime fetchDataDate;

        // Time to live for weather
        static private int weatherFetchTTL = Convert.ToInt32(Config.GetSetting("WeatherFetchTTL"));
        static private int weatherImageTTL = Convert.ToInt32(Config.GetSetting("WeatherImageTTL"));

        public CompiledWeather()
        {            
        }

        public string Language
        {
            get { return this.language; }
        }

        public string Location
        {
            get { return this.location; }
        }

        public XmlNode WeatherData
        {
            get { return this.weatherData; }
        }

        public DateTime FetchDataDate
        {
            get { return this.fetchDataDate; }
        }

        //
        // FetchWeatherData
        //
        // Gets the weather forecast data in XML format for the selected language/location
        //
        static public XmlNode FetchWeatherData(string language, string location)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load("http://weather.msn.com/weatherdata.aspx?wealocations=" + location);

            return xmlDoc.SelectSingleNode("weatherdata");
        }

        //
        // Load
        //
        // Loads the object from the database, and if it doesn't exist, creates it
        static public CompiledWeather Load(string language, string location, bool forImage, DateTime dateContext, bool bypassCaches)
        {
            CompiledWeather compiledWeather = null;
                        
            using (PhotoMixQuery query = new PhotoMixQuery("SelectCompiledWeather"))
            {
                query.Parameters.Add("@Language", SqlDbType.Char).Value = language;
                query.Parameters.Add("@Location", SqlDbType.VarChar).Value = location;
                if (forImage)
                    query.Parameters.Add("@ImageGeneratedDate", SqlDbType.DateTime).Value = dateContext;
                else
                    query.Parameters.Add("@CompiledDate", SqlDbType.DateTime).Value = dateContext;                

                if (query.Reader.Read())
                {
                    compiledWeather = new CompiledWeather();
                    compiledWeather.language = language;
                    compiledWeather.location = location;

                    compiledWeather.compiledDate = query.Reader.IsDBNull(3) ? DateTime.MinValue : query.Reader.GetDateTime(3);
                    compiledWeather.imageGeneratedDate = query.Reader.IsDBNull(4) ? DateTime.MinValue : query.Reader.GetDateTime(4);

                    if (!query.Reader.IsDBNull(2))
                    {
                        XmlDocument xmlDoc = new XmlDocument();

                        compiledWeather.weatherData = xmlDoc.CreateNode(XmlNodeType.Element, "weatherdata", "");
                        compiledWeather.weatherData.InnerXml = query.Reader.IsDBNull(2) ? null : query.Reader.GetString(2);
                        compiledWeather.fetchDataDate = query.Reader.IsDBNull(5) ? DateTime.MinValue : query.Reader.GetDateTime(5);
                    }
                    else
                    {
                        compiledWeather.weatherData = null;
                        compiledWeather.fetchDataDate = DateTime.MinValue;
                    }                                        
                }
            }

            if (compiledWeather != null &&
                (bypassCaches || compiledWeather.weatherData == null || compiledWeather.fetchDataDate.AddMinutes(CompiledWeather.weatherFetchTTL) < dateContext))
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load("http://weather.msn.com/weatherdata.aspx?wealocations=" + location);
                XmlNode weatherData = xmlDoc.SelectSingleNode("weatherdata");

                string sql = "update CompiledWeather " +
                        "set WeatherData = @WeatherData, FetchDataDate = @FetchDataDate " +
                        "where Language = @Language and Location = @Location";
                using (PhotoMixQuery query2 = new PhotoMixQuery(sql, CommandType.Text))
                {
                    query2.Parameters.Add("@Language", SqlDbType.Char).Value = language;
                    query2.Parameters.Add("@Location", SqlDbType.VarChar).Value = location;
                    query2.Parameters.Add("@WeatherData", SqlDbType.Text).Value = String.IsNullOrEmpty(weatherData.InnerXml) ? (Object)DBNull.Value : (Object)weatherData.InnerXml;
                    query2.Parameters.Add("@FetchDataDate", SqlDbType.DateTime).Value = dateContext;
                    query2.Execute();
                }

                compiledWeather.weatherData = weatherData;
                compiledWeather.fetchDataDate = dateContext;
            }

            return compiledWeather;
        }
        
        //
        // LoadForCompile
        //
        // Given data unique to a text rss feed, will either create a new compiled reference
        // or update an existing compile reference
        //
        static public CompiledWeather LoadForCompile(string language, string location, DateTime dateContext)
        {
            return Load(language, location, false, dateContext, false);
        }
        
        //
        // LoadForImage
        //
        // Given a language and location, will 
        //
        static public CompiledWeather LoadForImage(string language, string location, bool bypassCaches)
        {
            return Load(language, location, true, DateTime.Now, bypassCaches);            
        }

        //
        // GetImageFileName
        //
        // Will get the image file name representing the item
        //
        public static string GetImageFileName(string language, string location, string templateDirectory, string weatherImageDirectory, SlideShowImageSize imageSize, bool bypassCaches)
        {
            string fileName = ImageUtil.GetCompiledImageDirectory("Weather") + language + "_" + location + "_" + ((int)imageSize).ToString() + ".jpg";
            if (!bypassCaches && MiscUtil.TTLFileExists(fileName, CompiledWeather.weatherImageTTL))
            {
                return fileName;
            }
            else
            {
                CompiledWeather compiledWeather = CompiledWeather.LoadForImage(language, location, bypassCaches);

                Bitmap bitmap = compiledWeather.GenerateImage(fileName, templateDirectory, weatherImageDirectory, imageSize);

                return fileName;
            }
        }

        //
        // GenerateImage
        //
        // This will generate an image for the weather data
        //
        private Bitmap GenerateImage(string fileName, string templateDirectory, string weatherImageDirectory, SlideShowImageSize imageSize)
        {
            int width = SlideShow.slideShowImageWidths[(int)imageSize];
            int height = SlideShow.slideShowImageHeights[(int)imageSize];            

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.AppendChild(xmlDoc.ImportNode(this.weatherData, true));
            ImageUtil.AddXmlElement(xmlDoc.DocumentElement, "weatherimagedirectory", weatherImageDirectory);
            ImageUtil.AddXmlElement(xmlDoc.DocumentElement, "adurl", Config.GetSetting("AdUrl"));
            ImageUtil.AddXmlElement(xmlDoc.DocumentElement, "width", width.ToString());
            ImageUtil.AddXmlElement(xmlDoc.DocumentElement, "height", height.ToString());
            Bitmap bitmap = WebPageBitmap.LoadXsl(templateDirectory + "Weather.xsl", xmlDoc, width, height);
            ImageUtil.SaveJpeg(fileName, bitmap, 100);

            return bitmap;
        }
    }
}
