using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Net;
using System.IO;
using System.Xml;
using System.Text;

namespace PhotoStuff
{
    public class ConvMatrix
    {
        public int TopLeft = 0, TopMid = 0, TopRight = 0;
        public int MidLeft = 0, Pixel = 1, MidRight = 0;
        public int BottomLeft = 0, BottomMid = 0, BottomRight = 0;
        public int Factor = 1;
        public int Offset = 0;
        public void SetAll(int nVal)
        {
            TopLeft = TopMid = TopRight = MidLeft = Pixel = MidRight = BottomLeft = BottomMid = BottomRight = nVal;
        }
    }
    
    public class Conv
    {
        public static bool Invert(Bitmap b)
		{
			// GDI+ still lies to us - the return format is BGR, NOT RGB.
			BitmapData bmData = b.LockBits(new Rectangle(0, 0, b.Width, b.Height), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);

			int stride = bmData.Stride;
			System.IntPtr Scan0 = bmData.Scan0;

			unsafe
			{
				byte * p = (byte *)(void *)Scan0;

				int nOffset = stride - b.Width*3;
				int nWidth = b.Width * 3;
	
				for(int y=0;y<b.Height;++y)
				{
					for(int x=0; x < nWidth; ++x )
					{
						p[0] = (byte)(255-p[0]);
						++p;
					}
					p += nOffset;
				}
			}

			b.UnlockBits(bmData);

			return true;
		}

		public static bool GrayScale(Bitmap b)
		{
			// GDI+ still lies to us - the return format is BGR, NOT RGB.
			BitmapData bmData = b.LockBits(new Rectangle(0, 0, b.Width, b.Height), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);

			int stride = bmData.Stride;
			System.IntPtr Scan0 = bmData.Scan0;

			unsafe
			{
				byte * p = (byte *)(void *)Scan0;

				int nOffset = stride - b.Width*3;

				byte red, green, blue;
	
				for(int y=0;y<b.Height;++y)
				{
					for(int x=0; x < b.Width; ++x )
					{
						blue = p[0];
						green = p[1];
						red = p[2];

						p[0] = p[1] = p[2] = (byte)(.299 * red + .587 * green + .114 * blue);

						p += 3;
					}
					p += nOffset;
				}
			}

			b.UnlockBits(bmData);

			return true;
		}

		public static bool Brightness(Bitmap b, int nBrightness)
		{
			if (nBrightness < -255 || nBrightness > 255)
				return false;

			// GDI+ still lies to us - the return format is BGR, NOT RGB.
			BitmapData bmData = b.LockBits(new Rectangle(0, 0, b.Width, b.Height), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);

			int stride = bmData.Stride;
			System.IntPtr Scan0 = bmData.Scan0;

			int nVal = 0;

			unsafe
			{
				byte * p = (byte *)(void *)Scan0;

				int nOffset = stride - b.Width*3;
				int nWidth = b.Width * 3;

				for(int y=0;y<b.Height;++y)
				{
					for(int x=0; x < nWidth; ++x )
					{
						nVal = (int) (p[0] + nBrightness);
		
						if (nVal < 0) nVal = 0;
						if (nVal > 255) nVal = 255;

						p[0] = (byte)nVal;

						++p;
					}
					p += nOffset;
				}
			}

			b.UnlockBits(bmData);

			return true;
		}

		public static bool Contrast(Bitmap b, sbyte nContrast)
		{
			if (nContrast < -100) return false;
			if (nContrast >  100) return false;

			double pixel = 0, contrast = (100.0+nContrast)/100.0;

			contrast *= contrast;

			int red, green, blue;
			
			// GDI+ still lies to us - the return format is BGR, NOT RGB.
			BitmapData bmData = b.LockBits(new Rectangle(0, 0, b.Width, b.Height), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);

			int stride = bmData.Stride;
			System.IntPtr Scan0 = bmData.Scan0;

			unsafe
			{
				byte * p = (byte *)(void *)Scan0;

				int nOffset = stride - b.Width*3;

				for(int y=0;y<b.Height;++y)
				{
					for(int x=0; x < b.Width; ++x )
					{
						blue = p[0];
						green = p[1];
						red = p[2];
				
						pixel = red/255.0;
						pixel -= 0.5;
						pixel *= contrast;
						pixel += 0.5;
						pixel *= 255;
						if (pixel < 0) pixel = 0;
						if (pixel > 255) pixel = 255;
						p[2] = (byte) pixel;

						pixel = green/255.0;
						pixel -= 0.5;
						pixel *= contrast;
						pixel += 0.5;
						pixel *= 255;
						if (pixel < 0) pixel = 0;
						if (pixel > 255) pixel = 255;
						p[1] = (byte) pixel;

						pixel = blue/255.0;
						pixel -= 0.5;
						pixel *= contrast;
						pixel += 0.5;
						pixel *= 255;
						if (pixel < 0) pixel = 0;
						if (pixel > 255) pixel = 255;
						p[0] = (byte) pixel;					

						p += 3;
					}
					p += nOffset;
				}
			}

			b.UnlockBits(bmData);

			return true;
		}
	
		public static bool Gamma(Bitmap b, double red, double green, double blue)
		{
			if (red < .2 || red > 5) return false;
			if (green < .2 || green > 5) return false;
			if (blue < .2 || blue > 5) return false;

			byte [] redGamma = new byte [256];
			byte [] greenGamma = new byte [256];
			byte [] blueGamma = new byte [256];

			for (int i = 0; i< 256; ++i)
			{
				redGamma[i] = (byte)Math.Min(255, (int)(( 255.0 * Math.Pow(i/255.0, 1.0/red)) + 0.5));
				greenGamma[i] = (byte)Math.Min(255, (int)(( 255.0 * Math.Pow(i/255.0, 1.0/green)) + 0.5));
				blueGamma[i] = (byte)Math.Min(255, (int)(( 255.0 * Math.Pow(i/255.0, 1.0/blue)) + 0.5));
			}

			// GDI+ still lies to us - the return format is BGR, NOT RGB.
			BitmapData bmData = b.LockBits(new Rectangle(0, 0, b.Width, b.Height), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);

			int stride = bmData.Stride;
			System.IntPtr Scan0 = bmData.Scan0;

			unsafe
			{
				byte * p = (byte *)(void *)Scan0;

				int nOffset = stride - b.Width*3;

				for(int y=0;y<b.Height;++y)
				{
					for(int x=0; x < b.Width; ++x )
					{
						p[2] = redGamma[ p[2] ];
						p[1] = greenGamma[ p[1] ];
						p[0] = blueGamma[ p[0] ];

						p += 3;
					}
					p += nOffset;
				}
			}

			b.UnlockBits(bmData);

			return true;
		}

		public static bool Color(Bitmap b, int red, int green, int blue)
		{
			if (red < -255 || red > 255) return false;
			if (green < -255 || green > 255) return false;
			if (blue < -255 || blue > 255) return false;

			// GDI+ still lies to us - the return format is BGR, NOT RGB.
			BitmapData bmData = b.LockBits(new Rectangle(0, 0, b.Width, b.Height), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);

			int stride = bmData.Stride;
			System.IntPtr Scan0 = bmData.Scan0;

			unsafe
			{
				byte * p = (byte *)(void *)Scan0;

				int nOffset = stride - b.Width*3;
				int nPixel;

				for(int y=0;y<b.Height;++y)
				{
					for(int x=0; x < b.Width; ++x )
					{
						nPixel = p[2] + red;
						nPixel = Math.Max(nPixel, 0);
						p[2] = (byte)Math.Min(255, nPixel);

						nPixel = p[1] + green;
						nPixel = Math.Max(nPixel, 0);
						p[1] = (byte)Math.Min(255, nPixel);

						nPixel = p[0] + blue;
						nPixel = Math.Max(nPixel, 0);
						p[0] = (byte)Math.Min(255, nPixel);

						p += 3;
					}
					p += nOffset;
				}
			}

			b.UnlockBits(bmData);

			return true;
		}
        public static bool Conv3x3(Bitmap b, ConvMatrix m) 
        { 
            // Avoid divide by zero errors 
            if (0 == m.Factor) return false; 
            
            Bitmap bSrc = (Bitmap)b.Clone(); // GDI+ still lies to us - the return format is BGR, NOT RGB. 
            BitmapData bmData = b.LockBits(new Rectangle(0, 0, b.Width, b.Height), 
                                           ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb); 
            BitmapData bmSrc = bSrc.LockBits(new Rectangle(0, 0, bSrc.Width, bSrc.Height), 
                                            ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb); 
            int stride = bmData.Stride; 
            int stride2 = stride * 2; 
        
            System.IntPtr Scan0 = bmData.Scan0; 
            System.IntPtr SrcScan0 = bmSrc.Scan0; 
        
            unsafe { 
                byte * p = (byte *)(void *)Scan0; 
                byte * pSrc = (byte *)(void *)SrcScan0; 
                int nOffset = stride - b.Width*3; 
                int nWidth = b.Width - 2; 
                int nHeight = b.Height - 2; 
                
                int nPixel; 
                
                for(int y=0;y < nHeight;++y) 
                { 
                    for(int x=0; x < nWidth; ++x ) 
                    {
                        nPixel = ( ( ( (pSrc[2] * m.TopLeft) + (pSrc[5] * m.TopMid) + 
                            (pSrc[8] * m.TopRight) + (pSrc[2 + stride] * m.MidLeft) + 
                            (pSrc[5 + stride] * m.Pixel) + (pSrc[8 + stride] * m.MidRight) + 
                            (pSrc[2 + stride2] * m.BottomLeft) + 
                            (pSrc[5 + stride2] * m.BottomMid) + 
                            (pSrc[8 + stride2] * m.BottomRight)) 
                            / m.Factor) + m.Offset); 
                            
                        if (nPixel < 0) nPixel = 0; 
                        if (nPixel > 255) nPixel = 255; 
                        p[5 + stride]= (byte)nPixel; 
                        
                        nPixel = ( ( ( (pSrc[1] * m.TopLeft) + (pSrc[4] * m.TopMid) + 
                            (pSrc[7] * m.TopRight) + (pSrc[1 + stride] * m.MidLeft) + 
                            (pSrc[4 + stride] * m.Pixel) + (pSrc[7 + stride] * m.MidRight) + 
                            (pSrc[1 + stride2] * m.BottomLeft) + 
                            (pSrc[4 + stride2] * m.BottomMid) + 
                            (pSrc[7 + stride2] * m.BottomRight)) 
                            / m.Factor) + m.Offset); 
                            
                        if (nPixel < 0) nPixel = 0; 
                        if (nPixel > 255) nPixel = 255; 
                        p[4 + stride] = (byte)nPixel; 
                        
                        nPixel = ( ( ( (pSrc[0] * m.TopLeft) + (pSrc[3] * m.TopMid) + 
                                       (pSrc[6] * m.TopRight) + (pSrc[0 + stride] * m.MidLeft) + 
                                       (pSrc[3 + stride] * m.Pixel) + 
                                       (pSrc[6 + stride] * m.MidRight) + 
                                       (pSrc[0 + stride2] * m.BottomLeft) + 
                                       (pSrc[3 + stride2] * m.BottomMid) + 
                                       (pSrc[6 + stride2] * m.BottomRight)) 
                            / m.Factor) + m.Offset); 
                            
                        if (nPixel < 0) nPixel = 0; 
                        if (nPixel > 255) nPixel = 255; 
                        p[3 + stride] = (byte)nPixel; 
                        
                        p += 3; 
                        pSrc += 3; 
                    } 
                    
                    p += nOffset; 
                    pSrc += nOffset; 
                } 
            } 
            
            b.UnlockBits(bmData); 
            
            bSrc.UnlockBits(bmSrc); 
            
            return true; 
        }
        
        public static bool Smooth(Bitmap b)
        {
            return Smooth(b, 1);
        }
        
        public static bool Smooth(Bitmap b, int nWeight)
        {
            ConvMatrix m = new ConvMatrix();
            m.SetAll(1);
            m.Pixel = nWeight;
            m.Factor = nWeight + 8;

            return  Conv3x3(b, m);
        }
        
        public static bool Sharpen(Bitmap b)
        {
            return Sharpen(b, 11);
        }
        
        public static bool Sharpen(Bitmap b, int nWeight /* default to 11 */)
        {
            ConvMatrix m = new ConvMatrix();
			m.SetAll(0);
			m.Pixel = nWeight;
			m.TopMid = m.MidLeft = m.MidRight = m.BottomMid = -2;
			m.Factor = nWeight - 8;
			return  Conv3x3(b, m);            
        }
        
        public static bool EmbossLaplacian(Bitmap b)
		{
			ConvMatrix m = new ConvMatrix();
			m.SetAll(-1);
			m.TopMid = m.MidLeft = m.MidRight = m.BottomMid = 0;
			m.Pixel = 4;
			m.Offset = 127;

			return  Conv3x3(b, m);
		}	
		
		public static bool EdgeDetectQuick(Bitmap b)
		{
			ConvMatrix m = new ConvMatrix();
			m.TopLeft = m.TopMid = m.TopRight = -1;
			m.MidLeft = m.Pixel = m.MidRight = 0;
			m.BottomLeft = m.BottomMid = m.BottomRight = 1;
		
			m.Offset = 127;

			return  Conv3x3(b, m);
		}
		
		public static bool MeanRemoval(Bitmap b)
		{
		    return MeanRemoval(b, 9);
		}
		
		public static bool MeanRemoval(Bitmap b, int nWeight /* default to 9*/ )
		{
			ConvMatrix m = new ConvMatrix();
			m.SetAll(-1);
			m.Pixel = nWeight;
			m.Factor = nWeight - 8;

			return Conv3x3(b, m);
		}
		
		public static bool GaussianBlur(Bitmap b)
		{
		    return GaussianBlur(b, 4);
		}
		
		public static bool GaussianBlur(Bitmap b, int nWeight /* default to 4*/)
		{
			ConvMatrix m = new ConvMatrix();
			m.SetAll(1);
			m.Pixel = nWeight;
			m.TopMid = m.MidLeft = m.MidRight = m.BottomMid = 2;
			m.Factor = nWeight + 12;

			return  Conv3x3(b, m);
		}
    }
    
    public class ImageUtil
    {        
        static Random m_objRandom = new Random(unchecked((int)DateTime.Now.Ticks));                 
        
        public static bool GetXmlBool(
            XmlNode objXmlNode,
            string strKey,
            bool fDefault)
        {
            if (objXmlNode.SelectSingleNode(strKey) == null)
                return fDefault;
                
            if (objXmlNode.SelectSingleNode(strKey).InnerText == "1")
                return true;
            else
                return false;
        }
        
        public static int GetXmlInt(
            XmlNode objXmlNode,
            string strKey,
            int iDefault)
        {
            if (objXmlNode.SelectSingleNode(strKey) == null)
                return iDefault;
                
            return Convert.ToInt32(objXmlNode.SelectSingleNode(strKey).InnerText);
        }
        
        public static Color GetXmlColor(
            XmlNode objXmlNode,
            string strKey,
            string strDefault)
        {       
            string strColor = strDefault;     
            if (objXmlNode.SelectSingleNode(strKey) != null)
                strColor = objXmlNode.SelectSingleNode(strKey).InnerText;
                
            return ColorTranslator.FromHtml(strColor);
        }                                        
        
        // 
        // Method: SaveImage
        //
        // Saves an image based on the xml profile for the save
        //
        public static void SaveImage(
            string strSourceFile,   // File containing the source image
            string strDestFile,     // Complete path to the destination file
            XmlNode objXmlNode
        )
        {
            int iDesiredWidth = Convert.ToInt32(objXmlNode.SelectSingleNode("/Format/DesiredWidth").InnerText);
            int iDesiredHeight = Convert.ToInt32(objXmlNode.SelectSingleNode("/Format/DesiredHeight").InnerText);
            int iTopMargin = GetXmlInt(objXmlNode, "/Format/TopMargin", 0);
            int iLeftMargin = GetXmlInt(objXmlNode, "/Format/LeftMargin", 0);
            int iBottomMargin = GetXmlInt(objXmlNode, "/Format/BottomMargin", 0);
            int iRightMargin = GetXmlInt(objXmlNode, "/Format/RightMargin", 0);            
            Color clrMargin = GetXmlColor(objXmlNode, "/Format/ColorMargin", "white");
            int iHAlign = GetXmlInt(objXmlNode, "/Format/HAlign", 1);
            int iVAlign = GetXmlInt(objXmlNode, "/Format/VAlign", 1);            
            bool fGrow = GetXmlBool(objXmlNode, "/Format/Grow", true);
            bool fShrink = GetXmlBool(objXmlNode, "/Format/Shrink", true);
            int iRotate = GetXmlInt(objXmlNode, "/Format/Rotate", 0);
            bool fPad = GetXmlBool(objXmlNode, "/Format/Pad", false);                 
            Color clrPad = GetXmlColor(objXmlNode, "/Format/ColorPad", "black");                   
            int iQuality = GetXmlInt(objXmlNode, "/Format/Quality", 50);
            bool fSharpen = GetXmlBool(objXmlNode, "/Format/Sharpen", true);
        
            SaveImage(
                    strSourceFile, 
                    strDestFile, 
                    iDesiredWidth, 
                    iDesiredHeight, 
                    iTopMargin, 
                    iBottomMargin, 
                    iLeftMargin, 
                    iRightMargin, 
                    clrMargin, 
                    iVAlign,
                    iHAlign,
                    fGrow,
                    fShrink,
                    iRotate,
                    fPad,
                    clrPad,
                    fSharpen,
                    iQuality
                  );
        }
        
        //
        // Method: SaveImage
        //
        // Description:
        // Saves an image according to the passed in parameters
        //
        public static void SaveImage(
            string strSourceFile,   // File containing the source image
            string strDestFile,     // Complete path to the destination file
            int iDesiredWidth,      // Desired Width
            int iDesiredHeight,     // Desired Height
            int iMarginTop,
            int iMarginBottom,
            int iMarginLeft,
            int iMarginRight,
            Color clrMargin,
            int iVAlign,
            int iHAlign,
            bool fGrow,
            bool fShrink,
            int iRotate,
            bool fPad,
            Color clrPad,       // Padding color
            bool fSharpen,      // Sharpen on a grow
            int iQuality)       // Value is 0 to 100.  0 is highest compression, 100 best quality
        {            
            Image oOriginalImage;
            int iOriginalImageHeight, iOriginalImageWidth;
            int iTargetImageWidth, iTargetImageHeight;
            int iFinalImageWidth, iFinalImageHeight;
            int iFinalBitmapWidth, iFinalBitmapHeight;
            int iFinalMarginedWidth, iFinalMarginedHeight;
            int iOffsetLeft, iOffsetTop;
            Bitmap oBitmapSave;        
            
            // We only know how to save a .jpg
            if (strDestFile.Substring(strDestFile.Length - 4).ToLower() != ".jpg")
                throw new Exception("File must be a .jpg");
            
            // Load the image from the file        
            oOriginalImage = Image.FromFile(strSourceFile);        
            
            // Get the original size of the image
            iOriginalImageHeight = oOriginalImage.Height;
            iOriginalImageWidth = oOriginalImage.Width;
            
            // Rotate the image for optimal display
            if ((iRotate == 1 || iRotate == 2)&& 
                ((iOriginalImageWidth < iOriginalImageHeight && iDesiredWidth > iDesiredHeight) || (iOriginalImageWidth > iOriginalImageHeight && iDesiredWidth < iDesiredHeight))
               )
            {        
                if (iRotate == 1)
                    oOriginalImage.RotateFlip(RotateFlipType.Rotate90FlipNone);
                else
                    oOriginalImage.RotateFlip(RotateFlipType.Rotate270FlipNone);
                iOriginalImageWidth = oOriginalImage.Width;
                iOriginalImageHeight = oOriginalImage.Height;        
            }
            
            // Get the target size once framed
            iTargetImageWidth = iDesiredWidth - (iMarginLeft + iMarginRight);
            iTargetImageHeight = iDesiredHeight - (iMarginTop + iMarginBottom);        
            
            // Check for growing and shrinking the image
            if (iOriginalImageWidth < iTargetImageWidth && iOriginalImageHeight < iTargetImageHeight)
            {
                if (!fGrow)
                {
                    iFinalImageWidth = iOriginalImageWidth;
                    iFinalImageHeight = iOriginalImageHeight;
                }
                else 
                {                    
                    if ((float)iOriginalImageWidth / (float)iTargetImageWidth > (float)iOriginalImageHeight / (float)iTargetImageHeight)
                    {
                        iFinalImageWidth = iTargetImageWidth;
                        iFinalImageHeight = iTargetImageWidth * iOriginalImageHeight / iOriginalImageWidth;
                    }
                    else
                    {
                        iFinalImageHeight = iTargetImageHeight;
                        iFinalImageWidth = iTargetImageHeight * iOriginalImageWidth / iOriginalImageHeight;
                    }            
                }
            }
            else if (fShrink)
            {                
                if ((float)iOriginalImageWidth / (float)iTargetImageWidth > (float)iOriginalImageHeight / (float)iTargetImageHeight)
                {
                    iFinalImageWidth = iTargetImageWidth;
                    iFinalImageHeight = iTargetImageWidth * iOriginalImageHeight / iOriginalImageWidth;
                }
                else
                {
                    iFinalImageHeight = iTargetImageHeight;
                    iFinalImageWidth = iTargetImageHeight * iOriginalImageWidth / iOriginalImageHeight;
                }
            }
            else
            {
                iFinalImageWidth = iOriginalImageWidth;
                iFinalImageHeight = iOriginalImageHeight;
            }     
            
            // Get the final size of the image with margins
            iFinalMarginedWidth = iFinalImageWidth + iMarginLeft + iMarginRight;
            iFinalMarginedHeight = iFinalImageHeight + iMarginTop + iMarginBottom;                  
            
            // Now get the final bitmap size            
            iFinalBitmapWidth = iFinalMarginedWidth;
            iFinalBitmapHeight = iFinalMarginedHeight;
            if (fPad)
            {
                if (iFinalBitmapWidth < iDesiredWidth)
                    iFinalBitmapWidth = iDesiredWidth;
                    
                if (iFinalBitmapHeight < iDesiredHeight)
                    iFinalBitmapHeight = iDesiredHeight;                                        
            }                                
            
            // Create the final bitmap
            oBitmapSave = new Bitmap(iFinalBitmapWidth, iFinalBitmapHeight);
            
            // Create rectangle the entire width
            Rectangle rect = new Rectangle(0, 0, iFinalBitmapWidth, iFinalBitmapHeight);
            
            // Create the graphics context
            Graphics g = Graphics.FromImage(oBitmapSave);
            
            iOffsetLeft = 0;
            iOffsetTop = 0;
            
            // If we pad, fill with the pad color first
            if (fPad && (iFinalBitmapWidth > iFinalMarginedWidth || iFinalBitmapHeight > iFinalMarginedHeight))
            {
                // Fill the rectangle with the pad color
                SolidBrush brushPad = new SolidBrush(clrPad);
                g.FillRectangle(brushPad, rect);                                                    
                
                if (iHAlign == 1)
                    iOffsetLeft += (rect.Width - iFinalMarginedWidth) / 2;
                else if (iHAlign == 2)
                    iOffsetLeft += (rect.Width - iFinalMarginedWidth);
                else if (iHAlign == 3)
                    iOffsetLeft += m_objRandom.Next(rect.Width - iFinalMarginedWidth);
            
                if (iVAlign == 1)
                    iOffsetTop += (rect.Height - iFinalMarginedHeight) / 2;
                else if (iVAlign == 2)
                    iOffsetTop += (rect.Height - iFinalMarginedHeight);
                else if (iVAlign == 3)
                    iOffsetTop += m_objRandom.Next(rect.Height - iFinalMarginedHeight);
            }
            
            // If we have a margin, render the margin margin color
            if (iMarginTop != 0 || iMarginLeft != 0 || iMarginRight != 0 || iMarginBottom != 0)
            {
                rect.X += iOffsetLeft;
                rect.Y += iOffsetTop;
                rect.Width = iFinalMarginedWidth;
                rect.Height = iFinalMarginedHeight;
                // Create solid brush for the margin color
                SolidBrush brushMargin = new SolidBrush(clrMargin);                
                g.FillRectangle(brushMargin, rect);
                
                // Adjust the offset and rectangle by the margin
                iOffsetLeft += iMarginLeft;
                iOffsetTop += iMarginTop;
                rect.X += iMarginLeft;
                rect.Y += iMarginTop;
                rect.Width -= (iMarginLeft + iMarginRight);
                rect.Height -= (iMarginTop + iMarginBottom);                
            }                                
            
            // Draw the image into the bitmap    
            g.DrawImage(oOriginalImage, iOffsetLeft, iOffsetTop, iFinalImageWidth, iFinalImageHeight);
            
            // Sharpen if we grew
            if (fSharpen && fGrow && (iOriginalImageHeight < iFinalImageWidth || iOriginalImageWidth < iFinalImageWidth))
            {
                Conv.Sharpen(oBitmapSave);
            }
            
            // Save the final file    
            System.Drawing.Imaging.Encoder myEncoder = System.Drawing.Imaging.Encoder.Quality;
            EncoderParameters myEncoderParameters = new EncoderParameters(1);
            EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, iQuality);
            myEncoderParameters.Param[0] = myEncoderParameter;					                        
    	    								
    	    oBitmapSave.Save(strDestFile, GetEncoderInfo("image/jpeg"), myEncoderParameters);	    
    	    
    	    oOriginalImage.Dispose();
        }	
    	
    	
    	public static ImageCodecInfo GetEncoderInfo(String mimeType)
        {
            int j;
            ImageCodecInfo[] encoders;
            encoders = ImageCodecInfo.GetImageEncoders();
            for(j = 0; j < encoders.Length; ++j)
            {
                if(encoders[j].MimeType == mimeType)
                    return encoders[j];
            }
            return null;
        }
    }                 
}
