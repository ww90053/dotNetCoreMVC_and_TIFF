using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace ClassLibrary1
{
    public class TIFFHelper
    {
        /// <summary>
        /// 分割Tiff檔的範例
        /// </summary>
        public void DividTiffSample(
            string tiffFilePath = "D:\\SVN\\TIFF檔處理\\範例檔案\\稀見清代四部補編002冊__OK.tiff",
            string destFilePath = "D:\\SVN\\TIFF檔處理\\範例檔案\\"
            )
        {
            //載入tiff檔案
            System.Drawing.Bitmap image = new Bitmap(tiffFilePath);
            //下面這段是要抓
            //抓取Frame Dimension的清單,對應到一組guid陣列
            //這邊抓到的陣列好像都只會有一個元素
            Guid[] guid = image.FrameDimensionsList;
            //抓取第一個guid,產生FrameDimension物件
            FrameDimension fd = new FrameDimension(guid[0]);
            //得出有幾張圖片
            int frames = image.GetFrameCount(fd);
            for (int i = 0; i < frames; i++)
            {
                //根據剛才的FD物件,和圖片的索引位置,切換目前的圖片
                image.SelectActiveFrame(fd, i);
                //儲存目前的圖片
                image.Save(destFilePath + DateTime.Now.ToString("yyyyMMddHHmmss") + @"_pic_" + i.ToString() + ".tif"
                    , System.Drawing.Imaging.ImageFormat.Tiff);
            }
        }
        public static bool TiffAddWatermark(
    string tiffFilePath = "D:\\SVN\\TIFF檔處理\\範例檔案\\稀見清代四部補編002冊__OK.tiff",
    string WaterMarkFilePath = "D:\\SVN\\TIFF檔處理\\範例檔案\\新影像.PNG",
    string destFilePath = "D:\\SVN\\TIFF檔處理\\範例檔案\\"
    )
        {
            //讀取原始的tiff

            //分割成多組的Bitmap
            //每個Bitmap追加浮水印
            //合併所有的Bitmap

            //載入tiff檔案
            System.Drawing.Bitmap image = new Bitmap(tiffFilePath);
            System.Drawing.Bitmap image_W = new Bitmap(WaterMarkFilePath);
            //下面這段是要抓
            //抓取Frame Dimension的清單,對應到一組guid陣列
            //這邊抓到的陣列好像都只會有一個元素
            Guid[] guid = image.FrameDimensionsList;
            //抓取第一個guid,產生FrameDimension物件
            FrameDimension fd = new FrameDimension(guid[0]);

            //得出有幾張圖片
            int frames = image.GetFrameCount(fd);
            for (int i = 0; i < frames; i++)
            {
                //根據剛才的FD物件,和圖片的索引位置,切換目前的圖片
                image.SelectActiveFrame(fd, i);
                Image img = AddImgToImg(image, image_W, 1000, 1000
                    , (float)0.3, ".tiff");
                //img.Save
                //儲存目前的圖片
                img.Save(destFilePath + DateTime.Now.ToString("yyyyMMddHHmmss") + @"_pic_" + i.ToString() + ".tif"
                    , System.Drawing.Imaging.ImageFormat.Tiff);
            }

            return true;
        }
        //合併TIFF圖檔
        private void MergeTiffImages(string[] sourceFilePath, string destFilePath)
        {
            Bitmap multiPages = null;
            EncoderParameters eParam = new EncoderParameters(1);
            Encoder myEncoder = Encoder.SaveFlag;
            eParam.Param[0] = new EncoderParameter(myEncoder, (long)EncoderValue.MultiFrame);
            ImageCodecInfo codecInfo = null;

            foreach (ImageCodecInfo imgCodecInfo in ImageCodecInfo.GetImageEncoders())
            {
                if (imgCodecInfo.MimeType == "image/tiff")
                {
                    codecInfo = imgCodecInfo;
                    break;
                }
            }
            //逐筆處理來源的圖片檔案列
            for (int i = 0; i < sourceFilePath.Length; i++)
            {
                //如果是第一筆
                if (i == 0)
                {
                    //載入來源圖片 放到Bitmap中
                    multiPages = (Bitmap)Image.FromFile(sourceFilePath[i]);
                    //儲存成tiff檔 後面兩個參數不知道為何要寫到這麼複雜?
                    multiPages.Save(destFilePath, codecInfo, eParam);
                }
                else
                {
                    //不是第一筆的話
                    eParam.Param[0] = new EncoderParameter(myEncoder, (long)EncoderValue.FrameDimensionPage);
                    //持續加入Bitmap中,並儲存
                    multiPages.SaveAdd((Bitmap)Image.FromFile(sourceFilePath[i]), eParam);
                }
                //如果是最後一筆的話
                if (i == sourceFilePath.Length - 1)
                {
                    eParam.Param[0] = new EncoderParameter(myEncoder, (long)EncoderValue.Flush);
                    multiPages.SaveAdd(eParam);
                }
            }
        }


        /// <summary>
        /// 添加圖片水印
        /// </summary>
        /// <param name="image_from"></param>
        /// <param name="text"></param>
        /// <param name="rectX">水印開始X坐標（自動扣除圖片寬度）</param>
        /// <param name="rectY">水印開始Y坐標（自動扣除圖片高度</param>
        /// <param name="opacity">透明度 0-1</param>
        /// <param name="externName">文件后綴名</param>
        /// <returns></returns>
        public static Image AddImgToImg(Image image_from, Image watermark, float rectX, float rectY, float opacity, string externName)
        {
            //將來源圖片,複製到Bitmap上
            Bitmap bitmap = new Bitmap(image_from, image_from.Width, image_from.Height);
            //根據來源圖片,建立畫布
            Graphics g = Graphics.FromImage(bitmap);

            //下面定義一個矩形區域      
            float rectWidth = watermark.Width + 10;
            float rectHeight = watermark.Height + 10;
            //矩形區域 起始的X Y 座標 大概是原始圖片的10%
            float x = image_from.Width / 10;
            float y = image_from.Height / 10;
            //矩形區域 的寬 大概是原始圖片的90%
            rectWidth= (image_from.Width*9) / 10;
            //高則是依照浮水印的原圖比例
            rectHeight = (rectWidth * watermark.Height) / watermark.Width;

            //聲明矩形域
            RectangleF textArea = new RectangleF(x, y, rectWidth, rectHeight);
            //根據來源的浮水印原始照片,改變圖片的透明度,再產生Bitmap
            Bitmap w_bitmap = ChangeOpacity(watermark, opacity);
            //將剛才透明處理後的浮水印,加到畫布上
            g.DrawImage(w_bitmap, textArea);
            MemoryStream ms = new MemoryStream();

            //保存圖片 到 MemoryStream中
            switch (externName)
            {
                case ".jpg":
                    bitmap.Save(ms, ImageFormat.Jpeg);
                    break;
                case ".gif":
                    bitmap.Save(ms, ImageFormat.Gif);
                    break;
                case ".png":
                    bitmap.Save(ms, ImageFormat.Png);
                    break;
                case ".tiff":
                    bitmap.Save(ms, ImageFormat.Tiff);
                    break;
                default:
                    bitmap.Save(ms, ImageFormat.Jpeg);
                    break;
            }
            //再根據MemoryStream 轉換成新的Image
            Image h_hovercImg = Image.FromStream(ms);

            g.Dispose();
            bitmap.Dispose();

            return h_hovercImg;

        }

        /// <summary>
        /// 改變圖片的透明度
        /// </summary>
        /// <param name="img">圖片</param>
        /// <param name="opacityvalue">透明度</param>
        /// <returns></returns>
        public static Bitmap ChangeOpacity(Image img, float opacityvalue)
        {

            float[][] nArray ={ new float[] {1, 0, 0, 0, 0},

                                new float[] {0, 1, 0, 0, 0},

                                new float[] {0, 0, 1, 0, 0},

                                new float[] {0, 0, 0, opacityvalue, 0},

                                new float[] {0, 0, 0, 0, 1}};

            ColorMatrix matrix = new ColorMatrix(nArray);

            ImageAttributes attributes = new ImageAttributes();

            attributes.SetColorMatrix(matrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);

            Image srcImage = img;

            Bitmap resultImage = new Bitmap(srcImage.Width, srcImage.Height);

            Graphics g = Graphics.FromImage(resultImage);

            g.DrawImage(srcImage, new Rectangle(0, 0, srcImage.Width, srcImage.Height), 0, 0, srcImage.Width, srcImage.Height, GraphicsUnit.Pixel, attributes);

            return resultImage;
        }
    }
}
