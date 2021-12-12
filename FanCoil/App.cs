using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Media.Imaging;
using System.IO;
using System.Windows.Media;
using System.Reflection;




namespace FanCoil
{

    public class App : IExternalApplication
    {

        public Result OnStartup(UIControlledApplication application)
        {

            string tabName = "Fan Coil";
            string panelName = "Create Fan Coils";
            application.CreateRibbonTab(tabName);
            // Initialize whole plugin's user interface.

            RibbonPanel panel = application.CreateRibbonPanel(tabName, panelName);



            Image img = FanCoil.Properties.Resources.fancoilimg;

            ImageSource imgSrc = GetSoruceImage(img);



            PushButtonData btnData = new PushButtonData(
                "Button",
                "Run",
                Assembly.GetExecutingAssembly().Location,
                "FanCoil.Command"
                );

            PushButton button = panel.AddItem(btnData) as PushButton;

            button.Image = imgSrc;
            button.LargeImage = imgSrc;
            button.ToolTip = "Short Description";
            button.Enabled = true;


            return Result.Succeeded;
        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private BitmapSource GetSoruceImage(Image img)
        {
            BitmapImage bmp = new BitmapImage();
            using (MemoryStream ms = new MemoryStream())
            {
                img.Save(ms, ImageFormat.Png);
                ms.Position = 0;

                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.UriSource = null;
                bmp.StreamSource = ms;


                bmp.EndInit();
            }
            return bmp;
        }
    }


}


