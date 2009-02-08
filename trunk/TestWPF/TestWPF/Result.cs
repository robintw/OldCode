using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.IO;
using System.Windows.Media.Imaging;
using System.Diagnostics;

namespace TestWPF
{
    class Result
    {
        private System.Drawing.Icon DisplayIcon;

        public Result(string _name, string _path)
        {
            Name = _name;
            Path = _path;
        }

        public string Name
        {
            get;
            set;
        }

        public override string ToString()
        {
            return Name;
        }

        public string Path
        {
            get;
            set;
        }

        public System.Drawing.Icon Icon
        {
            set
            {
                DisplayIcon = value;
            }
        }

        public System.Windows.Media.ImageSource Image
        {
            get
            {
                System.Drawing.Icon ico = System.Drawing.Icon.ExtractAssociatedIcon(Path);
                Bitmap bmp = ico.ToBitmap();

                MemoryStream strm = new MemoryStream();
                bmp.Save(strm, System.Drawing.Imaging.ImageFormat.Png);
                strm.Seek(0, SeekOrigin.Begin);

                PngBitmapDecoder pbd = new PngBitmapDecoder(strm, BitmapCreateOptions.None, BitmapCacheOption.Default);

                return pbd.Frames[0];
            }
        }

        public void Run()
        {
            Process proc = new Process();
            proc.StartInfo.FileName = Path;
            proc.Start();
        }

    }
}
