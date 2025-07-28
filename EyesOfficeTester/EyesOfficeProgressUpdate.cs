using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EyesOfficeTesterLib
{
    public class EyesOfficeProgressUpdate
    {
        public string? progressMessage;
        public int? progressValue;
        public Bitmap? bitmap;
        public byte[] pngBytes = null;
        public bool hasError = false;
    }
}
