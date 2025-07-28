using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EyesOfficeTesterLib;
using static System.Net.Mime.MediaTypeNames;


namespace EyesOfficeTesterCli
{
    /// <summary>
    /// An ASCII progress bar
    /// </summary>
    public class EyesOfficeTesterCliProgressBar : 
        IDisposable, IProgress<EyesOfficeProgressUpdate>
    {
        private const int blockCount = 10;
        private readonly TimeSpan animationInterval = TimeSpan.FromSeconds(1.0 / 8);
        private const string animation = @"|/-\";

        private readonly System.Threading.Timer timer;

        private double currentProgress = 0;
        private string currentText = string.Empty;
        private bool disposed = false;
        private int animationIndex = 0;
        private bool isWriteImage = false;
        private string message;

        public EyesOfficeTesterCliProgressBar(bool isWriteImage = false)
        {
            timer = new System.Threading.Timer(TimerHandler);

            // A progress bar is only for temporary display in a console window.
            // If the console output is redirected to a file, draw nothing.
            // Otherwise, we'll end up with a lot of garbage in the target file.
            if (!Console.IsOutputRedirected)
            {
                ResetTimer();
            }
            this.isWriteImage = isWriteImage;
        }

        public void Report(EyesOfficeProgressUpdate eyesOfficeProgressUpdate)
        {
            //Console.WriteLine(eyesOfficeProgressUpdate.progressMessage);
            if (eyesOfficeProgressUpdate.progressMessage != null)
            {
                message = eyesOfficeProgressUpdate.progressMessage;
            }
            
            // Make sure value is in [0..1] range
            eyesOfficeProgressUpdate.progressValue = 
                Math.Max((byte)0, 
                Math.Min((byte)1, 
                (byte)eyesOfficeProgressUpdate.progressValue));

            Interlocked.Exchange(
                ref currentProgress, 
                (double)eyesOfficeProgressUpdate.progressValue);

            if(isWriteImage && eyesOfficeProgressUpdate.bitmap != null)
            {
                string timestamp = DateTime.UtcNow.ToString("yyyy-MM-dd_HH:mm:ss.fff",
                                            CultureInfo.InvariantCulture);

                using (MemoryStream ms = new MemoryStream())
                {
                    eyesOfficeProgressUpdate.bitmap.Save(ms, ImageFormat.Png);
                    eyesOfficeProgressUpdate.pngBytes = ms.ToArray();
                }

                string imageDir = 
                    Path.GetDirectoryName(
                        System.Reflection.Assembly.GetExecutingAssembly().Location);

                string imagePath = 
                    imageDir + Path.DirectorySeparatorChar + timestamp + ".png";

                File.WriteAllBytes(imagePath, eyesOfficeProgressUpdate.pngBytes);
            }
        }

        private void TimerHandler(object state)
        {
            lock (timer)
            {
                if (disposed) return;

                int progressBlockCount = (int)(currentProgress * blockCount);
                int percent = (int)(currentProgress * 100);
                string text = string.Format("[{0}{1}] {2,3}% {3} - " + message,
                    new string('#', progressBlockCount), 
                    new string('-', blockCount - progressBlockCount),
                    percent,
                    animation[animationIndex++ % animation.Length]);
                UpdateText(text);

                ResetTimer();
            }
        }

        private void UpdateText(string text)
        {
            // Get length of common portion
            int commonPrefixLength = 0;
            int commonLength = Math.Min(currentText.Length, text.Length);

            while (commonPrefixLength < commonLength && 
                text[commonPrefixLength] == currentText[commonPrefixLength])
            {
                commonPrefixLength++;
            }

            // Backtrack to the first differing character
            StringBuilder outputBuilder = new StringBuilder();
            outputBuilder.Append('\b', currentText.Length - commonPrefixLength);

            // Output new suffix
            outputBuilder.Append(text.Substring(commonPrefixLength));

            // If the new text is shorter than the old one: delete overlapping characters
            int overlapCount = currentText.Length - text.Length;
            if (overlapCount > 0)
            {
                outputBuilder.Append(' ', overlapCount);
                outputBuilder.Append('\b', overlapCount);
            }

            Console.Write(outputBuilder);
            currentText = text;
        }

        private void ResetTimer()
        {
            timer.Change(animationInterval, TimeSpan.FromMilliseconds(-1));
        }

        public void Dispose()
        {
            lock (timer)
            {
                disposed = true;
                UpdateText(string.Empty);
            }
        }

    }
}
