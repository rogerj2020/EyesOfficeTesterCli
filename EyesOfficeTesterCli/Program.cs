// See https://aka.ms/new-console-template for more information
using EyesOfficeTesterLib;
using System.Globalization;


namespace EyesOfficeTesterCli
{
    internal class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var cliParser = new CommandLineParser(args);

            String? workingDir = cliParser.GetStringArgument("directory", 'd');
            workingDir = workingDir ?? Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            bool isSaveImages = cliParser.GetSwitchArgument("saveImages", 's');

            bool hasProgressBar = cliParser.GetSwitchArgument("progressBar", 'p');
            bool notifyOnCompletion = cliParser.GetSwitchArgument("notify", 'n');

            String? apiKey = cliParser.GetStringArgument("apiKey", 'k');
            String? serverUrl = cliParser.GetStringArgument("serverUrl", 'u');


            IProgress<EyesOfficeProgressUpdate> progress;
            EyesOfficeTester eyesOfficeTester;
            if (hasProgressBar)
            {
                progress = new EyesOfficeTesterCliProgressBar(isSaveImages);
            }
            else
            {
                progress = new Progress<EyesOfficeProgressUpdate>(update =>
                {
                    string timestamp = DateTime.UtcNow.ToString("yyyy-MM-dd_HH.mm.ss.fff",
                                                    CultureInfo.InvariantCulture);

                    Console.WriteLine("[" + timestamp + "] " + update.progressMessage + " (" + update.progressValue + "%) ***");
                    if (isSaveImages && update.pngBytes != null)
                    {
                        string imageDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                        string imagePath = imageDir + Path.DirectorySeparatorChar + timestamp + ".png";

                        File.WriteAllBytes(imagePath, update.pngBytes);
                        Console.WriteLine("Saved Image file: " + imagePath);
                    }
                });
            }

            if (String.IsNullOrEmpty(apiKey))
            {
                eyesOfficeTester = new EyesOfficeTester(progress);
            }
            else
            {
                eyesOfficeTester = new EyesOfficeTester(apiKey, progress);
            }

            if (!String.IsNullOrEmpty(serverUrl))
            {
                eyesOfficeTester.ServerUrl(serverUrl);
            }

            eyesOfficeTester.ReportImages(isSaveImages);
            eyesOfficeTester.NotifyOnCompletion(notifyOnCompletion);
            eyesOfficeTester.CheckOfficeFiles(workingDir, hasProgressBar);
        }
    }
}







