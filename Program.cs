using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JoinDocFileIssue
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start processing...");

            string templatePath = Path.Combine(Environment.CurrentDirectory, "Templates", "main.docx");
            string outputDir = ConfigurationManager.AppSettings.Get("OutputDirectoryName");
            string outputFileName = ConfigurationManager.AppSettings.Get("OutputFileName");
            string outputPath = Path.Combine(Environment.CurrentDirectory, outputDir, outputFileName);

            try
            {
                //do file generation
                ReportService.GenerateReport(templatePath, outputPath, new { });

                Console.WriteLine("File generated successfully!");
                Console.WriteLine($"Path: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR: {ex.Message}");
            }

            Console.WriteLine();
            Console.Write("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
