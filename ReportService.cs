using Docentric.Documents.Reporting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace JoinDocFileIssue
{
    public static class ReportService
    {
        public static void GenerateReport(string inputTemplatePath, string outputPath, object data)
        {
            CreateDirectory(outputPath);

            using (Stream templateStream = File.OpenRead(inputTemplatePath))
            using (Stream reportDocumentStream = File.Create(outputPath))
            {
                var dg = new DocumentGenerator(null);

                dg.WriteErrorsAndWarningsToDocument = false;
                dg.SetDefaultDataSourceValue(data);
                
                //generate Docentric document
                var result = dg.GenerateDocument(templateStream, reportDocumentStream);

                if (result.HasErrors)
                {
                    IEnumerable<string> errorMessages = result.Errors.Select(x => $"Severity: {x.Severity.ToString()}; Message: {x.Message}");
                    throw new Exception(string.Join(Environment.NewLine, errorMessages));
                }
            }
        }

        #region -- Private Methods --
        private static void CreateDirectory(string path)
        {
            string dirPath = Path.GetDirectoryName(path);
            if (!Directory.Exists(dirPath))
                Directory.CreateDirectory(dirPath);
        }
        #endregion -- Private Methods --
    }
}
