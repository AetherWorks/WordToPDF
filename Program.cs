using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace WordToPDF
{
    /// <summary>
    /// Program which saves a Micrsoft Word document to PDF.
    /// Based on http://msdn.microsoft.com/en-us/library/bb412305.aspx?cs-save-lang=1&cs-lang=csharp#code-snippet-4.
    /// </summary>
    class Program
    {
        /// <summary>
        /// Take a MS Word file, and output a corresponding PDF>
        /// </summary>
        /// <param name="args">[0]: Path to the file to be converted; [1]: Path to the output file</param>
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.Out.WriteLine("Incorrect number of parameters. Expected: <source> <result>, was " + args.Length);
                return;
            }
            
            // Initialize Word interop references.
            Application wordApplication = new Application();
            Document wordDocument = null;

            // Get full path of input and output files.
            object paramSourceDocPath = Path.GetFullPath(args[0]);
            string paramExportFilePath = Path.GetFullPath(args[1]);
            
            ///////////////////////////////////////////////
            /// STATIC CONFIGURATION PARAMETERS
            //////////////////////////////////////////////

            WdExportFormat paramExportFormat = WdExportFormat.wdExportFormatPDF;
            bool paramOpenAfterExport = false;
            WdExportOptimizeFor paramExportOptimizeFor =
                WdExportOptimizeFor.wdExportOptimizeForPrint;
            WdExportRange paramExportRange = WdExportRange.wdExportAllDocument;
            
            int paramStartPage = 0;
            int paramEndPage = 0;
            
            WdExportItem paramExportItem = WdExportItem.wdExportDocumentContent;

            bool paramIncludeDocProps = true;
            bool paramKeepIRM = true;
            WdExportCreateBookmarks paramCreateBookmarks =
                WdExportCreateBookmarks.wdExportCreateWordBookmarks;
            
            bool paramDocStructureTags = true;
            bool paramBitmapMissingFonts = true;
            bool paramUseISO19005_1 = false;

            object paramMissing = Type.Missing;

            ///////////////////////////////////////////////
            /// CONVERSION
            //////////////////////////////////////////////

            try
            {
                // Open the source document.
                wordDocument = wordApplication.Documents.Open(
                    ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing);

                // Export it in the specified format.
                if (wordDocument != null)
                    wordDocument.ExportAsFixedFormat(paramExportFilePath,
                        paramExportFormat, paramOpenAfterExport,
                        paramExportOptimizeFor, paramExportRange, paramStartPage,
                        paramEndPage, paramExportItem, paramIncludeDocProps,
                        paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                        paramBitmapMissingFonts, paramUseISO19005_1,
                        ref paramMissing);
            }
            catch (Exception ex)
            {
                Console.Out.WriteLine(ex);
            }
            finally
            {
                // Close and release the Document object.
                if (wordDocument != null)
                {
                    wordDocument.Close(ref paramMissing, ref paramMissing,
                        ref paramMissing);
                    wordDocument = null;
                }

                // Quit Word and release the ApplicationClass object.
                if (wordApplication != null)
                {
                    wordApplication.Quit(ref paramMissing, ref paramMissing,
                        ref paramMissing);
                    wordApplication = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
