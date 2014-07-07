using System;
using System.IO;
using System.Collections.Generic;
using System.Text;


namespace convertTool
{
    class Program
    {
        static void Main(string[] args)
        {
            
            if (args.Length <2 )
            {
                Console.Write("Usage: convertTooll inputFileName outputFileName [fileType]");
                return;
            }
            string inputFileName = @args[0];
            string outputFileName = @args[1];
            string fileType = "doc";
            if (args.Length == 3)
            {
                fileType = args[2];
            }

            if (File.Exists(inputFileName) == false){
                Console.WriteLine("Error: Input file is not exists.");
                return;
            }

            Console.WriteLine("inputFileName:" + inputFileName);
            Console.WriteLine("outputFileName:" + outputFileName);
            Console.WriteLine("fileType:" + fileType);

            ConvertClass c = new ConvertClass();

            bool ret = false;

            
                if (fileType == "doc")
                {
                    ret = c.Convert(inputFileName, outputFileName, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                }
                else if (fileType == "xls")
                {
                    ret = c.Convert(inputFileName, outputFileName, Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF);
                }
                else if (fileType == "ppt")
                {
                    ret = c.Convert(inputFileName, outputFileName, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPDF);
                }
                else if (fileType == "pdf")
                {
                    //outputFileName = inputFileName;
                    ret = c.Convert(inputFileName, outputFileName);
                }
            

            if (ret == true)
            {
                //PdfToSwf p = new PdfToSwf();
                //ret = p.Convert(outputFileName);

                Console.WriteLine("Convert to pdf complated.");

                //Environment.Exit(1);
            }

        }
    }
}
