using System;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.IO;

/// <summary>
/// Class1 的摘要说明
/// </summary>
public class ConvertClass
{
    public ConvertClass()
	{
		//
		// TODO: 在此处添加构造函数逻辑
		//
	}

    //将word文档转换成PDF格式
    public bool Convert(string sourcePath, string targetPath, Word.WdExportFormat exportFormat)
    {
        bool result;
        object paramMissing = Type.Missing;
        Word.ApplicationClass wordApplication = new Word.ApplicationClass();
        Word.Document wordDocument = null;
        try
        {
            object paramSourceDocPath = sourcePath;
            string paramExportFilePath = targetPath;

            Word.WdExportFormat paramExportFormat = exportFormat;
            bool paramOpenAfterExport = false;
            Word.WdExportOptimizeFor paramExportOptimizeFor =
                    Word.WdExportOptimizeFor.wdExportOptimizeForPrint;
            Word.WdExportRange paramExportRange = Word.WdExportRange.wdExportAllDocument;
            int paramStartPage = 0;
            int paramEndPage = 0;
            Word.WdExportItem paramExportItem = Word.WdExportItem.wdExportDocumentContent;
            bool paramIncludeDocProps = true;
            bool paramKeepIRM = true;
            Word.WdExportCreateBookmarks paramCreateBookmarks =
                    Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks;
            bool paramDocStructureTags = true;
            bool paramBitmapMissingFonts = true;
            bool paramUseISO19005_1 = false;

            wordDocument = wordApplication.Documents.Open(
                    ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing, ref paramMissing, ref paramMissing,
                    ref paramMissing);

            if (wordDocument != null)
                wordDocument.ExportAsFixedFormat(paramExportFilePath,
                        paramExportFormat, paramOpenAfterExport,
                        paramExportOptimizeFor, paramExportRange, paramStartPage,
                        paramEndPage, paramExportItem, paramIncludeDocProps,
                        paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                        paramBitmapMissingFonts, paramUseISO19005_1,
                        ref paramMissing);
            result = true;
        }
        finally
        {
            if (wordDocument != null)
            {
               wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                wordDocument = null;
            }
            if (wordApplication != null)
            {
                wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                wordApplication = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        return result;
    }

    //将excel文档转换成PDF格式
    public bool Convert(string sourcePath, string targetPath, Excel.XlFixedFormatType targetType)
    {
        bool result;
        object missing = Type.Missing;
        Excel.ApplicationClass application = null;
        Excel.Workbook workBook = null;


        try
        {
            application = new Excel.ApplicationClass();
            object target = targetPath;
            object type = targetType;
            workBook = application.Workbooks.Open(sourcePath, missing, missing, missing, missing, missing,
                true, missing, missing, missing, missing, missing, missing, missing, missing);

            workBook.ExportAsFixedFormat(targetType, target, Excel.XlFixedFormatQuality.xlQualityStandard, true, false, missing, missing, missing, missing);
            result = true;
        }
        catch
        {
            result = false;
        }
        finally
        {
            if (workBook != null)
            {
                workBook.Close(true, missing, missing);
                workBook = null;
            }
            if (application != null)
            {
                application.Quit();
                application = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        return result;
    }

    //将ppt文档转换成PDF格式
    public bool Convert(string sourcePath, string targetPath, PowerPoint.PpSaveAsFileType targetFileType)
    {
        bool result;
        object missing = Type.Missing;
        PowerPoint.ApplicationClass application = null;
        PowerPoint.Presentation persentation = null;
        try
        {
            application = new PowerPoint.ApplicationClass();
            persentation = application.Presentations.Open(sourcePath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
            persentation.SaveAs(targetPath, targetFileType, Microsoft.Office.Core.MsoTriState.msoTrue);

            result = true;
        }
        catch
        {
            result = false;
        }
        finally
        {
            if (persentation != null)
            {
                persentation.Close();
                persentation = null;
            }
            if (application != null)
            {
                application.Quit();
                application = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        return result;
    }

    //remove password from pdf
    public bool Convert(string inputFileName, string outputFileName)
    {

        PdfDocument pdf = new PdfDocument();
        try
        {
            pdf = PdfReader.Open(inputFileName, PdfDocumentOpenMode.Import);
        }
        catch (PdfSharp.Pdf.IO.PdfReaderException)
        {
            try
            {
                string newName = this.WriteCompatiblePdf(inputFileName);
                pdf = PdfReader.Open(newName, PdfDocumentOpenMode.Import);
            }
            catch 
            {
                return false;
            }
        }
        PdfDocument newPdf = new PdfDocument();
        foreach (PdfPage page in pdf.Pages)
        {
            newPdf.Pages.Add(page);
        }
        newPdf.Save(outputFileName);

        return true;
    }

    private string WriteCompatiblePdf(string sFilename)
    {
        string sNewPdf = System.IO.Path.GetTempFileName();

        iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(sFilename);

        int n = reader.NumberOfPages;

        iTextSharp.text.Document document = new iTextSharp.text.Document(reader.GetPageSizeWithRotation(1));

        iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(sNewPdf, FileMode.Create));
        writer.SetPdfVersion(iTextSharp.text.pdf.PdfWriter.PDF_VERSION_1_4);
        document.Open();
        iTextSharp.text.pdf.PdfContentByte cb = writer.DirectContent;
        iTextSharp.text.pdf.PdfImportedPage page;

        int rotation;

        int i = 0;

        while (i < n)
        {
            i++;
            document.SetPageSize(reader.GetPageSizeWithRotation(i));
            document.NewPage();
            page = writer.GetImportedPage(reader, i);
            rotation = reader.GetPageRotation(i);
            if (rotation == 90 || rotation == 270)
            {
                cb.AddTemplate(page, 0, -1f, 1f, 0, 0, reader.GetPageSizeWithRotation(i).Height);
            }
            else
            {
                cb.AddTemplate(page, 1f, 0, 0, 1f, 0, 0);
            }
        }
        // step 5: we close the document
        document.Close();
        return sNewPdf;
    }
}
