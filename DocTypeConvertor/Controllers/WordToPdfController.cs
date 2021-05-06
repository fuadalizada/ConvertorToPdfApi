using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Web.Http;
using DocTypeConvertor.Models;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using Document = Microsoft.Office.Interop.Word.Document;

namespace DocTypeConvertor.Controllers
{
    public class WordToPdfController : ApiController
    {
        #region Word to Pdf

        [HttpGet]
        [Route("api/WordToPdf/ConvertWordToPdf")]
        public ResponseModel ConvertWordToPdf(string fileLocation, string outLocation)
        {
            try
            {
                string newPath = "";
                if (!File.Exists(outLocation))
                {
                    Application app = new Application();
                    app.Visible = false;
                    app.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                    Documents documents = app.Documents;
                    Document doc = documents.Open(fileLocation);
                    newPath = Path.GetDirectoryName(fileLocation);
                    if (newPath != null)
                    {
                        newPath = newPath.Replace(newPath, outLocation);
                        if (!File.Exists(newPath))
                        {
                            doc.SaveAs2(newPath, WdSaveFormat.wdFormatPDF);
                        }
                        else
                        {
                            Marshal.ReleaseComObject(documents);
                            doc.Close();
                            Marshal.ReleaseComObject(doc);
                            app.Quit();
                            Marshal.ReleaseComObject(app);
                            return (new ResponseModel { IsSucceed = false, ErrorMessage = newPath + " file-ı artıq mövcuddur." });
                        }
                    }
                    else
                    {
                        Marshal.ReleaseComObject(documents);
                        doc.Close();
                        Marshal.ReleaseComObject(doc);
                        app.Quit();
                        Marshal.ReleaseComObject(app);
                        return (new ResponseModel { IsSucceed = false, ErrorMessage = newPath + " null dəyər alıb." });
                    }

                    Marshal.ReleaseComObject(documents);
                    doc.Close();
                    Marshal.ReleaseComObject(doc);
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }
                else
                {
                    return (new ResponseModel { IsSucceed = false, ErrorMessage = outLocation + " file-ı artıq mövcuddur." });
                }

                if (!File.Exists(newPath))
                {
                    return (new ResponseModel { IsSucceed = false, ErrorMessage = newPath + " file-ı tapılmadı." });
                }
                return (new ResponseModel { Data = newPath, IsSucceed = true, ErrorMessage = string.Empty });
            }
            catch (Exception e)
            {
                return (new ResponseModel { IsSucceed = false, ErrorMessage = "File convert oluna bilmədi: " + e.Message });
            }
        }
        #endregion

        #region Excel to Pdf
        [HttpGet]
        [Route("api/WordToPdf/ConvertExcelToPdf")]
        public ResponseModel ConvertExcelToPdf(string fileLocation, string outLocation)
        {
            try
            {
                if (!File.Exists(outLocation))
                {

                    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                    app.Visible = false;
                    app.DisplayAlerts = false;
                    app.Interactive = false;
                    Workbooks workbooks = app.Workbooks;
                    Workbook wkb = workbooks.Open(fileLocation, ReadOnly: true);
                    wkb.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outLocation);

                    Marshal.ReleaseComObject(workbooks);
                    wkb.Close();
                    Marshal.ReleaseComObject(wkb);
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }
                else
                {
                    return (new ResponseModel { IsSucceed = false, ErrorMessage = outLocation + " file-ı artıq mövcuddur." });
                }

                if (!File.Exists(outLocation))
                {
                    return (new ResponseModel { IsSucceed = false, ErrorMessage = outLocation + " file-ı tapılmadı." });
                }
                return (new ResponseModel { Data = outLocation, IsSucceed = true, ErrorMessage = string.Empty });
            }
            catch (Exception e)
            {
                return (new ResponseModel { IsSucceed = false, ErrorMessage = "File convert oluna bilmədi: " + e.Message });
            }

        }
        #endregion

        #region PPT to Pdf

        [HttpGet]
        [Route("api/WordToPdf/ConvertPptToPdf")]
        public ResponseModel ConvertPptToPdf(string fileLocation, string outLocation)
        {
            try
            {
                if (!File.Exists(outLocation))
                {
                    Microsoft.Office.Interop.PowerPoint.Application app = new Microsoft.Office.Interop.PowerPoint.Application
                    {
                        Visible = MsoTriState.msoTrue
                    };
                    Presentations presentations = app.Presentations;
                    Presentation presentation = presentations.Open(fileLocation, ReadOnly: MsoTriState.msoCTrue);
                    presentation.ExportAsFixedFormat(outLocation, PpFixedFormatType.ppFixedFormatTypePDF);

                    Marshal.ReleaseComObject(presentations);
                    presentation.Close();
                    Marshal.ReleaseComObject(presentation);
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }
                else
                {
                    return (new ResponseModel { IsSucceed = false, ErrorMessage = outLocation + " file-ı artıq mövcuddur." });
                }

                if (!File.Exists(outLocation))
                {
                    return (new ResponseModel { IsSucceed = false, ErrorMessage = outLocation + " file-ı tapılmadı." });
                }

                return (new ResponseModel { Data = outLocation, IsSucceed = true, ErrorMessage = string.Empty });
            }
            catch (Exception e)
            {
                return (new ResponseModel { IsSucceed = false, ErrorMessage = "File convert oluna bilmədi: " + e.Message });
            }
        }

        #endregion

        #region Jpeg to Pdf
        [HttpGet]
        [Route("api/WordToPdf/ConvertImageToPdf")]
        public ResponseModel ConvertImageToPdf(string fileLocation, string outLocation)
        {
            try
            {
                if (!File.Exists(outLocation))
                {
                    var document = new iTextSharp.text.Document(PageSize.A4, 25, 25, 25, 25);
                    using (var stream = new FileStream(outLocation, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        PdfWriter.GetInstance(document, stream);
                        document.Open();
                        using (var imageStream = new FileStream(fileLocation, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            var image = Image.GetInstance(imageStream);
                            if (image.Height > PageSize.A4.Height - 25)
                            {
                                image.ScaleToFit(PageSize.A4.Width - 25, PageSize.A4.Height - 25);
                            }
                            else if (image.Width > PageSize.A4.Width - 25)
                            {
                                image.ScaleToFit(PageSize.A4.Width - 25, PageSize.A4.Height - 25);
                            }
                            image.Alignment = Element.ALIGN_MIDDLE;
                            document.Add(image);
                        }

                        document.Close();
                    }
                }
                else
                {
                    return (new ResponseModel { IsSucceed = false, ErrorMessage = outLocation + " file-ı artıq mövcuddur." });
                }

                return (new ResponseModel { Data = outLocation, IsSucceed = true, ErrorMessage = string.Empty });
            }
            catch (Exception e)
            {
                return (new ResponseModel { IsSucceed = false, ErrorMessage = "File convert oluna bilmədi: " + e.Message });
            }
        }

        #endregion
    }
}
