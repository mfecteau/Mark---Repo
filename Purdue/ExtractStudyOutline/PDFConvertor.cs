using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing.Printing;
using System.Reflection;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;


namespace ExtractStudyOutline
{
    
        /// <summary>
        /// PDFConverter is used to convert a .doc file into a .pdf file.
        /// The class has a static method convertDocToPdf which can convert a Word.Document to a pdf file.
        /// </summary>
        public class PDFConverter
        {
            private CDIntfEx.CDIntfExClass _pdfPrinter;
            private const string PDF_PRINTER_NAME = "TSD PDF Converter"; // do NOT change it unless you change installer as well.
            private const string LICENSEE = "Medidata Solutions Inc";
            private const string LICENSE_CODE =
                "07EFCDAB010001007E9F201B7096A71240190AD259AA5E784379C9BC8380DCF56CDFF58D4B36BEAA16E54194F2726A01320C96F4BE99A34C057E465A2D315F";

            private const int FNO_NoPrompt = 0x1;
            private const int FNO_UseFileName = 0x2;
            private const int FNO_EmbedFonts = 0x10;
            private const int FNO_MultilingualSupport = 0x80;
            private const int FNO_FullEmbed = 0x200;

            private const string AUTHOR_PWD = "FASTTRACKTSD"; // used in encryption

            private PrinterOptions _opts = null;

            private PDFConverter()
            {
               // Log.log(System.Diagnostics.TraceLevel.Info, "Enter in PDFConverter(), create a PDF printer object.");

                try
                {
                    _opts = new PrinterOptions();
                    // initialize the PDF printer
                    _pdfPrinter = new CDIntfEx.CDIntfExClass();
                    _pdfPrinter.DriverInit(PDF_PRINTER_NAME);
                    // enable printer after init driver.
                    enablePrinter();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Unable to create a PDFConverter." + ex.Message + ex.StackTrace);
                    //Log.exception(ex, "Unable to create a PDF printer. " + ex.Message + ex.StackTrace);
                    //throw new FTException("Unable to create a PDFConverter.", ex);
                }
            }

            private void setPrinterOptions(PrinterOptions opts)
            {
                _opts = opts;
            }

            /// <summary>
            /// To convert a doc to a pdf in either background or foreground thread.
            /// used from the posting call
            /// </summary>
            /// <param name="wdDoc"></param>
            /// <param name="pdfOutputFileName"></param>
            /// <param name="printInBackground"></param>
            public static void convertDocToPdf(Word.Document wdDoc,
                string pdfOutputFileName, string author,
                bool printInBackground, bool useRange)
            {
                //Log.log(System.Diagnostics.TraceLevel.Info, "Enter in PDFConverter::convertDocToPdf()," +
                //    " pdfOutputFileName is: " + pdfOutputFileName);

                PDFConverter pdfConverter = new PDFConverter();
                try
                {
                    if (useRange == true)
                    {
                        //commented the below 2 lines, as we dont need to give user selection for pages to convert.

                        ////PrinterSettings printOpts = new PrinterSettings();
                        ////pdfConverter.setPrinterOptions(printOpts.getPrinterOptions());

                        /* this would be nice... pity it crashes
                        Word.Application word = wdDoc.Application;
                        object missing = Missing.Value;
                        Word.Dialog dlg = word.Dialogs.Item(Word.WdWordDialog.wdDialogFilePrint);
                        dlg.Show(ref missing);
                        */
                    }
                    pdfConverter.convertDocToPdfInternal(wdDoc, pdfOutputFileName, author, printInBackground);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Unable to convert document to pdf file. " + ex.Message + ex.StackTrace);
                    //throw new FTException("Unable to convert document to pdf file.", ex);
                }
                finally
                {
                    // need to disconnect printer no matter if the convertion succeeds or not.
                    pdfConverter.disconnectPrinter();
                }
            }

            /// <summary>
            /// Enable the PDF Printer
            /// </summary>
            private void enablePrinter()
            {
             //   Log.log(System.Diagnostics.TraceLevel.Info, "Enter in PDFConverter::enablePrinter().");

                try
                {
                    _pdfPrinter.EnablePrinter(LICENSEE, LICENSE_CODE);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Unable to enablePrinter(). " + ex.Message + ex.StackTrace);
                    //throw new FTException("Unable to enable a PDF printer.", ex);
                }
            }

            /// <summary>
            /// To disconnect PDF printer
            /// </summary>
            private void disconnectPrinter()
            {
               // Log.log(System.Diagnostics.TraceLevel.Info, "Enter in PDFConverter::disconnectPrinter().");

                try
                {
                    _pdfPrinter.DriverEnd();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Unable to disconnectPrinter(). " + ex.Message + ex.StackTrace);
                    //throw new FTException("Unable to disconnect a PDF printer.", ex);
                }
            }

            /// <summary>
            /// To convert a word doc to a pdf file.
            /// This takes a boolean argument printInBackground to indicate if we want to do 
            /// the printing pdf job in background or foreground. Please note, currently we
            /// decided Not to set the security stuff, hence we can have a choice of print in
            /// background or foreground. If we decide to add security setting later on, then
            /// foreground printing is a must.
            /// </summary>
            /// <param name="?"></param>
            /// <param name="pdfFileName">PDF output file name with full path</param>
            private void convertDocToPdfInternal(Word.Document wdDoc,
                        string pdfOutputFileName, string author,
                        bool printInBackground)
            {
               // Log.log(System.Diagnostics.TraceLevel.Info, "Enter in PDFConverter::convertDocToPdfInternal()," +
                 //   " pdfOutputFileName is: " + pdfOutputFileName);

                String tempFile = Path.GetTempFileName();

                // set pdf printer setting and output file
                _pdfPrinter.FileNameOptionsEx =
                    FNO_NoPrompt |
                    FNO_UseFileName |
                    FNO_EmbedFonts |
                    FNO_MultilingualSupport;

                _pdfPrinter.DefaultFileName = tempFile;

                _pdfPrinter.JPEGCompression = true;
                _pdfPrinter.Resolution = 100;

                // remember the original printer for word
                string originalPrinter = wdDoc.Application.ActivePrinter;

                object oldWarnValue = null;

                try
                {
                    // set the PDF printer as active printer
                    wdDoc.Application.ActivePrinter = PDF_PRINTER_NAME;

                    // call enablePrinter right before each printing.
                    enablePrinter();

                    //Word.Dialog dlg = wdDoc.Dialogs.Item(Word.WdWordDialog.wdDialogFilePrint);
                    //dlg.Show(ref Missing.Value);
                    // 'print' the doc to the pdf printer
                    object missing = Missing.Value;
                    object false1 = false;
                    object true1 = true;
                    object range = _opts.getPrintRange(); //FromTo; //CurrentPage;
                    object pages = _opts.getPrintPages();

                //    oldWarnValue = WordHelper.disableSavingWithCommentsWarning(wdDoc.Application);

                    if (printInBackground)
                    {
                        // Print in background - this may be ok when we do NOT set security stuff.
                        wdDoc.PrintOut(ref true1, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing);
                    }
                    else
                    {
                        // Print in foreground - this is a Must if we want to set security stuff
                        // note: we pass false in the 1st argument to force word does the printing job in
                        // foreground thread. This is important since we don't know when the printing is 
                        // done if it is done in backgroud thread, which will cause blocking / failure
                        // when we try to open the pdf file for setting security stuff in the next step.			
                        wdDoc.PrintOut(ref false1, ref missing, ref range, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref pages, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing);
                    }
                    // restore thr pdf printer setting
                    _pdfPrinter.FileNameOptions = 0;
                }
                catch (Exception e)
                {
                    MessageBox.Show("Got an exception in convertDocToPdfInternal().");
                    // let us just throw the exception to let caller handle it.
                    throw e;
                }
                finally
                {
                    // restore the word printer to the original one, no matter if pdf succeeds or fails
                    wdDoc.Application.ActivePrinter = originalPrinter;

                   // WordHelper.retoreSavingWithCommentsWarningValue(wdDoc.Application, oldWarnValue);
                }

                setPDFFileSecurity(tempFile, author);

                // Copy the temp file over now
                File.Copy(tempFile, pdfOutputFileName);
            }

            /// <summary>
            /// This is to set security stuff to a pdf doc.
            /// The security we set here includes:
            ///		- Not allow editing
            ///		- Not allow copying
            ///		- Allow printing
            ///		- Allow adding/changing notes
            ///		- user does Not need a password to view the pdf document.
            /// </summary>
            /// <param name="pdfFileName">PDF file name with full path</param>
            private void setPDFFileSecurity(string pdfFileName, string author)
            {
                //Log.log(System.Diagnostics.TraceLevel.Info, "Enter in PDFConverter::setPDFFileSecurity()," +
                //    " pdfOutputFileName is: " + pdfFileName);

                try
                {
                    // open the pdf file
                    CDIntfEx.Document pdfDoc = new CDIntfEx.DocumentClass();
                    pdfDoc.Open(pdfFileName);

                    // set creator & author?
                    //pdfDoc.Creator = "Fast Track Systems Inc.";
                    pdfDoc.Author = author;

                    // set license Key
                    pdfDoc.SetLicenseKey(LICENSEE, LICENSE_CODE);

                    //Encrypt the file, we enable printing (+4), adding notes (+32), but disable editing (+8)
                    //and copying (+16). We don't set a user password here. Go to Amyuni developer manual for details (P121)
                    pdfDoc.Encrypt(AUTHOR_PWD, "", -64 + 4 + 32);

                    // save thh encrypted document
                    pdfDoc.Save(pdfFileName);
                }
                catch (Exception ex)
                {
                    //Log.exception(ex, "Unable to set PDF security. " + ex.Message + ex.StackTrace);
                    //throw new FTException("Unable to set PDF security.", ex);
                }
            }
        }
        public class PrinterOptions
        {
            object range_ = Word.WdPrintOutRange.wdPrintAllDocument;
            object pages_ = Missing.Value;

            public PrinterOptions() { }

            public PrinterOptions(object range, object pages)
            {
                range_ = range;
                pages_ = pages;
            }

            public object getPrintRange()
            {
                return range_;
            }

            void setPrintRange(object r)
            {
                range_ = r;
            }

            public object getPrintPages()
            {
                return pages_;
            }

            void setPrintPages(object p)
            {
                pages_ = p;
            }
        }  
}

