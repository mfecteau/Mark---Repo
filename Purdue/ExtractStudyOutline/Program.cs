using System;
using System.Collections;
using System.Windows.Forms;
using System.Data;
using Tspd.Utilities;
using Word = Microsoft.Office.Interop.Word;

namespace ExtractStudyOutline
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	class Program
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main(string[] args)
		{
			try 
			{
				extractStudyOutline(args[0]);
			}
			catch (Exception ex) 
			{
				ErrorForm dlg = new ErrorForm();

				dlg.setErrorMessage(ex.Message + "\r\n" + ex.StackTrace);
				dlg.ShowDialog();
			}
		}

        private static void extractStudyOutline(string fileName)
        {
            Word.Application wdApp = null;

            object oFalse = false;
            object oTrue = true;

            object missing = System.Reflection.Missing.Value;
            object wdCollapseEnd = Word.WdCollapseDirection.wdCollapseEnd;

            wdApp = new Word.Application();
            // wdApp.Visible = true;

            object oFileName = fileName;
            object password = "";

            wdApp.Application.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            Word.WdPasteOptions _currSetting = wdApp.Options.PasteFormatBetweenStyledDocuments;


            wdApp.Options.PasteFormatBetweenStyledDocuments = Word.WdPasteOptions.wdKeepSourceFormatting;



            Word.Document wdDocSource = wdApp.Documents.Open(
                ref oFileName, ref oFalse, ref oFalse, ref oFalse, ref password,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref oFalse, ref missing,
                ref missing, ref missing, ref missing);

            try
            {
                wdDocSource.Sections.First.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
                wdDocSource.Sections.First.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
            }
            catch (Exception e)
            {
                //
            }
            Word.Document wdDoc1 = wdApp.Documents.Add(ref missing, ref missing, ref missing, ref oFalse);
            Word.Document wdDoc2 = wdApp.Documents.Add(ref missing, ref missing, ref missing, ref oFalse);         


            object oStart = wdDocSource.Sections.First.Range.Start;
            object oEnd = wdDocSource.Sections.First.Range.End;
            Word.Range mySecRng = null;
            string str = "";
            bool _found = false;
            int START_ = 0;
            int secCnt_sdd = 0;
            object wdEnd = Word.WdCollapseDirection.wdCollapseEnd;

            IEnumerator secEnum = wdDocSource.Sections.GetEnumerator();
            while (secEnum.MoveNext())
            {
                Word.Section currsec = (Word.Section)secEnum.Current;
                mySecRng = currsec.Range;
                str = mySecRng.Text.Trim();
                str = str.Replace("\r", "");
                str = str.Replace("\t", "");

                Log.trace(str.Substring(0, 10));

                if (str.Substring(0, 10).ToLower().StartsWith("title page"))
                {
                 //   Log.trace(str);
                    secCnt_sdd = currsec.Index - 1; 
                    _found = true;
                    START_ = currsec.Range.Start;
                    oEnd = mySecRng.End;
                    break;
                }
               
            }
            if (_found)
            {
                Log.trace("SDD Starts from " + START_);
                Log.trace("SDD Ends @ :" + oEnd);
            }
            else
            {
                Log.trace("SDD not found");
            }
            Word.Range wdRange1 = wdDocSource.Range(ref oStart, ref oEnd);
            wdRange1.Copy();


            Word.Range iRng = wdDoc1.ActiveWindow.Selection.Range;
            Word.Section reportSec;


            object[] styleArray;
            styleArray = new object[] { "tspdHV", "tspdNV", "tspdRHV", "tspdRNV", "tspdD", "tspdDR", "tspdHDR" };

            if (_found)
            {
                iRng.PasteAndFormat(Word.WdRecoveryType.wdPasteDefault);
                Word.Selection sel_ = wdDoc1.ActiveWindow.Selection;
                sel_.WholeStory();
                RemoveGrayBackcolor(sel_);
                ////			//removing styles for elements in the report document.

                foreach (object oStyle in styleArray)
                {
                    RemoveStyle(oStyle, wdDoc1);
                }

                //Copying and Pasting Headers in SDD Document
                secEnum.Reset();
                int idx = 0;            
                while (secEnum.MoveNext())
                {
                    Word.Section secList = (Word.Section)secEnum.Current;

                    if (secList.Index <= secCnt_sdd)
                    {
                        try
                        {
                            idx++;  //Starts with 1.
                            secList.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Copy();
                            reportSec = wdDoc1.Sections[idx];
                            reportSec.PageSetup.DifferentFirstPageHeaderFooter = 0;
                            reportSec.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                            reportSec.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.WholeStory();
                            reportSec.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paste();
                            //sel_.Collapse(ref WordHelper.COLLAPSE_END);
                            //sel_.TypeBackspace();      //to remove extra paragraph               
                            sel_.Collapse(ref wdEnd);


                            secList.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Copy();
                            reportSec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                            reportSec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paste();
                            sel_.Collapse(ref wdEnd);
                            //sel2_.TypeBackspace();
                            //sel2_.Collapse(ref wdEnd);

                        }
                        catch (Exception ex)
                        {
                            Log.exception(ex, ex.StackTrace);
                            //System.Windows.Forms.MessageBox.Show(ex.ToString() + "  ___ " + secList.Index);
                        }
                    }
                }


            }

            



            object f1 = fileName.Substring(0, fileName.Length - 4) + "_1.doc";

            wdDoc1.SaveAs(ref f1, ref missing, ref oFalse, ref password,
                ref oFalse, ref password, ref oFalse, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing);


            ///Protocol Document -- 
            
            oStart = START_;
            oEnd = wdDocSource.Content.End;
            Word.Range wdRange2 = wdDocSource.Range(ref oStart, ref oEnd);
            wdRange2.Copy();

            //wdDoc2.ActiveWindow.Selection.SetRange(0, 0);
            Word.Range tmpRng = wdDoc2.ActiveWindow.Selection.Range;
            //tmpRng.Paste();
            tmpRng.PasteAndFormat(Word.WdRecoveryType.wdPasteDefault);


            Word.Selection sel2_ = wdDoc2.ActiveWindow.Selection;
            sel2_.WholeStory();
            RemoveGrayBackcolor(sel2_);

            ////			//removing styles for elements in the report document.
            foreach (object oStyle in styleArray)
            {
                RemoveStyle(oStyle, wdDoc2);
            }

            /// Copying and Pasting Headers and Footers in Protocol Document.
            /// 

            
            sel2_.Collapse(ref wdEnd);
            secEnum.Reset();

			int idx2 =0;
			
            while (secEnum.MoveNext())
            {
                Word.Section secList = (Word.Section)secEnum.Current;
                
                if (secList.Index > secCnt_sdd)
                {
                    try
                    {
                        idx2++;  //Starts with 1.
                        secList.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Copy();                        
                        reportSec = wdDoc2.Sections[idx2];
                        reportSec.PageSetup.DifferentFirstPageHeaderFooter = 0;
                        reportSec.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                        reportSec.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.WholeStory();
                        reportSec.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paste();
                        //sel_.Collapse(ref WordHelper.COLLAPSE_END);
                        //sel_.TypeBackspace();      //to remove extra paragraph               
                        sel2_.Collapse(ref wdEnd);
                        
                        
                        secList.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Copy();                        
                        reportSec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
                        reportSec.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paste();
                        sel2_.Collapse(ref wdEnd);
                        //sel2_.TypeBackspace();
                        //sel2_.Collapse(ref wdEnd);

                    }
                    catch (Exception ex)
                    {
                        Log.exception(ex, ex.StackTrace);
                        //System.Windows.Forms.MessageBox.Show(ex.ToString() + "  ___ " + secList.Index);
                    }
                }
            }


            object f2 = fileName.Substring(0, fileName.Length - 4) + "_2.doc";

            wdDoc2.SaveAs(ref f2, ref missing, ref oFalse, ref password,
                ref oFalse, ref password, ref oFalse, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing);

            try
            {
                object dataObj = Clipboard.GetDataObject();
                Clipboard.SetDataObject("", true);
            }
            catch (Exception) { }

            try
            {
                string pdfFilename = "";
                pdfFilename = fileName.Substring(0, fileName.Length - 4) + "_1.pdf";

                if (System.IO.File.Exists(pdfFilename))
                { //DELETE the file if it exists.
                    System.IO.File.Delete(pdfFilename);
                }

                Tspd.Utilities.PDFConverter.convertDocToPdf(wdDoc1, pdfFilename, "Designer", false, false);

                pdfFilename = fileName.Substring(0, fileName.Length - 4) + "_2.pdf";
                if (System.IO.File.Exists(pdfFilename))
                {
                    System.IO.File.Delete(pdfFilename);
                }
                Tspd.Utilities.PDFConverter.convertDocToPdf(wdDoc2, pdfFilename, "Designer", false, false);
            }
            catch (Exception ex)
            {
                Log.exception(ex, ex.Message + ex.StackTrace);
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
            finally
            {
                wdApp.Options.PasteFormatBetweenStyledDocuments = _currSetting;
                wdDocSource.Close(ref oFalse, ref missing, ref missing);
                wdDoc1.Close(ref oFalse, ref missing, ref missing);
                wdDoc2.Close(ref oFalse, ref missing, ref missing);
                wdApp.Quit(ref oFalse, ref missing, ref missing);
            }
        }


        private static void RemoveGrayBackcolor(Word.Selection sel_)
        {
            sel_.Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;

            sel_.Range.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            sel_.Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            sel_.Range.Borders.InsideColor = Word.WdColor.wdColorBlack;
            sel_.Range.Borders.OutsideColor = Word.WdColor.wdColorBlack;

            sel_.Range.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;

            IEnumerator tables = sel_.Tables.GetEnumerator();
            while (tables.MoveNext())
            {
                Word.Table currTable = tables.Current as Word.Table;
                currTable.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
            }
            //sel_.ParagraphFormat.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
        }

        private static void RemoveStyle(object stylename, Word.Document wdDoc_)
        {
            try
            {
                Word.Style foundStyle = wdDoc_.Styles.get_Item(ref stylename);
                //Removing the TSPD Stlyes from the Report Document.
                foundStyle.Font.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                foundStyle.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                foundStyle.ParagraphFormat.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                //				foundStyle.Font.Borders.Item(Word.WdBorderType.wdBorderBottom).LineStyle= Word.WdLineStyle.wdLineStyleNone;
                //				foundStyle.Font.Borders.Item(Word.WdBorderType.wdBorderLeft).LineStyle= Word.WdLineStyle.wdLineStyleNone;
                //				foundStyle.Font.Borders.Item(Word.WdBorderType.wdBorderRight).LineStyle= Word.WdLineStyle.wdLineStyleNone;

            }
            catch (Exception ex) { }

        }
	}
}
