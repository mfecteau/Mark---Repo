using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text;
using System.Collections;
using MSXML2;

using Tspd.Businessobject;
using Tspd.Tspddoc;
using Tspd.Icp;
using Tspd.Utilities;
using Tspd.Rules;
using Tspd.Context;
using WorX;
using Word = Microsoft.Office.Interop.Word;


namespace TspdCfg.FastTrack.Rules
{
    public class TableRefCheckRule : ITSPDRule
	{
#if false
<rule id="Advisory05" type="CSHARP" displayName="TableRefCheckRule" source="TspdCfg.FastTrack.Rules.ReferenceCheckRule.,FTRules.dll" categories="StyleAdv" debug="false"/>
#endif
        public static Hashtable Sel_Section = new Hashtable();
        public static readonly string MY_ID = "TableRefCheckRule";
		private ArrayList finalResult = new ArrayList();
		private static string _ruleId = "";
        public TspdDocument doc = null;
		private static bool _debug = false;
        public ArrayList arr_Ranges = new ArrayList();
		public void Init(string ruleId, bool debug) 
		{
			_ruleId = ruleId;
			_debug = debug;
		}

		public string AdvisoryPrefix
		{
			get { return MY_ID; }
		}


        public bool canRunInStandaloneDocument
        {
            get { return true; }
        }
		public ICollection Run() 
		{
            ArrayList advisories = new ArrayList();
            try
            {
                ArrayList docSec = new ArrayList();
                ContextManager ctx = ContextManager.getInstance();
                doc = ctx.getActiveDocument();
                BusinessObjectMgr bom = doc.getBom();
                finalResult.Clear();

                int sRng = doc.getActiveWordDocument().ActiveWindow.Selection.Start;
                int eRng = doc.getActiveWordDocument().ActiveWindow.Selection.End;
                
                //Checking for List of Tables
                docSec = CheckListofTables(doc);
                string smsg = "";

                foreach (string str in docSec)
                {
                    smsg = str;
                    RuleAdvisory adv = new RuleAdvisory(
                        _ruleId, MY_ID + str, smsg);
                    advisories.Add(adv);
                }

                doc.getActiveWordDocument().ActiveWindow.Selection.SetRange(sRng,eRng); //to reset cursor.
                Word.Selection tmpSel = doc.getActiveWordDocument().ActiveWindow.Selection;
                tmpSel.Collapse(ref WordHelper.COLLAPSE_START); //resetting your cursor 
                
               object what = Word.WdGoToItem.wdGoToPage;
               object pageno = tmpSel.get_Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndAdjustedPageNumber);
               tmpSel.GoTo(ref what, ref VBAHelper.OPT_MISSING, ref VBAHelper.OPT_MISSING, ref pageno);

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return advisories;
		}

        public ArrayList CheckListofTables(TspdDocument currDoc)
        {
            /*This method will get list of Tables, Assuming only Only one list of table(figures) exists in doc. Look for TSD Section = "List of References". */

            ArrayList arrLOT = new ArrayList();
            ArrayList arrDoc = new ArrayList();

            //List all the Ranges you want to exclude. It uses a configuration file (txt).
            excludeRanges();

            try
            {
                if (currDoc.getActiveWordDocument().TablesOfFigures.Count > 0)
                {
                    Word.Range rng = currDoc.getActiveWordDocument().TablesOfFigures[1].Range;
                    IEnumerator TOCParas = rng.Paragraphs.GetEnumerator();
                    string[] secName = null;
                    while (TOCParas.MoveNext())
                    {
                        Word.Paragraph toc_ = (Word.Paragraph)TOCParas.Current;
                        if (toc_.Range.Text.Trim().Length != 0)
                        {
                            secName = null;
                            secName = toc_.Range.Text.Split('\t');
                            if (secName.Length > 2)
                            {
                                arrLOT.Add(secName[1]);
                            }
                            else
                            {
                                arrLOT.Add(secName[0]);
                            }
                            //MessageBox.Show(toc_.Range.Text + "        " +  toc_.Range.Bookmarks.ToString());
                        }
                    } //End While
                } //END IF ToF count


                arrDoc = getTableswithCaptionStyle();
                try
                {
                    //currDoc.unsetDocProtectionState();
                    string mesg = " is listed in the List of Tables, but the corresponding table cannot be identified within the document";
                    compareList(arrLOT, arrDoc, mesg);

                    mesg = " is found in document, but not found in List of Tables.";
                    compareList(arrDoc, arrLOT, mesg);
                }
                catch (Exception ex)
                {
                    Log.exception(ex, ex.Message);
                    MessageBox.Show(ex.ToString());
                }

 
            }
            catch (Exception ex)
            {
                Log.exception(ex, ex.Message);
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                
            }
            return finalResult;
        }


        private void excludeRanges()
        {
            DocumentSectionEntry dse = null;
            IElementBase worxEntry = null; // (IElementBase)dse.getWorxElement();
            arr_Ranges.Clear();
            try
            {

                string filepath = doc.getTrialProject().getTemplateDirPath() + "\\rules\\Protocol\\DocSection.txt";
                if (System.IO.File.Exists(filepath))
                {
                    System.IO.StreamReader strReader = new System.IO.StreamReader(filepath);
                    string currLine = null;
                    string reportName, elemPath;
                    bool flagReportheader = false;

                    //strReader.Peek(
                    while (strReader.Peek() >= 0)
                    {
                        currLine = strReader.ReadLine();

                        // compare currLine to Report Type and set FLAG = TRUE for reading lines then after.

                        if (currLine.IndexOf("ReportType") >= 0)
                        {
                            currLine = currLine.Substring(currLine.IndexOf("/"));
                            flagReportheader = true;
                        }
                        else
                        {
                            if (flagReportheader == true)
                            {
                                if (currLine.IndexOf("DocSection") >= 0)
                                {
                                    elemPath = currLine.Substring(currLine.IndexOf("=")+1);
                                    dse = doc.getDocSectionManager().getByPath(elemPath);
                                    if (dse != null)
                                    {
                                        worxEntry = (IElementBase)dse.getWorxElement();
                                        arr_Ranges.Add(worxEntry.WdRange);
                                    }
                                    //  sel_Doc_Section.Add(elemPath);
                                }
                                else if (currLine.IndexOf("LastSection") >= 0)
                                {
                                    elemPath = currLine.Substring(currLine.IndexOf("=")+1);
                                    dse = doc.getDocSectionManager().getByPath(elemPath);
                                    if (dse != null)
                                    {
                                        worxEntry = (IElementBase)dse.getWorxElement();

                                        //find the end of document
                                        object what = Word.WdGoToItem.wdGoToLine;
                                        object which = Word.WdGoToDirection.wdGoToLast;
                                        doc.getActiveWordDocument().ActiveWindow.Selection.GoTo(ref what, ref which, ref VBAHelper.OPT_MISSING, ref VBAHelper.OPT_MISSING);
                                        int endofDoc = doc.getActiveWordDocument().ActiveWindow.Selection.End;

                                        int startrng = worxEntry.WdRange.Start;
                                        Word.Selection tsel_ = doc.getActiveWordDocument().ActiveWindow.Selection;
                                        tsel_.SetRange(startrng, endofDoc);
                                        arr_Ranges.Add(tsel_.Range);
                                        tsel_.Collapse(ref WordHelper.COLLAPSE_START);
                                    }
                                }
                                else if (currLine.IndexOf("EndReport") >= 0)
                                {
                                    flagReportheader = false;
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                Log.exception(ex, "One or more document section path has been modified.");
                MessageBox.Show(ex.ToString());
            }
        }

        public ArrayList getTableswithCaptionStyle()
        {
            //this method, will build a List of all tables with headers and checking their field code. If no field code then check Ranges, which we donot want to include

            string[] secName = null;
            ArrayList arrtableName = new ArrayList();
            object oWordUnit = Word.WdUnits.wdParagraph;
            object o2 = 1; //Count for paragraph
            bool found = false;
            object start = 0;
            object end=0;
            try
            {

                foreach (Word.Table tbl in doc.getActiveWordDocument().Tables)
                {
                    try
                    {
                        tbl.Range.Collapse(ref WordHelper.COLLAPSE_START);
                        end= tbl.Range.Start;
                        start = tbl.Range.Previous(ref oWordUnit, ref o2).Start;
                        Word.Range thisRng = doc.getActiveWordDocument().Range(ref start, ref end).Duplicate;
                     //   Word.Style txtStyl = (Word.Style)thisRng.get_Style();

                            IEnumerator fdenum = thisRng.Fields.GetEnumerator();
                            while (fdenum.MoveNext())
                            {
                                Word.Field fd = (Word.Field)fdenum.Current;
                                if (fd.Code.Text.Trim().ToLower().Contains("seq table"))
                                {
                                    found = true;
                                    if (thisRng.Text.Trim().Length != 0)
                                    {
                                        secName = null;
                                        secName = thisRng.Text.Split('\t');
                                        if (secName.Length > 1)
                                        {
                                            arrtableName.Add(secName[1].Trim());
                                        }
                                        else
                                        {
                                            arrtableName.Add(secName[0].Trim());
                                        }
                                        //MessageBox.Show(toc_.Range.Text + "        " +  toc_.Range.Bookmarks.ToString());
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (!CheckRanges(thisRng))
                                {
                                    thisRng.Collapse(ref WordHelper.COLLAPSE_START);
                                     doc.getActiveWordDocument().ActiveWindow.Selection.SetRange(thisRng.Start, thisRng.End);
                                     Word.Selection selObj = doc.getActiveWordDocument().ActiveWindow.Selection;

                                    string strpageno = selObj.get_Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndAdjustedPageNumber).ToString();
                                    string msg = "A table on page " + strpageno + " has not been properly referenced in the List of Tables.";
                                    finalResult.Add(msg);
                                }
                            }
                        
                        found = false;
                    }
                    catch (Exception ex1)
                    {
                        //''ignore
                    }
                }
            }
            catch (Exception ex)
            {
// Word.WdInformation.wdActiveEndAdjustedPageNumber
            }
            finally
            { }
            return arrtableName;
        }


        private bool CheckRanges(Word.Range tblRng)
        {
            //This method checks if the given Range falls between any of the ranges we need to skip it.
            bool inRng = false;
            try
            {
                foreach (Word.Range dsRng in arr_Ranges)
                {
                    if ((tblRng.Start >= dsRng.Start) && (tblRng.End <= dsRng.End))
                    { //This method determines whether the range or selection returned by expression is contained in the specified Range by comparing the starting and ending character positions, as well as the story type.
                        inRng = true;    
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.exception(ex,ex.ToString());
                MessageBox.Show(ex.ToString());
            }
            return inRng;
        }
        public void compareList(ArrayList fromList, ArrayList toList, string mesg)
        {
            try
            {
                foreach (string findstr  in fromList)
                {
                    if (toList.IndexOf(findstr) < 0)
                    {
                        finalResult.Add(findstr + mesg);
                    }
                }
 
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

		public void CheckCount(string findString,int start,int end)
		{
		   try
            {
                bool found = false;
                object _optMissing = System.Reflection.Missing.Value;
                Word.Selection sel_ = doc.getActiveWordDocument().ActiveWindow.Selection;
                sel_.Find.Text = "<(" + findString.ToString() + ")>";
                sel_.Find.Forward = true;
                sel_.Find.Wrap = Word.WdFindWrap.wdFindContinue;
                sel_.Find.MatchCase = false;
                sel_.Find.MatchWholeWord = false;
                sel_.Find.MatchWildcards = true; 
             
                int currRange = 0;

                found = sel_.Find.Execute(
                      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                      ref _optMissing, ref _optMissing, ref _optMissing,
                      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing);
                currRange = sel_.Range.Start;

                bool exitFlag = true;

                while (exitFlag)
                {
                    //See if there is another instance of Reference text.
                    found = false;   //Resetting the flag
                    sel_ = doc.getActiveWordDocument().ActiveWindow.Selection;
                    sel_.Find.Text = "<(" + findString.ToString() + ")>";
                    sel_.Find.Forward = true;
                    sel_.Find.Wrap = Word.WdFindWrap.wdFindContinue;
                    sel_.Find.MatchCase = false;
                    sel_.Find.MatchWholeWord = false;
                    sel_.Find.MatchWildcards = true;
                 

                    found = sel_.Find.Execute(
                      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                      ref _optMissing, ref _optMissing, ref _optMissing,
                      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing);

                    if (found)
                    {
                        if (currRange == sel_.Range.Start)
                        {
                            exitFlag = false;
                            found = false;
                        }
                        else if (((sel_.Range.Start <= start) && (sel_.Range.End <= start))  || ((sel_.Range.Start >= end) && (sel_.Range.End <= end)))
                        {
                            exitFlag = false;
                            found = true;
                        }
                    }
                    else 
                    {
                        exitFlag = false;
                    }
                }

                if (found == false)
                {
                    finalResult.Add("The reference " + findString + " is identified in the List of References but is not used in the Protocol.");
                }
    
               //if ((found) && (currRange == sel_.Range.Start))
                    //{
                    //    exitFlag = false;
                    //}
                    //if ((found) && ((sel_.Range.Start >= start) && (sel_.Range.End <= end)))
                    //{
                    //    found = false;
                    //}
                //}

                //if (found == true)
                //{
                //    //See if there is another instance of Reference text.
                //    found = false;   //Resetting the flag
                //    sel_ = doc.getActiveWordDocument().ActiveWindow.Selection;
                //    sel_.Find.Text = findString.ToString(); sel_.Find.Forward = true;
                //    sel_.Find.Wrap = Word.WdFindWrap.wdFindContinue;
                //    sel_.Find.MatchCase = false;
                //    sel_.Find.MatchWholeWord = false;

                //    found = sel_.Find.Execute(
                //      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                //      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing,
                //      ref _optMissing, ref _optMissing, ref _optMissing,
                //      ref _optMissing, ref _optMissing, ref _optMissing, ref _optMissing);

                    

                //    if ((found) && (currRange == sel_.Range.Start))
                //    {
                //        found = false;
                //    }

                //    if ((found) && ((sel_.Range.Start >= start) && (sel_.Range.End <= end)))
                //    { 
                //            found = false;                        
                //    }
                //}
               
            }
            catch (Exception ex)
            {
                throw ex;
            }           
		}
	}

	}

