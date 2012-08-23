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
    public class TOCReferenceCheckRule : ITSPDRule
	{
#if false
<rule id="Advisory05" type="CSHARP" displayName="ReferenceCheckRule" source="TspdCfg.FastTrack.Rules.ReferenceCheckRule.,FTRules.dll" categories="StyleAdv" debug="false"/>
#endif
        public static Hashtable Sel_Section = new Hashtable();
        public static readonly string MY_ID = "TOCReferenceCheckRule";
		private ArrayList finalResult = new ArrayList();
		private static string _ruleId = "";
        public TspdDocument doc = null;
		private static bool _debug = false;

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

                //Checking for Table of Contents
                docSec = CheckReference(doc);

              
                string smsg = "";

                foreach (string str in docSec)
                {
                    smsg = str;
                    RuleAdvisory adv = new RuleAdvisory(
                        _ruleId, MY_ID + str, smsg);
                    advisories.Add(adv);
                }

                //doc.getActiveWordDocument().ActiveWindow.Selection.SetRange(rngStart, rngEnd); //to reset cursor.
                //doc.getActiveWordDocument().ActiveWindow.Selection.Select();
               // return advisories;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return advisories;
		}

   


		public ArrayList CheckReference(TspdDocument currDoc)
		{
            /*This method will check if TOC matches the actual Document sections.. */

           ArrayList arrTOC = new ArrayList();
           ArrayList arrDocSec = new ArrayList();
         
            //Collecting all TOC in ArrayList

           Word.Range tocRange = currDoc.getActiveWordDocument().TablesOfContents[1].Range;
           IEnumerator TOCParas = tocRange.Paragraphs.GetEnumerator();
           string[] secName =null;
            while (TOCParas.MoveNext())
            {
                Word.Paragraph toc_ = (Word.Paragraph)TOCParas.Current;
                if (toc_.Range.Text.Trim().Length != 0)
                {
                    secName = null;
                    secName = toc_.Range.Text.Split('\t');
                    if (secName.Length > 2)
                    {
                        arrTOC.Add(secName[1]);
                    }
                    else
                    {
                        arrTOC.Add(secName[0]);
                    }
                    //MessageBox.Show(toc_.Range.Text + "        " +  toc_.Range.Bookmarks.ToString());
                }
            }

            IEnumerator sections = currDoc.getDocSectionList();
            DocumentSectionEntry selSection;
            IElementBase sel_WorxEntry =null;

            while (sections.MoveNext())
            {
                DocumentSectionEntry dse = (DocumentSectionEntry)sections.Current;
                if (dse.getDocumentState() == Tspd.Businessobject.ChooserEntry.DocumentState.InDoc)
                {
                    arrDocSec.Add(dse.getElementLabel());
                }
            }

            try
            {
                //currDoc.unsetDocProtectionState();
                string mesg = " is found in Trial document, but not found in Table of Contents";
                compareList(arrDocSec, arrTOC, mesg);

                mesg = " is found in Table of Contents, but not found in Trial Document.";
                compareList(arrTOC, arrDocSec, mesg);
            }
            catch (Exception ex)
            {
                Log.exception(ex, ex.Message);
                MessageBox.Show(ex.ToString());
            }
            finally
                {
                 //   currDoc.setDocProtectionState();
                }
				return finalResult;
			}

   

        public void compareList(ArrayList fromList, ArrayList toList, string mesg)
        {
            try
            {
                
                ArrayList arr_toList = Convert2Lower(toList);  // --> to convert all items in LOWER CASE
                
                foreach (string findstr  in fromList)
                {
                    if (arr_toList.IndexOf(findstr.ToLower()) < 0)
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

        private ArrayList Convert2Lower(ArrayList sourceList)
        {
            string stext = null;
            for(int i=0; i<=sourceList.Count-1; i++)
            {
                stext = sourceList[i].ToString().ToLower();
                sourceList[i] = stext;
            }

            return sourceList;
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
    
              
               
            }
            catch (Exception ex)
            {
                throw ex;
            }           
		}
	}

	}

