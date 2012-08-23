using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices;
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
    public class DocSecCheckRule : ITSPDRule
    {
#if false
<rule id="Advisory05" type="CSHARP" displayName="DocSecCheckRule" source="TspdCfg.FastTrack.Rules.DocSecCheckRule.,FTRules.dll" categories="StyleAdv" debug="false"/>
#endif
        public static Hashtable Sel_Section = new Hashtable();
        public static readonly string MY_ID = "DocSecCheckRule";
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

            docSec = CheckDocSection(doc);
            string smsg = "";

            foreach (string str in docSec)
            {
                smsg = str;
                RuleAdvisory adv = new RuleAdvisory(
                    _ruleId, MY_ID + str, smsg);
                advisories.Add(adv);
            }
            
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());                
            }
            return advisories;

        }

        public ArrayList CheckDocSection(TspdDocument currDoc)
        {
            /*This method will check if a particular reference exists in document more then Once. Reference is identified
            by style 'Heading 9'. Look for TSD Section = "List of References". */

            IEnumerator sections = currDoc.getDocSectionList();
           // DocumentSectionEntry selSection;
     //       IElementBase sel_WorxEntry = null;


            try
            { 
            while (sections.MoveNext())
            {
                DocumentSectionEntry dse = (DocumentSectionEntry)sections.Current;

                object tr = true;

                if (dse is CustomSectionEntry)
                {
                    continue;
                }
                else
                {

                    if (dse.getValueForNode("@required").ToString() == "true")
                    {
                        if (dse.getDocumentState() != Tspd.Businessobject.ChooserEntry.DocumentState.InDoc)
                        {
                            finalResult.Add("Section " + dse.getElementLabel().ToString() + " is a required section but it has been removed from the protocol document.");
                        }
                        if (dse.getElementLabel().ToLower().Trim() != dse.getActualDisplayValue().ToLower().Trim())
                        {
                            finalResult.Add("Section " + dse.getActualDisplayValue().ToString() + " is a required section but has been renamed from " + dse.getElementLabel().ToString() + " within the protocol document.");
                        }
                    }
                }
            }
            }
            catch (Exception ex)
            {
                Log.exception(ex, ex.Message);
                    MessageBox.Show(ex.ToString());
            }

            return finalResult;
        }


    }

}

