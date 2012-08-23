using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections;
using System.Xml;

using Tspd.Businessobject;
using Tspd.Tspddoc;
using Tspd.Icp;
using Tspd.Utilities;
using Tspd.Rules;
using Tspd.Context;
using WorX;


namespace TspdCfg.FastTrack.Rules
{
	public class Advisory01 : ITSPDRule
	{
#if false
<rule id="Advisory01" type="CSHARP" displayName="Advisory01" source="TspdCfg.FastTrack.Rules.Advisory01.,FTRules.dll" categories="testcs" debug="false"/>
#endif
		public static Hashtable Sel_Section = new Hashtable();
		public static readonly string MY_ID = "Advisory01";

		private static string _ruleId = "";
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
			ArrayList docSec =  new ArrayList();

			ContextManager ctx = ContextManager.getInstance();
			TspdDocument doc = ctx.getActiveDocument();
			BusinessObjectMgr bom = doc.getBom();

			docSec = compareDocSection(doc);

			string smsg="" ;
			foreach (DocumentSectionEntry dse in docSec) 
			{
				 smsg = "Level " + dse.getSectionLevel() + " - "  + dse.getElementLabel() + " is not included in a Document.";	
				RuleAdvisory adv = new RuleAdvisory(
				_ruleId, MY_ID + dse.getElementPath(), 
				smsg);
			    advisories.Add(adv);		
			}
		

			return advisories;
		}

		public ArrayList compareDocSection(TspdDocument currDoc)
		{
			//Document Section that are unchecked/removed a document section that is included in the template.

			ArrayList myList1 = new ArrayList();
			ArrayList myList2 = new ArrayList();
			IEnumerator sections = currDoc.getDocSectionList();
			
			while(sections.MoveNext())
			{				
				DocumentSectionEntry dse = (DocumentSectionEntry)sections.Current;
				IElementBase worxEntry = (IElementBase)dse.getWorxElement();
				
				if (dse.getDocumentState() == Tspd.Businessobject.ChooserEntry.DocumentState.InDoc)
				{
					
				}
				else
				{
                    if (dse.getSectionLevel() <= 3)  //Limit to 3 Levels
                    {
                        myList2.Add(dse);
                    }
				}
			}
			return myList2;

		}


	}
}
