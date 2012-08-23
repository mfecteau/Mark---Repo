using System;
using System.Diagnostics;
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


namespace TspdCfg.FastTrack.Rules
{
	public class Advisory04 : ITSPDRule
	{
#if false
<rule id="Advisory04" type="CSHARP" displayName="Advisory04" source="TspdCfg.FastTrack.Rules.Advisory04.,FTRules.dll" categories="testcs" debug="false"/>
#endif
		public static Hashtable Sel_Section = new Hashtable();
		public static readonly string MY_ID = "Advisory04";

		private static string _ruleId = "";
		private static bool _debug = false;

		public void Init(string ruleId, bool debug) 
		{
			_ruleId = ruleId;
			_debug = debug;
		}

        public bool canRunInStandaloneDocument
        {
            get { return true; }

        }
		public string AdvisoryPrefix
		{
			get { return MY_ID; }
		}

		public ICollection Run() 
		{
			ArrayList advisories = new ArrayList();
			ArrayList docSec =  new ArrayList();

			ContextManager ctx = ContextManager.getInstance();
			TspdDocument doc = ctx.getActiveDocument();
			BusinessObjectMgr bom = doc.getBom();

			docSec = CompareDocSectionLevel(doc);

			string smsg="" ;
		
			foreach (string str in docSec) 
            {
				 smsg = str;
			RuleAdvisory adv = new RuleAdvisory(
				_ruleId, MY_ID + str , smsg);
			advisories.Add(adv);
			}
			

			return advisories;
		}

		public ArrayList CompareDocSectionLevel(TspdDocument currDoc)
		{
		
			//Get the Docuement Sections cannot exceed Level 4.

			ArrayList myList2 = new ArrayList();
		
			IEnumerator sections = currDoc.getDocSectionList();
			
			while(sections.MoveNext())
			{
				DocumentSectionEntry dse = (DocumentSectionEntry)sections.Current;
				IElementBase worxEntry = (IElementBase)dse.getWorxElement();
				if (dse.getSectionLevel() > 4)
				{
					if (dse.getDocumentState() == Tspd.Businessobject.ChooserEntry.DocumentState.InDoc)
					{
						myList2.Add(dse.getSectionNumber() + " - "  + dse.getActualDisplayValue() + " exceeds 4 levels.");
					}
				}
                			
			}

			
			return myList2;

		}


	}
}
