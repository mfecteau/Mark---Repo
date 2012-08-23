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
	public class Advisory02 : ITSPDRule
	{
#if false
<rule id="Advisory02" type="CSHARP" displayName="Advisory02" source="TspdCfg.FastTrack.Rules.Advisory02.,FTRules.dll" categories="stats" debug="false"/>
#endif
		public static Hashtable Sel_Section = new Hashtable();
		public static readonly string MY_ID = "Advisory02";

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
		
			foreach(string str in docSec)
            {
				 smsg = str;
			RuleAdvisory adv = new RuleAdvisory(
				_ruleId, MY_ID + str , 
				smsg);
			advisories.Add(adv);
			}
		
		

			return advisories;
		}

		public ArrayList compareDocSection(TspdDocument currDoc)
		{
		
			//Document Section that are renamed from the Orignial Name found in the Template
			ArrayList myList1 = new ArrayList();
			ArrayList myList2 = new ArrayList();

			

			IEnumerator sections = currDoc.getDocSectionList();
			
			while(sections.MoveNext())
			{
				DocumentSectionEntry dse = (DocumentSectionEntry)sections.Current;
				IElementBase worxEntry = (IElementBase)dse.getWorxElement();
				string labelEntry = "";
				
				if (dse.getDocumentState() == Tspd.Businessobject.ChooserEntry.DocumentState.InDoc)
				{
					if(dse.getElementLabel() != dse.getActualDisplayValue())
					{
						myList2.Add(dse.getElementLabel()+ " is renamed to " + dse.getActualDisplayValue());
					}
				}
			}		



			return myList2;

		}


	}
}
