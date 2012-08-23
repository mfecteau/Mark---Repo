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
	public class Advisory03 : ITSPDRule
	{
#if false
<rule id="Advisory03" type="CSHARP" displayName="Advisory03" source="TspdCfg.FastTrack.Rules.Advisory03.,FTRules.dll" categories="testcs" debug="false"/>
#endif
		public static Hashtable Sel_Section = new Hashtable();
		public static readonly string MY_ID = "Advisory03";

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

			docSec = GetCustomSections(doc);

			string smsg="" ;
		
			foreach (DocumentSectionEntry dse in docSec) 
            {
				 smsg = dse.getElementLabel() + " is Custom Document Section.";
			RuleAdvisory adv = new RuleAdvisory(
				_ruleId, MY_ID + dse.getElementPath() , smsg);
			advisories.Add(adv);
			}
			

			return advisories;
		}

		public ArrayList GetCustomSections(TspdDocument currDoc)
		{
		
			//Get the CUSTOM Docuement Section.  OR Document Section that are created by Author and used in the Protocol Document

			ArrayList myList2 = new ArrayList();
			IEnumerator sections = currDoc.getDocSectionList();
			
			while(sections.MoveNext())
			{
				DocumentSectionEntry dse = (DocumentSectionEntry)sections.Current;
				IElementBase worxEntry = (IElementBase)dse.getWorxElement();
				string labelEntry = "";
				
				if(dse is CustomSectionEntry)
				{
					myList2.Add(dse);
				}
			}
			return myList2;

		}


	}
}
