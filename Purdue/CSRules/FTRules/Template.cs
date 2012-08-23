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

namespace TspdCfg.FastTrack.Rules
{
	public class Template : ITSPDRule
	{
#if false
<rule id="Template" type="CSHARP" displayName="Template" source="TspdCfg.FastTrack.Rules.Template,FTRules.dll" categories="testcs" debug="false"/>
#endif

		public static readonly string MY_ID = "Template";

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

			ContextManager ctx = ContextManager.getInstance();
			TspdDocument doc = ctx.getActiveDocument();
			BusinessObjectMgr bom = doc.getBom();

			string smsg = doc.getDocumentTitle() + ", " + DateTime.Now.ToString();
			RuleAdvisory adv = new RuleAdvisory(
				_ruleId, MY_ID, // add element path to my_id if based on a specific object
				smsg);
			advisories.Add(adv);

			return advisories;
		}
	}
}
