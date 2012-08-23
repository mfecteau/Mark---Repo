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
	public class TestRule3 : ITSPDRule
	{
#if false
<rule id="testcs3" type="CSHARP" displayName="Test CSharp 3" source="TspdCfg.FastTrack.Rules.TestRule3,FTRules.dll" categories="testcs" debug="false"/>
#endif

		public static readonly string MY_ID = "TestRule3";

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
				_ruleId, MY_ID, 
				smsg);
			advisories.Add(adv);

			return advisories;
		}
	}
}
