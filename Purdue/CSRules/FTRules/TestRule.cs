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
	public class TestRule : ITSPDRule
	{
#if false
<rule id="testcs1" type="CSHARP" displayName="Test CSharp 1" source="TspdCfg.FastTrack.Rules.TestRule,FTRules.dll" categories="testcs" debug="false"/>
#endif

		public static readonly string MY_ID = "TestRule";

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

			runTest1(advisories);

			return advisories;
		}

		private void runTest1(ArrayList advisories) 
		{

			ContextManager ctx = ContextManager.getInstance();
			TspdDocument doc = ctx.getActiveDocument();
			BusinessObjectMgr bom = doc.getBom();

			IcpInstanceManager icpInstMgr = doc.getBaseInstanceMgr();

			bool isOther;

			// Path to custom variable Child Bearing Potential
			string cbpPath = "/FTICP/StudyPopulation/Population/ChildBearingPotential";

			// Path to custom attribute in test article Teratogenic Potential
			string tpAttribute = "TeratogenicPotential";

			// Get the value for CBP
			string cbpType = icpInstMgr.getTypedDisplayValue(cbpPath, out isOther);
			string cbpOtherValue = null;
			if (isOther) 
			{
				cbpOtherValue = icpInstMgr.getTypedOtherDisplayValue(cbpPath, out isOther);
			}

			bool haveCBP = false;

			// We have one
			if (!LittleUtilities.isEmpty(cbpType)) 
			{
				haveCBP = true;
			}

			// Create advisory if missing
			if (!haveCBP) 
			{
				RuleAdvisory adv = new RuleAdvisory(
					_ruleId, MY_ID + "cbp", 
					"Pregnancy conditions not specified.  Specify: Population/Child Bearing Potential.");
				advisories.Add(adv);
			}

			// Check test articles for value
			CTMaterialEnumerator ctEnum = bom.getCTMaterialEnumerator();
			while (ctEnum.MoveNext()) 
			{
				ClinicalTrialMaterial ctm = ctEnum.getCurrent();

				string tpValue = (string )ctm.getValueForNode(tpAttribute);

				// Advisory if mising
				if (LittleUtilities.isEmpty(tpValue)) 
				{

					RuleAdvisory adv = new RuleAdvisory(
						_ruleId, MY_ID + ctm.getElementPath(), 
						"Pregnancy conditions not specified.  Specify: Test Article/Teratogenic Potential for test article <" +
						ctm.getMaterialName() + ">");
					advisories.Add(adv);
				}
			}
		}

	}

	public class TestRule2 : ITSPDRule
	{
#if false
<rule id="testcs2" type="CSHARP" displayName="Test CSharp 2" source="TspdCfg.FastTrack.Rules.TestRule2,FTRules.dll" categories="testcs" debug="false"/>
#endif

		public static readonly string MY_ID = "TestRule2";

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

			SOAEnumerator soaEnum = bom.getAllSchedules();
			while (soaEnum.MoveNext()) 
			{
				SOA soa = soaEnum.getCurrent();
				ProtocolEventEnumerator peEnum = soa.getAllVisits();

				while (peEnum.MoveNext()) 
				{
					ProtocolEvent pe = peEnum.getCurrent();

					TaskEnumerator taskEnum = soa.getTaskEnumerator();
					while (taskEnum.MoveNext()) 
					{
						Task task = taskEnum.getCurrent();

						TaskVisit tv = soa.getOrCreateTaskVisit(task.getObjID(), pe.getObjID(), false);
						if (tv == null) 
						{
							string smsg = "The visit <" + pe.getBriefDescription() + "> ";
							smsg += "in the schedule of activities <" + soa.getName() + "> ";
							smsg += "has been defined in the protocol, but no tasks have ";
							smsg += "been designated to be performed during the visit.";

							RuleAdvisory adv = new RuleAdvisory(_ruleId,MY_ID + pe.getObjID().ToString() + "x"+ task.getObjID().ToString(),smsg);
							
							advisories.Add(adv);
						}
					}
				}
			}

			return advisories;
		}
	}
}
