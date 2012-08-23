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
	public class TestStatistics : ITSPDRule
	{
#if false
<rule id="TestStatistics" type="CSHARP" displayName="Test Statistics" source="TspdCfg.FastTrack.Rules.TestStatistics,FTRules.dll" categories="stats" debug="false">
Verify that for each Test Statistic assigned to a Hypothesis, that it is an output of at least one of the Analyses assigned to the Hypothesis.
</rule>
#endif

		public static readonly string MY_ID = "TestStatistics";

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

			VariableDictionary vd = bom.getVariableDictionary();

			ModelMappingEnumerator mmEnum = bom.getStatsMgr().getAllModelMappings();

			HypothesisEnumerator hypothesisEnum = bom.getStatsMgr().getHypothesisEnumerator();
			while (hypothesisEnum.MoveNext())
			{
				Hypothesis hypothesis = hypothesisEnum.getCurrent();

				MappedVariableCollection hypMvc = hypothesis.getTestStatisticsCollection();
				IEnumerator hypVarRefEnum = hypMvc.getVarRefs();
				while (hypVarRefEnum.MoveNext()) 
				{
					bool ruleSatisfied = false;
					VarRef hypVarRef = hypVarRefEnum.Current as VarRef;

					mmEnum.Reset();
					while (mmEnum.MoveNext() && !ruleSatisfied) 
					{
						AnalysisToModel a2m = mmEnum.getCurrent();
						MappedVariableCollection a2mMvc = a2m.getVariableCollection();

						IEnumerator a2mVarRefEnum = a2mMvc.getVarRefs();
						while (a2mVarRefEnum.MoveNext() && !ruleSatisfied) 
						{
							VarRef a2mVarRef = a2mVarRefEnum.Current as VarRef;

							if (hypVarRef.getVariableID() == a2mVarRef.getVariableID() &&
								a2mVarRef.getIORole() == VarRef.InputOutputRole.Output) 
							{
								// Rule is satisfied if there is at least one xvsa fdsf asdfsda fsda sdfa sdfsdamfkl
								ruleSatisfied = true;
							}
						}
					}

					if (!ruleSatisfied) 
					{
						StudyVariable sv = vd.findBySourceID(hypVarRef.getVariableID());

						string smsg = "Hypothesis: <" + hypothesis.getActualDisplayValue();
						smsg += "> refers to Test Statistic <" + sv.getActualDisplayValue();
						smsg += "> that is not generated by any analysis.";

						RuleAdvisory adv = new RuleAdvisory(
							_ruleId, 
							MY_ID + hypothesis.getElementPath() + sv.getElementPath(), // add element path to my_id if based on a specific object
							smsg);
						advisories.Add(adv);
					}
				}
			}

			return advisories;
		}

	}
}
