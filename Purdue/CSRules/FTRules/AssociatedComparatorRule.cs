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
	/// <summary>
	/// Summary description for ComparatorRegimen.
	/// </summary>
	public class AssociatedComparatorRule : ITSPDRule
	{
#if false
<category id="testcs" labelKey="Test C#" /> 
<rule id="assoccomparesharp2" type="CSHARP" source="TspdCfg.FastTrack.Rules.AssociatedComparatorRule,FTRules.dll" categories="testcs" debug="false"/>
#endif		
		public static readonly string MY_ID = "AssociatedComparatorRule";
		private static string _ruleId = "";
		private static bool _debug = false;

			
		
		//
		// TODO: Add constructor logic here
		//

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

			runComparator(advisories);

			return advisories;
		}

		private void runComparator(ArrayList advisories) 
		{

			ContextManager ctx = ContextManager.getInstance();
			TspdDocument doc = ctx.getActiveDocument();
			BusinessObjectMgr bom = doc.getBom();

			IcpInstanceManager icpInstMgr = doc.getBaseInstanceMgr();
			ClinicalTrialMaterial placebo = null;
			ClinicalTrialMaterial primary = null;

			string primaryRole = "investigationalproduct";
			string placeboRole = "comparator";
			// Check test articles for value
			CTMaterialEnumerator ctEnum = bom.getCTMaterialEnumerator();

			//first find the itme
			while (ctEnum.MoveNext()) 
			{
				ClinicalTrialMaterial ctm = ctEnum.getCurrent();

				if(ctm.getParentLikeChild() == null)
				{
					string role = ctm.getPrimaryRole();								
					if (!LittleUtilities.isEmpty(role) && role.ToLower().CompareTo(primaryRole) == 0)
					{
						primary = ctm;
						break;
					}
				}
			}
			if(primary == null)
			{
				//missing something
				//could give an advisory too
				return;
			}
			
			//now find the associated placebo
			CTMaterialEnumerator assocEnum = bom.getAssociatedTrialMaterials(primary);
			while (assocEnum.MoveNext()) 
			{
				ClinicalTrialMaterial ctm = assocEnum.getCurrent();

				string role = ctm.getPrimaryRole();
				if (!LittleUtilities.isEmpty(role) && role.ToLower().CompareTo(placeboRole) == 0)
				{	
					placebo = ctm;
						break;
				}
			}

			if(placebo == null)
			{
				//missing something
				//could give an advisory too
				return;
			}

			string placeboFormula = placebo.getFormulation().ToLower();
			string primaryFormula = primary.getFormulation().ToLower();
			if(primaryFormula.CompareTo(placeboFormula) != 0)
			{
				RuleAdvisory adv = new RuleAdvisory(
					_ruleId, MY_ID + primary.getElementPath(), 
					"<" + primary.getMaterialName() + "> has a formulation of <" + primaryFormula + 
					"> but the <" + placebo.getMaterialName() + "> for this study has a formulation of <" +
					placeboFormula + ">. The formulations should be the same.");
				advisories.Add(adv);
			}

			string childFormula = "";
			ctEnum.Reset();
			while(ctEnum.MoveNext()) 
			{
				//could inforce association here too
				ClinicalTrialMaterial child = ctEnum.getCurrent();
				if(primary.getObjID().Equals(child.getParentID()))
				{
					childFormula = child.getFormulation().ToLower();
					if(childFormula.CompareTo(placeboFormula) != 0)
					{
						RuleAdvisory adv = new RuleAdvisory(
							_ruleId, MY_ID + child.getElementPath(), 
							"The formulation of the children of <" + primary.getMaterialName() + 
							"> do not match <" + placeboFormula + "> for this study." +
							" The formulations should be the same.");
						advisories.Add(adv);
						break;
					}
				}
			}
		}
	}
}
