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
using Tspd.Bridge;

namespace TspdCfg.FastTrack.Rules
{
	public class ProcedureSetRule : ITSPDRule
	{
#if false
<rule id="Template" type="CSHARP" displayName="Procedure Set" source="TspdCfg.FastTrack.Rules.ProcedureSetRule,Rules.dll" categories="Document" debug="false"/>
#endif

        public static readonly string MY_ID = "ProcedureSet";
        public ArrayList finalResult = new ArrayList();
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

			string smsg = "";
            advisories = getProcSets(bom);
			RuleAdvisory adv = new RuleAdvisory(
				_ruleId, MY_ID, // add element path to my_id if based on a specific object
				smsg);
			advisories.Add(adv);
            
			return advisories;
		}

        private ArrayList getProcSets(BusinessObjectMgr bom_)
        {

            getglobalTaskSets(bom_);
            IList soaList = bom_.getAllSchedules().getList();  //Get list of all Schedules.
            foreach (SOA _soa in soaList)
            {
                Task tv = null;
             //   tv.get

            }

            return finalResult;
        }

        private void getglobalTaskSets(BusinessObjectMgr bom_)
        {
            

            TaskSetCV tcv =  BridgeProxy.getInstance().getTaskSetList();
            IEnumerator taskEnum = tcv.iterator();
            while (taskEnum.MoveNext())
            {
                TaskSet taskSet = (TaskSet)taskEnum.Current;
                System.Windows.Forms.MessageBox.Show(taskSet.getName());
                
            }

        }
	}
}
