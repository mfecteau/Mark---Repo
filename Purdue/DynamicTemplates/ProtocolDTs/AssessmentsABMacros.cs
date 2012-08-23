using System;
using System.Collections;
using System.Windows.Forms;
using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;

using TspdCfg.FastTrack.DynTmplts;
using Word = Microsoft.Office.Interop.Word;

namespace VersionControl 
{
	internal sealed class AssessmentsABMacros 
	{
		private static readonly string header_ = @"$Header: C:\work_tspd\tsdcfg_200\Common\DynamicTemplates\FTProtocolDTs\AbstractBased\AssessmentsABMacros.cs, 5, 9/15/2008 3:40:50 PM, Larry Peterson$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	public class AssessmentsABMacros : AbstractMacro
	{
        #region Dynamic Template Methods
		
		public class AssessmentInfo 
		{
			private string type_;
			private string leaderTxt_;
			private string endTxt_;
			public AssessmentInfo(string type) 
			{
				type_ = type;
				leaderTxt_ = "The following will be measured for ";
				endTxt_ = " evaluation(s):";
			}
			public AssessmentInfo(string type, string leaderTxt, string endTxt) 
			{
				type_ = type;
				leaderTxt_ = leaderTxt;
				endTxt_ = endTxt;
			}
			public string type 
			{
				get {return type_;}
			}
			/*public string bodyText 
			{
				get {return leaderTxt_ + type.ToLower() + endTxt_;}
			}*/

			public string getBodyText(IcpSchemaManager icpSchemaMgr)
			{
				// fix bug 46Q: use user label rather than system label in output
				return leaderTxt_ + icpSchemaMgr.getUserLabel("PurposeTypes", type) + endTxt_;
			}
		}

		/// <summary>
		/// Displays a single organization based upon given context.
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
        /// 
        public static MacroExecutor.MacroRetCd Assessment(
      MacroExecutor.MacroParameters mp)
        {
            try
            {
                mp.pba_.setOperation("Efficacy Assessments", "Generating information...");

                
                AssessmentsABMacros macro = null;
                macro = new AssessmentsABMacros(mp, new AssessmentInfo("Efficacy"));
                macro.preProcess();
                macro.display();
                //macro.postprocess();
                return macro.macroStatusCode_;
            }
            catch (Exception e)
            {
                mp.inoutRng_.Text = "Failed in Efficacy Assessments: " + e.Message;
            }
            return MacroExecutor.MacroRetCd.Failed;
        }

     
		public static MacroExecutor.MacroRetCd OtherAssessments(
			MacroExecutor.MacroParameters mp) 
		{
			try 
			{
				mp.pba_.setOperation("Other Assessments", "Generating information...");
				
				// needs to pick up the otherText, sort by and display other Text instead of type if otherText exists 
				// may want to derive from AssessmentInfo
				AssessmentsABMacros macro = null;
				macro = new AssessmentsABMacros(mp, new AssessmentInfo("Other"));
				macro.preProcess();
				macro.displayOther();
				//macro.postprocess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				mp.inoutRng_.Text = "Failed in other Assessments: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}
		

		/// <summary>
		/// Displays a single organization based upon given context.
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		protected static MacroExecutor.MacroRetCd OtherAssessments(string type,
			MacroExecutor.MacroParameters mp) 
		{
			try 
			{
				AssessmentsABMacros macro = null;
				macro = new AssessmentsABMacros(mp, new AssessmentInfo(type));	
				string typeUserLabel = macro.IcpSchemaMgr.getUserLabel("PurposeTypes", type);
				if (typeUserLabel != null && typeUserLabel.Length > 0)
					typeUserLabel = typeUserLabel.Substring(0, 1).ToUpper() + typeUserLabel.Substring(1);
				mp.pba_.setOperation(typeUserLabel +"  Assessments", "Generating information...");
				
				macro.preProcess();
				macro.display();
				//macro.postprocess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				mp.inoutRng_.Text = "Failed in " + type + " Assessments: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		
		#endregion

		#region Tasks evaluation region
		/// <summary>
		/// Wrapper around Task. Used for equality testing.
		/// </summary>
		private struct WTask 
		{
			private Task task_;
			public WTask(Task task) {
				task_ = task;
			}
			public Task task {
				get { return task_; }
			}
			// Ok, when are two tasks equal? It depends. In our case, they are equal if the task
			// name is the same. I know, the user can change the task name so that two technically
			// different tasks appear to be the same, and the user can change the task name
			// making two identical tasks appear to be different. But since "equal" is in the
			// eye of the user and it's the user that will be seeing the list, taskName will be
			// the object of comparison.
			public override bool Equals(object obj) {
				if (!(obj is WTask)) {
					throw new ArgumentException("Array element is not of type 'WTask'");
				}
				string thisDN = task_.getDisplayName();
				string thatDN = ((WTask)obj).task_.getDisplayName();
				if (thisDN == null && thatDN == null) {
					return true;
				} else if (thisDN == null || thatDN == null) {
					return false;
				}
				return thisDN.Equals(thatDN);
			}
			public override int GetHashCode() {
				string taskName = task_.getDisplayName();
				if (taskName == null) {
					return 0;
				}
				return taskName.GetHashCode();
			}

		}
	
		#endregion

		// Macro variables.
		AssessmentInfo assInfo_;
		ArrayList assTasks_; // Element: WTask

		/// <summary>
		/// 
		/// </summary>
		/// <param name="tspdDoc"></param>
		/// <param name="inoutRng"></param>
		/// <param name="insertNew"></param>
		/// <param name="pba"></param>
		public AssessmentsABMacros(MacroExecutor.MacroParameters mp, AssessmentInfo assInfo)
			: base(mp) 
		{
			assInfo_ = assInfo;
			assTasks_ = new ArrayList();
		}

		/// <summary>
		/// 
		/// </summary>
		public override void preProcess() 
		{
			try 
			{
                AssessmentSelection frmSelection = new AssessmentSelection();
                frmSelection.label1.Text = "Select Assessments";
                frmSelection.Fill_Assessments(bom_);

               // frmSelection.cmbOutcome.Items.Add("");
                if (frmSelection.ShowDialog() == DialogResult.OK)
                {
                   // EnumPair ep = 
 
                }

				// Loop through all the SOAs.
				SOAEnumerator soaIter = bom_.getAllSchedules();
				while (soaIter.MoveNext()) 
				{
					SOA soa = (SOA)soaIter.Current;
					ProtocolEventEnumerator visitIter = soa.getAllVisits();

					// Loop through tasks of the SOA.
					TaskEnumerator taskIter = soa.getTaskEnumerator();
					while (taskIter.MoveNext()) 
					{
						Task task = (Task)taskIter.Current;
						WTask wTask = new WTask(task);

						// Skip the task if we already included it. To determine identity, see
						// WTask.Equals().
						if (assTasks_.Contains(wTask)) 
						{
							continue;
						}

						// Loop through the visits to find a TaskVisit for the current task. If
						// there is one, then we're interested. Otherwise, this task is not
						// worthy.
						long taskID = task.getObjID();
						bool taskAdded = false;
						visitIter.Reset();
						while (!taskAdded && visitIter.MoveNext()) 
						{
							ProtocolEvent visit = (ProtocolEvent)visitIter.Current;
							TaskVisit taskVisit =
								soa.getOrCreateTaskVisit(taskID, visit.getObjID(), false);
							if (taskVisit == null) 
							{
								continue;
							}

							// We want the current task if one of the purpose types associated with
							// one of the task's visits is of the type we are looking for.
							IEnumerator tvpIter = soa.getTaskVisitPurposes(taskVisit);
							while (tvpIter.MoveNext()) 
							{
								string tvpName = ((TaskVisitPurpose)tvpIter.Current).getPurposeName();

								if (tvpName != null && 
									string.Compare(assInfo_.type, tvpName, true) == 0) 
								{
									assTasks_.Add(wTask);
									taskAdded = true;
									break;
								}
							}
						}
					}
				}
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Problem in preprocess()");
				throw e;
			}
		}
		
		public override void postProcess(){}

        public  void sst()
        { 
        }
		public new static bool canRun(BaseProtocolObject bpo)
		{
			return true;
		}

		/// <summary>
		/// 
		/// </summary>
		public override void display() 
		{
			try 
			{
				Word.Range wrkRng = startAtBeginningOfParagraph();
				Word.Range begRng = wrkRng.Duplicate;

				wdDoc_.UndoClear();
				//wrkRng.End = MacroBaseUtilities.insertHiddenNBSpace(wrkRng);

				

				// Display assessment leader text.
				wrkRng.InsertAfter(assInfo_.getBodyText(icpSchemaMgr_));
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

				// If there are none, then there aren't any.
				if (assTasks_.Count < 1) 
				{
					wrkRng.InsertAfter("	There are none defined.");
					goto EndOfMyCode;
				}

				// Start the bulleted list.
				WordListHelper.ListTemplate wlt = WordListHelper.getBulletListTemplate(wdApp_);
                object O1 = 1;
				// Display the list of assessment tasks.
				for (int i = 0; i < assTasks_.Count; ++i) 
				{
					Task task = ((WTask)assTasks_[i]).task;

					wlt.BeginListItem(ref wrkRng);

					// Do elem-ref-take-on-style-hack
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
					wrkRng.InsertAfter("_");
					Word.Range hackRng = wrkRng.Duplicate;
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, task, Task.DISPLAYNAME, wrkRng);

                  //  hackRng.Delete(
                    hackRng.Delete(ref WordHelper.CHARACTER, ref O1);

					wlt.EndListItem(ref wrkRng);
				}


			EndOfMyCode:

				setOutgoingRng(begRng.Start, wrkRng.End);

			} 
			catch (Exception e) 
			{
				Log.exception(e, "Problem in display()");
				throw e;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		public void displayOther() 
		{
			try 
			{
				Word.Range wrkRng = startAtBeginningOfParagraph();
				Word.Range begRng = wrkRng.Duplicate;

				wdDoc_.UndoClear();
				//wrkRng.End = MacroBaseUtilities.insertHiddenNBSpace(wrkRng);

				// If there are none, then there aren't any.
				if (assTasks_.Count < 1) 
				{
					wrkRng.InsertAfter(assInfo_.getBodyText(icpSchemaMgr_));
					wrkRng.InsertParagraphAfter();
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
					wrkRng.InsertAfter("	There are none defined.");
					goto EndOfMyCode;
				}

				// Display assessment leader text.
				// sort by OtherText and ......
				wrkRng.InsertAfter(assInfo_.getBodyText(icpSchemaMgr_)); // needs to pick up the OtherText......
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

				// Start the bulleted list.
				WordListHelper.ListTemplate wlt = WordListHelper.getBulletListTemplate(wdApp_);

                object O1 = 1;

				// Display the list of assessment tasks.
				for (int i = 0; i < assTasks_.Count; ++i) 
				{
					Task task = ((WTask)assTasks_[i]).task;

					wlt.BeginListItem(ref wrkRng);

					// Do elem-ref-take-on-style-hack
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
					wrkRng.InsertAfter("_");
					Word.Range hackRng = wrkRng.Duplicate;
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, task, Task.DISPLAYNAME, wrkRng);

					hackRng.Delete(ref WordHelper.CHARACTER, ref O1);

					wlt.EndListItem(ref wrkRng);
				}

			EndOfMyCode:

				setOutgoingRng(begRng.Start, wrkRng.End);

			} 
			catch (Exception e) 
			{
				Log.exception(e, "Problem in display()");
				throw e;
			}
		}
	}
}


