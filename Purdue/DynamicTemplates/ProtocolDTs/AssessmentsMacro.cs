//-------------------------------------------------------------------------------------------------
/// <remarks>
/// This program is the confidential and proprietary product of
/// Fast Track Systems, Inc. Any unauthorized use, reproduction,
/// or transfer of this program is strictly prohibited.
/// Copyright (C) 2009 by Medidata Worldwide solutions
/// All rights reserved.
/// $Workfile: C:\work_tspd\tsdcfg_200\SalesDemo\DynamicTemplates\FTProtocolDTs\AssessmentsMacro.cs$
/// $Revision: 5$
/// $Date: 11/15/2009 3:40:51 PM$
/// </remarks>
//-------------------------------------------------------------------------------------------------

using System;
using System.Collections;
using System.Diagnostics;
using System.Text;
using System.Runtime.InteropServices;

using Tspd.Tspddoc;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;
using Tspd.FormBase.ProgressBar;

using Word = Microsoft.Office.Interop.Word;

namespace VersionControl
{
	internal sealed class AssessmentsMacro {
		private static readonly string header_ = @"$Header: C:\work_tspd\DynamicTemplates\FTProtocolDTs\AssessmentsMacro.cs, 5, 11/19/2009 3:40:51 PM, Larry Peterson$";
	}
}

namespace TspdCfg.SalesDemo.DynTmplts {

	/// <summary>
	/// Summary description for AssessmentsMacro.
	/// </summary>
	public class AssessmentsMacro : MacroBase {

		private struct AssessmentInfo {
			private string type_;
			private string leaderTxt_;
			public AssessmentInfo(string type, string leaderTxt) {
				type_ = type;
				leaderTxt_ = leaderTxt;
			}
			public string type {
				get {return type_;}
			}
			public string leaderText {
				get {return leaderTxt_;}
			}
		}
		static private AssessmentInfo effAssInfo_ =
			new AssessmentInfo("Efficacy",
            "The following efficacy measures will be employed in this study:");
		static private AssessmentInfo safetyAssInfo_ =
			new AssessmentInfo("Safety",
			"The following measures of safety and tolerability will be employed in this study:");

		/// <summary>
		/// Efficacy assessments macro
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd EfficacyAssessments(
			MacroExecutor.MacroParameters mp) {
			try {
				mp.pba_.setOperation("Efficacy Assessments", "Generating efficacy assessments...");
				AssessmentsMacro macro = null;
				macro = new AssessmentsMacro(mp, effAssInfo_);
				macro.preprocess();
				macro.display();
				//macro.postprocess();
				return macro.macroStatusCode_;
			} catch (Exception e) {
				mp.inoutRng_.Text = "Failed in EfficacyAssessments: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		/// <summary>
		/// Safety assessments macro
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd SafetyAssessments(
			MacroExecutor.MacroParameters mp) {
			try {
				mp.pba_.setOperation("Safety Assessments", "Generating safety assessments...");
				AssessmentsMacro macro = null;
				macro = new AssessmentsMacro(mp, safetyAssInfo_);
				macro.preprocess();
				macro.display();
				//macro.postprocess();
				return macro.macroStatusCode_;
			} catch (Exception e) {
				mp.inoutRng_.Text = "Failed in SafetyAssessments: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		/// <summary>
		/// Wrapper around Task. Used for equality testing.
		/// </summary>
		private struct WTask {
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
		private AssessmentsMacro(MacroExecutor.MacroParameters mp, AssessmentInfo assInfo)
			: base(mp) {
			assInfo_ = assInfo;
			assTasks_ = new ArrayList();
		}

		/// <summary>
		/// 
		/// </summary>
		private void preprocess() {
			try {

				// Loop through all the SOAs.
				SOAEnumerator soaIter = bom_.getAllSchedules();
				while (soaIter.MoveNext()) {
					SOA soa = (SOA)soaIter.Current;
					ProtocolEventEnumerator visitIter = soa.getAllVisits();

					// Loop through tasks of the SOA.
					TaskEnumerator taskIter = soa.getTaskEnumerator();
					while (taskIter.MoveNext()) {
						Task task = (Task)taskIter.Current;
						WTask wTask = new WTask(task);

						// Skip the task if we already included it. To determine identity, see
						// WTask.Equals().
						if (assTasks_.Contains(wTask)) {
							continue;
						}

						// Loop through the visits to find a TaskVisit for the current task. If
						// there is one, then we're interested. Otherwise, this task is not
						// worthy.
						long taskID = task.getObjID();
						bool taskAdded = false;
						visitIter.Reset();
						while (!taskAdded && visitIter.MoveNext()) {
							ProtocolEvent visit = (ProtocolEvent)visitIter.Current;
							TaskVisit taskVisit =
								soa.getOrCreateTaskVisit(taskID, visit.getObjID(), false);
							if (taskVisit == null) {
								continue;
							}

							// We want the current task if one of the purpose types associated with
							// one of the task's visits is of the type we are looking for.
							IEnumerator tvpIter = soa.getTaskVisitPurposes(taskVisit);
							while (tvpIter.MoveNext()) {
								string tvpName =
									((TaskVisitPurpose)tvpIter.Current).getPurposeName();
								if (tvpName != null && 
									string.Compare(tvpName, assInfo_.type, true) == 0) 
								{
									assTasks_.Add(wTask);
									taskAdded = true;
									break;
								}
							}
						}
					}
				}
			} catch (Exception e) {
				Log.exception(e, "Problem in preprocess()");
				throw e;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		private void display() {
			try {
				Word.Range wrkRng = startAtBeginningOfParagraph();
				Word.Range begRng = wrkRng.Duplicate;

				wdDoc_.UndoClear();

				// The list template used to create specific lists. This is a single bullet list
				WordListHelper.ListTemplate wlt = WordListHelper.getBulletListTemplate(wdApp_);

				// Display assessment leader text.
				wrkRng.InsertAfter(assInfo_.leaderText);
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

				// If there are none, then there aren't any.
				if (assTasks_.Count < 1) {
					wrkRng.InsertAfter("	There are none defined.");
					goto EndOfMyCode;
				}

				// Display the list of assessment tasks.
				for (int i = 0; i < assTasks_.Count; ++i) {

					wlt.BeginListItem(ref wrkRng);

					Task task = ((WTask)assTasks_[i]).task;

					// Do elem-ref-take-on-style-hack
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
					wrkRng.InsertAfter("_");
					Word.Range hackRng = wrkRng.Duplicate;
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

					wrkRng.End = putElemRef(task, Task.DISPLAYNAME, wrkRng);

					hackRng.Delete(ref WordHelper.CHARACTER, ref Utils.O1);

					wlt.EndListItem(ref wrkRng);
				}

			EndOfMyCode:

				setOutgoingRng(begRng.Start, wrkRng.End);

			} catch (Exception e) {
				Log.exception(e, "Problem in display()");
				throw e;
			}
		}
	}
}
