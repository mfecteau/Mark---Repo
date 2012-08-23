using System;
using System.Collections;
using System.Windows.Forms;

using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;
using MSXML2;

using Word = Microsoft.Office.Interop.Word;

namespace VersionControl 
{
	internal sealed class LabAssessmentsMacro
	{
		private static readonly string header_ = @"$Header: LabAssessmentsMacro.cs, 1, 18-Aug-09 12:04:40, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for LabAssessmentsMacro.
	/// </summary>
	public class LabAssessmentsMacro : AbstractMacroImpl
	{
		class TaskVariables 
		{
			public Task task = null;
			public ArrayList pvList = new ArrayList();
			public ArrayList variables = new ArrayList();
		}

		public class TVByTaskComparer : IComparer  
		{
			public TVByTaskComparer(){}

			int IComparer.Compare(Object x, Object y)  
			{
				TaskVariables tv1 = x as TaskVariables;
				TaskVariables tv2 = y as TaskVariables;
				
				return tv1.task.getSequence().CompareTo(tv2.task.getSequence());
			}
		}

		ArrayList _taskVariableList = new ArrayList();
		ArrayList _badEvents = new ArrayList();
		SOA _currentSOA = null;

		bool _showVariableAbbreviation = false;
		bool _includeScheduledTimes = false;

		public string _headingStyle = null;
		public string _listStyle = null;

		public LabAssessmentsMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public static new bool canRun(BaseProtocolObject bpo)
		{
			SOA soa = bpo as SOA;
			if (soa == null)
			{
				return false;
			}

			if (soa.isSchemaDesignMode()) 
			{
				return false;
			}

			return true;
		}
		
		#region Dynamic Template Methods
		
		#region LabAssessments

		public static MacroExecutor.MacroRetCd LabAssessments (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.LabAssessmentsMacro.LabAssessments,ProtocolDTs.dll" elementLabel="Laboratory Assessments" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="Laboratory Assessments." shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
#endif
			try 
			{
				mp.pba_.setOperation("Lab Assessments Macro", "Generating information...");
				
				LabAssessmentsMacro macro = null;
				macro = new LabAssessmentsMacro(mp);

				macro._headingStyle = PfizerUtil.PFIZER_STYLE_TEXT_TI12;
				macro._listStyle = PfizerUtil.PFIZER_STYLE_TEXT_BULL;

				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Lab Assessments Macro");
				mp.inoutRng_.Text = "Lab Assessments Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		#endregion

		#region LabAssessments

		public static MacroExecutor.MacroRetCd LabAssessmentsSynopsis (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.LabAssessmentsMacro.LabAssessmentsSynopsis,ProtocolDTs.dll" elementLabel="Laboratory Assessments Synopsis" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="Laboratory Assessments for Synopsis." shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
#endif
			try 
			{
				mp.pba_.setOperation("Lab Assessments Synopsis Macro", "Generating information...");
				
				LabAssessmentsMacro macro = null;
				macro = new LabAssessmentsMacro(mp);

				macro._headingStyle = PfizerUtil.PFIZER_STYLE_TABLETEXT_10;
				macro._listStyle = PfizerUtil.PFIZER_STYLE_TABLETEXT_BULL_10;

				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Lab Assessments Synopsis Macro");
				mp.inoutRng_.Text = "Lab Assessments Synopsis Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		#endregion

		#endregion

		public override void preProcess()
		{
			pba_.updateProgress(1.0);

			// Hashtable of unique tasks
			Hashtable htTasks = new Hashtable();

			string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);
			if (MacroBaseUtilities.isEmpty(elementPath)) 
			{
				return;
			}

			SOAEnumerator soaEnum = bom_.getAllSchedules();
				
			while (soaEnum.MoveNext())
			{
				pba_.updateProgress(2.0);

				SOA soa = soaEnum.getCurrent();
				if (soa.getElementPath().Equals(elementPath)) 
				{
					_currentSOA = soa;
					break;
				}
			}

			if (_currentSOA == null) return;

			// Get stored parameters
			string sParms = execParms_.getParm(MacroExecutor.MacroExecParameters.PARM_1);
			string[] aParms = null;

			if (!MacroBaseUtilities.isEmpty(sParms)) 
			{
				aParms = sParms.Split('|');
			}

			bool parmsValid = false;

			if (aParms != null && aParms.Length == 2)
			{
				parmsValid = true;

				if (!MacroBaseUtilities.isEmpty(aParms[0])) 
				{
					try 
					{
						_showVariableAbbreviation = bool.Parse(aParms[0]);
					}
					catch (Exception ex) 
					{
						parmsValid = false;
					}
				}

				if (!MacroBaseUtilities.isEmpty(aParms[1])) 
				{
					try 
					{
						_includeScheduledTimes = bool.Parse(aParms[1]);
					}
					catch (Exception ex) 
					{
						parmsValid = false;
					}
				}
			}

			// Ask the user if the parms are missing/invalid
			if (!parmsValid) 
			{
				LabSelections labSelect = new LabSelections();

				DialogResult res = labSelect.ShowDialog();
				if (res == DialogResult.OK) 
				{
					_showVariableAbbreviation = labSelect.ShowVariableAbbreviation;
					_includeScheduledTimes = labSelect.IncludeScheduledTimes;

					// save it for next time so we don't ask
					sParms = _showVariableAbbreviation.ToString() + "|";
					sParms += _includeScheduledTimes.ToString();

					execParms_.setParm(MacroExecutor.MacroExecParameters.PARM_1, sParms);
				}
			}

			
			// Now collect task visits
			TaskVisitEnumerator tvEnum = _currentSOA.getAllTaskVisits();

			while (tvEnum.MoveNext()) 
			{
				TaskVisit tv = tvEnum.getCurrent();
				if (tv.isCentralFacility() || tv.isLocalFacility())
				{
					pba_.updateProgress(1.0);
					
					findVariables(_currentSOA, tv, htTasks);
				}
			}
		}

		private void findVariables(SOA soa, TaskVisit tv, Hashtable htTasks) 
		{

			// For all of the found central/local lab tasks,
            // find all of the variables associated with them.
			//
			// Collect them my the unique task key so that the task
			// not repeated.
			VariableDictionary dict = bom_.getVariableDictionary();

			Task task = soa.getTaskOfTaskVisit(tv);
			ProtocolEvent visit = soa.getVisitOfTaskVisit(tv);

			string studyDay = visit.getStudyDayTime();
			if (MacroBaseUtilities.isEmpty(studyDay)) 
			{
				if (!_badEvents.Contains(visit)) 
				{
					_badEvents.Add(visit);
				}
			}

			long taskKey;

			if (task.isLocalTask()) 
				taskKey = task.getObjID();
			else
				taskKey = task.getProcDefId();
			
			// Look it up
			TaskVariables tVar = htTasks[taskKey] as TaskVariables;
			if (tVar == null) 
			{
				// Not found, create it
				tVar = new TaskVariables();
				tVar.task = task;
				htTasks[taskKey] = tVar;

				_taskVariableList.Add(tVar);
			}

			Period per = bom_.getParentOfScheduleItem(visit);
			PfizerUtil.PeriodAndVisit pv = new PfizerUtil.PeriodAndVisit();
			pv.per = per;
			pv.visit = visit;
			tVar.pvList.Add(pv);

			// Get the mappings to get the variables
			IEnumerator mapEnum = task.getVariableMappings();
			while (mapEnum.MoveNext()) 
			{
				pba_.updateProgress(2.0);

				VarRef varref = mapEnum.Current as VarRef;

				// unchecked, don't include
				if (varref.isHidden()) 
				{
					continue;
				}

                StudyVariable var = dict.findBySourceID(varref.getVariableID());

				// Save off if we don't already have it
				if (var != null && !tVar.variables.Contains(var)) 
				{
					tVar.variables.Add(var);
				}
			}
		}

		public override void display()
		{
			Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;

			if (_currentSOA == null)
			{
				pba_.updateProgress(70.0);

				wrkRng.InsertAfter("This schedule that this macro refers to was removed, delete this macro.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
						
				inoutRange.End = wrkRng.End;
				setOutgoingRng(inoutRange);
				wdDoc_.UndoClear();
				return;
			}

			displayLabAssessments(wrkRng);

			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}

		public override void postProcess()
		{
			// Clean up memory
			_taskVariableList.Clear();
			_badEvents.Clear();
			_currentSOA = null;
		}

		private TaskVisit findTaskVisit(SOA soa, long taskID, long visitID) 
		{
			TaskVisit tv = soa.getOrCreateTaskVisit(taskID, visitID, false);
			if (tv != null) 
			{
				tv.getComplexChildren();
				return tv;
			}

			return null;
		}

		private void displayLabAssessments(Word.Range wrkRng) 
		{
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			object oStyle = tspdDoc_.getStyleHelper().setNamedStyle(_headingStyle, wrkRng);
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			wrkRng.InsertAfter("Laboratory tests");

			Word.Range underlineRange = wrkRng.Duplicate;
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			wrkRng.InsertAfter(" ");

			// Underline the text
			underlineRange.Font.Underline = Word.WdUnderline.wdUnderlineSingle;

			wrkRng.InsertParagraphAfter();
			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				
			if (_badEvents.Count != 0) 
			{
				wrkRng.InsertAfter("Study Events must have a Study Day entered.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
			else if (_taskVariableList.Count == 0) 
			{
				wrkRng.InsertAfter("No Task-Events defined with central or local lab selected.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
			else
			{
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

				oStyle = tspdDoc_.getStyleHelper().setNamedStyle(_listStyle, wrkRng);
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

				// Sort by Task sequence
				_taskVariableList.Sort(new TVByTaskComparer());
				foreach (TaskVariables tVar in _taskVariableList) 
				{
					pba_.updateProgress(2.0);

					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, tVar.task, Task.BRIEF_DESCRIPTION, wrkRng, macroEntry_);

					if (!MacroBaseUtilities.isEmpty(tVar.task.getFullDescription()))
					{
						wrkRng.End = MacroBaseUtilities.putAfterElemRef(":", tspdDoc_, wrkRng);

						wrkRng.InsertAfter(tVar.task.getFullDescription());
						wrkRng.InsertAfter(" ");
					}
					else
					{
						wrkRng.End = MacroBaseUtilities.putAfterElemRef(":", tspdDoc_, wrkRng);
					}

					if (_includeScheduledTimes) 
					{
						wrkRng.InsertAfter("to be collected at ");


						// Sort visits by sequence
						tVar.pvList.Sort(new PfizerUtil.PeriodAndVisitComparer());

						Period lastPer = null;
						foreach (PfizerUtil.PeriodAndVisit pv in tVar.pvList) 
						{
							Period curPer = pv.per;
							if (lastPer == null || curPer.getObjID() != lastPer.getObjID()) 
							{
								if (lastPer != null) 
								{
									wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);
								}

								wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, curPer, Period.BRIEF_DESCRIPTION, wrkRng, macroEntry_);
								wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);
							}
							else if (lastPer != null) 
							{
								wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);
							}

							wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, pv.visit, ProtocolEvent.STUDY_DAYTIME, wrkRng, macroEntry_);

							TaskVisit tv = findTaskVisit(_currentSOA, tVar.task.getObjID(), pv.visit.getObjID());
							if (tv != null) 
							{
								if (!MacroBaseUtilities.isEmpty(tv.getFullDescription()))
								{
									wrkRng.InsertAfter(tv.getFullDescription());
									wrkRng.InsertAfter(" ");
								}
							}

							lastPer = curPer;
						}
					
						if (tVar.variables.Count != 0) 
						{
							wrkRng.End = MacroBaseUtilities.putAfterElemRef(":", tspdDoc_, wrkRng);
						}
					}

					bool haveFirstMapping = false;
					foreach (StudyVariable var in tVar.variables) 
					{
						pba_.updateProgress(2.0);

						if (haveFirstMapping) 
						{
							wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);
						}
						
						haveFirstMapping = true;

						wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, var, StudyVariable.BRIEF_DESCRIPTION, wrkRng, macroEntry_);

						if (_showVariableAbbreviation) 
						{
							wrkRng.InsertAfter("(");
							wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, var, StudyVariable.LAB_DESCRIPTION, wrkRng, macroEntry_);
							wrkRng.End = MacroBaseUtilities.putAfterElemRef(")", tspdDoc_, wrkRng);
						}
					}

					if (!haveFirstMapping) 
					{
						// wrkRng.InsertAfter("No Mappings.");
					}
					
					wrkRng.InsertParagraphAfter();
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
					wdDoc_.UndoClear();
				}

				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				
				oStyle = tspdDoc_.getStyleHelper().setNamedStyle(PfizerUtil.NORMAL, wrkRng);
			}

			wdDoc_.UndoClear();
		}
	}
}
