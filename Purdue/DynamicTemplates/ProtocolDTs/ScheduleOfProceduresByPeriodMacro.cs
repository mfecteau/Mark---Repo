using System;
using System.Collections;

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
	internal sealed class ScheduleOfProceduresByPeriodMacro
	{
		private static readonly string header_ = @"$Header: ScheduleOfProceduresByPeriodMacro.cs, 1, 18-Aug-09 12:05:44, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for ScheduleOfProceduresByPeriodMacro.
	/// [PERIOD]
	///                     [Visit1]                      [Visit2]                       [Visit3]
	///               [predose]        [postdose]        [predose]      [postdose]
	/// </summary>
	public class ScheduleOfProceduresByPeriodMacro : AbstractMacroImpl
	{
		public static string TREATMENT_TYPE = "treatment"; // from scheduleItemTypes
		public static readonly string DOSING_TREATMENT = "dosingtreatment"; // from scheduleItemTypes

		public class TVCell
		{
			public ArrayList tvList = new ArrayList();
			public string spanLabel = "";
		}

		// wrapper class for a TreatmentVisit in a Treatment period with all the gory details
		public class TreatmentVisit
		{
			public ProtocolEvent	_visit;
			public ArrayList	_tvCells; // array of tcWrappers
			public ArrayList	_tvSpanCells; // array of tcWrappers
			public Hashtable	_taskVisitsByTaskID;  // only topLevel 

			public TreatmentVisit(ProtocolEvent visit, SOA soa)
			{
				_visit = visit;
				_tvCells = new ArrayList();  
				_tvSpanCells = new ArrayList();
				_taskVisitsByTaskID = new Hashtable();
				IList ls = soa.getTaskVisitsForVisit(visit).getList();
				foreach (TaskVisit v in ls)
				{
					int copies = soa.getCopiesOfTaskVisit(v).getList().Count;
					_taskVisitsByTaskID[v.getAssociatedTaskID()] = copies;
				}
			}

			/// <summary>
			/// Deep Equals
			/// Are the top level task visits the same?
			/// Are the counts of the task visit copies the same?
			/// Are the time columns equal
			/// Are the column spans the same.
			/// One case is not checked: mirror image patterns.
			///		Visit 1    task A at 30 min, task A and B 40 min
			///		Visit 2    task A and B 30 min task A at 40 min
			/// </summary>
			/// <param name="other"></param>
			/// <returns></returns>
			public override bool Equals(object other)
			{
				// -- do we need some kind of test around the visit name?

				IComparer comparer = new PurdueUtil.TCWrapperComparer();
				TreatmentVisit tvo = other as TreatmentVisit;
				// -- do we have the same tasks in the overall sense?
				if (tvo._taskVisitsByTaskID.Count != _taskVisitsByTaskID.Count)
					return false;

				Hashtable hto = tvo._taskVisitsByTaskID;

				foreach(long tid in _taskVisitsByTaskID.Keys)
				{
					if (!hto.Contains(tid))
					{
						return false;
					}
					int copiesOfThis = (int)(_taskVisitsByTaskID[tid]);
					int copiesOfOther = (int)(hto[tid]);

					if (copiesOfThis != copiesOfOther)
					{
						return false;
					}
				}
				// We have same task visits and the same number of each.
				// -- do we have the same column layout
				if (tvo._tvCells.Count == _tvCells.Count)
				{
					for(int i = 0; i < _tvCells.Count; i++)
					{
						PurdueUtil.TCWrapper tc1 = _tvCells[i] as PurdueUtil.TCWrapper;
						PurdueUtil.TCWrapper tc2 = tvo._tvCells[i] as PurdueUtil.TCWrapper;
						if (comparer.Compare(tc1, tc2) != 0 || tc1.Label.Equals(tc2.Label) == false)
						{
							return false;
						}
					}
				}
				else
				{
					return false;
				}
				if (tvo._tvSpanCells.Count == _tvSpanCells.Count)
				{
					for(int i = 0; i < _tvSpanCells.Count; i++)
					{
						PurdueUtil.TCWrapper tc1 = _tvSpanCells[i] as PurdueUtil.TCWrapper;
						PurdueUtil.TCWrapper tc2 = tvo._tvSpanCells[i] as PurdueUtil.TCWrapper;
						if (comparer.Compare(tc1, tc2) != 0 || tc1.Label.Equals(tc2.Label) == false)
						{
							return false;
						}
					}
				}
				else
				{
					return false;
				}
		
				return true;
			}
		}

		// wrapper class for a TreatmentPeriod within an SOA with all its gory details
		public class TreatmentPeriod
		{
			public Period _period;
			public ArrayList _treatmentVisits; // list of TreatmentVisits 
			public ArrayList _taskList;  // the all tasks with TVlabels in this period
			public Task _anchorTask;
			public TreatmentPeriod(Period p)
			{
				_period = p;
				_treatmentVisits = new ArrayList();
				_taskList = new ArrayList();
			}

			public override int GetHashCode()
			{
				return _period.getObjectRoot().GetHashCode();
			}


			/// <summary>
			/// Deeper Equals
			/// </summary>
			/// <param name="other"></param>
			/// <returns></returns>
			public override bool Equals(object other)
			{
				IComparer comp = new PurdueUtil.TCWrapperComparer();
				TreatmentPeriod tpo = other as TreatmentPeriod;
				if (_treatmentVisits.Count == tpo._treatmentVisits.Count)
				{
					for(int i = 0; i < _treatmentVisits.Count; i++)
					{
						TreatmentVisit v1 = _treatmentVisits[i] as TreatmentVisit;
						v1._tvCells.Sort(comp);
						TreatmentVisit v2 = tpo._treatmentVisits[i] as TreatmentVisit;
						v2._tvCells.Sort(comp);
						if (!v1.Equals(v2))
							return false;
					}
				}
				else
				{
					return false;
				}

				if (_taskList.Count == tpo._taskList.Count)
				{
					for(int i = 0; i < _taskList.Count; i++)
					{
						Task v1 = _taskList[i] as Task;
						Task v2 = tpo._taskList[i] as Task;
						if (!v1.getObjID().Equals(v2.getObjID()))
							return false;
					}
				}
				else
				{ 
					return false;
				}

				return true;
			}
		}

		Tspd.Icp.SOA _currentSOA = null;
		long _currentArm = ArmRule.ALL_ARMS;
		//ArrayList _periodVisitList = new ArrayList();
		ArrayList _taskList = new ArrayList();
		ArrayList _invalidTC = new ArrayList();
		ArrayList _treatmentPeriods = new ArrayList();
		IFootnoter _footNoter = null;
		public bool	_preprocessErrors = false;
		bool	_isBlinded = false;

		public static readonly string CTMROLE_INVESTIGATIONAL_PRODUCT = "investigationalProduct";
		
		public ScheduleOfProceduresByPeriodMacro(MacroExecutor.MacroParameters mp, bool blinded) : base (mp)
		{
			_isBlinded = blinded;
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region ScheduleOfProcedures
		public static MacroExecutor.MacroRetCd ScheduleOfProceduresByPeriod
			(MacroExecutor.MacroParameters mp) 
		{
			return scheduleOfProceduresByPeriodInternal(mp, false); 
		}

		public static MacroExecutor.MacroRetCd BlindedScheduleOfProceduresByPeriod
			(MacroExecutor.MacroParameters mp) 
		{
			return scheduleOfProceduresByPeriodInternal(mp, true); 
		}

#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.ScheduleOfProceduresByPeriodMacro.ScheduleOfProceduresByPeriod,ProtocolDTs.dll" elementLabel="SOP By Treatment Period" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="Prints table of Schedule of Procedures for a given Schedule of Assessments for each treatment period." shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.ScheduleOfProceduresByPeriodMacro.BlindedScheduleOfProceduresByPeriod,ProtocolDTs.dll" elementLabel="Blinded SOP By Treatment Period" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="Prints blinded table of Schedule of Procedures for a given Schedule of Assessments for each treatment period." shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
#endif

		public static MacroExecutor.MacroRetCd scheduleOfProceduresByPeriodInternal( 
			MacroExecutor.MacroParameters mp, bool blinded) 
		{
			try 
			{
				mp.pba_.setOperation("SOP By Treatment Period Macro", "Generating information...");
				
				ScheduleOfProceduresByPeriodMacro macro = null;
				macro = new ScheduleOfProceduresByPeriodMacro(mp, blinded);
				macro.preProcess();
				if (macro._preprocessErrors == false)
				{
					macro.display(); 
					macro.postProcess();
				}
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in SOP By Treatment Period Macro"); 
				mp.inoutRng_.Text = "Error in SOP By Treatment Period Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}



		#endregion

		#endregion

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

		public override void preProcess() 
		{
			string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);
			if (MacroBaseUtilities.isEmpty(elementPath)) 
			{
				return;
			}

			ArrayList errorList = new ArrayList();

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

			if (_currentSOA == null) 
			{
				_preprocessErrors = true;
				return;
			}


			// Collect the ordered period list filtered by those of type treatment
			ArrayList orderedTopLevelEvents = new ArrayList(); //XX

			ArrayList tmp = new ArrayList();
			_currentSOA.getTopLevelActivityList(_currentArm, null, tmp);
			foreach(EventScheduleBase eb in tmp)
			{
				Period period = eb as Period;
				if (period != null && period.getScheduleItemType().Equals(TREATMENT_TYPE))
					this._treatmentPeriods.Add(new TreatmentPeriod(period));
			}

			if (_treatmentPeriods.Count > 0)
			{
				tmp.Clear();
				// each of these is going to represent one table
				foreach (TreatmentPeriod tp in _treatmentPeriods)
				{
					if (preprocessPeriod(tp, errorList))
						tmp.Add(tp);
					// otherwise there were no dosing visits in this period
				}

				if (tmp.Count == 0)
				{
					addErrorMessage("The Schedule of Activities should contain at least one"+
						" Treatment Period containing a Visit marked as \'Dosing Treatment\'", errorList);
				}
				else
				{
					// now re-initialize our processing array with only those
					// periods we have successfully prepped for.
					_treatmentPeriods.Clear();
					_treatmentPeriods.AddRange(tmp);
				}
			}
			else
			{
				addErrorMessage("The Schedule of Activities should contain at least one"+
					" Period marked as \'Treatment\'", errorList);
			}

			if (errorList.Count > 0)
			{
				emitErrors(errorList);
				_preprocessErrors = true;
				return;
			}
			return;
		}

		/// <summary>
		/// copied in from SOA.cs where it was private in 2.4.5 and slightly modified
		/// </summary>
		/// <param name="child"></param>

		public void constructParentElementPath(EventScheduleBase child)
		{
			IXMLDOMNode correctParent = child.getObjectRoot().parentNode.parentNode;
			if (!(correctParent.nodeName.Equals(IcpDefines.PeriodNode) || correctParent.nodeName.Equals(IcpDefines.SubPeriodNode)))
				return;
			
			IXMLDOMNode seqNode = MsXmlHelper.selectSingleNode(correctParent, BaseProtocolObject.SEQUENCE);
			string predicate = child.getSchedulePathSlash() + 
				"Periods/Period[./sequence=\"" + seqNode.text + "\"]";

			// if our parent is a period then use the simpler path
			// this will happen if we are a sub period or we are a visit with a direct
			// parent that is is a Period.
			if (correctParent.nodeName.Equals(IcpDefines.PeriodNode))
			{
				child.setParentElementPath(predicate);
				return;
			}

				// else we are a Visit and our parent is a sub period.  In this case, the parent
				// element path should trace all the way from the top.

			else if (correctParent.nodeName.Equals(IcpDefines.SubPeriodNode))
			{
				string predicate2 = "Children/SubPeriod[./sequence=\"" + seqNode.text + "\"]";
				correctParent = correctParent.parentNode.parentNode;
				// now construct the uppermost part of the path
				seqNode = MsXmlHelper.selectSingleNode(correctParent, BaseProtocolObject.SEQUENCE);
				predicate = child.getSchedulePathSlash() + "Periods/Period[./sequence=\"" + seqNode.text + "\"]";
				child.setParentElementPath(predicate + "/" + predicate2);
			}
		}

		public class OrderedVisitsBySubperiod : System.Collections.IComparer
		{
			#region IComparer Members
			SOA _soa;

			public OrderedVisitsBySubperiod(SOA soa)
			{
				_soa = soa;
			}

			public int Compare(object x, object y)
			{
				TreatmentVisit ev1 = x as TreatmentVisit;
				TreatmentVisit ev2 = y as TreatmentVisit;
				// first check same subperiod and order
				Period ev1Parent = _soa.getParentOfScheduleItem(ev1._visit);
				Period ev2Parent = _soa.getParentOfScheduleItem(ev2._visit);

				if (ev1Parent.getObjID() == ev2Parent.getObjID())
				{
					return ev1._visit.getSequence() - ev2._visit.getSequence();
				}
				else
				{
					return ev1Parent.getSequence() - ev2Parent.getSequence();
				}
			}

			#endregion
		}



		public bool preprocessPeriod(TreatmentPeriod treatmentPeriod, ArrayList errorList)
		{
			try
			{
				string tpName = treatmentPeriod._period.getBriefDescription();

				ProtocolEventEnumerator pve = 
					new ProtocolEventEnumerator(
					this._currentSOA.getProtocolEventsForPeriod(treatmentPeriod._period),
					this.icdSchemaMgr_.getVisitTemplate(),
					this._currentSOA.getObjID());

				// find all the treatment visits in this period
				while (pve.MoveNext())
				{
					ProtocolEvent visit = (ProtocolEvent)pve.Current;
					constructParentElementPath(visit);
					//visit.setParentElementPath(treatmentPeriod._period.getElementPath());
					visit.getComplexChildren();

					string siType = visit.getScheduleItemType();

					if ( siType != null && 
						siType.Equals(DOSING_TREATMENT) && 
						visit.getEventType().getSubtype().Equals(EventType.EventSubType.Scheduled))
					{
						treatmentPeriod._treatmentVisits.Add(new TreatmentVisit(visit, _currentSOA));
					}
				}

				treatmentPeriod._treatmentVisits.Sort(
					new OrderedVisitsBySubperiod(this._currentSOA));

				// if there aren't any, its an error
				if (treatmentPeriod._treatmentVisits.Count == 0)
				{
					return false;
				}

				// Collect the ordered tasks
				ArrayList includedTasks = new ArrayList();
				IList taskList = new ElementListHelpers(tspdDoc_).
					getLiveChooserEntryListForTasks(this._currentSOA);
				foreach (Task task in taskList)
				{
					// Collect by Dosing Task
					if (task.isDosingTask()) 
					{
						DosingTask dosingTask = new DosingTask(
							task.getObjectRoot(), icdSchemaMgr_.getTemplateByClass(typeof(DosingTask)));

						long ctmID = dosingTask.getctMaterialID();

						ClinicalTrialMaterial ctm = findCTM(ctmID);
						string ctmRole = ctm.getPrimaryRole();

						if (!MacroBaseUtilities.isEmpty(ctmRole) && ctmRole.Equals(CTMROLE_INVESTIGATIONAL_PRODUCT)) 
						{
							includedTasks.Add(dosingTask);
						}
					}
				}
				if (includedTasks.Count == 0)
				{
					string msg = "There must be a Dosing Task linked to an Investigational-Product declared " +
						"in the Schedule of Activities";
					this.addErrorMessage(msg, errorList);
					// no point going ahead without this
					return false;
				}


				// be optimistic that the dosing tasks will match timepoint
				// we will exit below if this is not the case
				treatmentPeriod._anchorTask = includedTasks[0] as Task;
				bool anchorTaskIsChunked = false;

				// find the timechunk list for each visit and also get a list of the usedTasks
				Hashtable uniqueTasks = new Hashtable();
				int visitIndex = 0;
				foreach (TreatmentVisit tv in treatmentPeriod._treatmentVisits)
				{
					ArrayList tcList = PEExploder.getTimeChunks(_currentSOA, tv._visit);
					// Accumulation of time chunks in this Visit
					
					foreach (TimeChunk tc in tcList) 
					{
						if (tc.isDefaultLabel()) 
						{
							continue;
						}
	
						string startTime;
						string endTime;
						string unit;
						string serr;
	
						bool success = PurdueUtil.parseTimePoint(tc.Label, out startTime, out endTime, out unit, out serr);
						if (!success) 
						{
							string errorMessage = "Period: " + tpName + ", Visit: " +
								tv._visit.getBriefDescription() + ", Time: " + tc.Label + ": " + serr;

							this.addErrorMessage(errorMessage, errorList);
							continue;
						}
	
						// Add time chunks
						// make sure we add a visit index in case the 
						PurdueUtil.TCWrapper tcwrap = new PurdueUtil.TCWrapper(tc, startTime, endTime, unit);
						tcwrap.DayIndex = visitIndex;
						
						if (tcwrap.isSpan())
							tv._tvSpanCells.Add(tcwrap);
						else
							tv._tvCells.Add(tcwrap);

						// check that at least one dose occurs at the 0 timepoint and use that as the anchor
						foreach (Task incTask in includedTasks)
						{
							ArrayList tvList = tc.getByTaskID(incTask.getObjID());
							if (tvList.Count != 0) 
							{
								if (!tc.Label.Equals("0"))
									continue;
								if (!anchorTaskIsChunked)
								{
									treatmentPeriod._anchorTask = incTask;
									anchorTaskIsChunked = true;
								}
							}
						}
					}


					// get the tasks used in this visit and add to the unique set
					foreach (Task task in taskList)
					{
						PurdueUtil.PeriodAndVisit pv = new PurdueUtil.PeriodAndVisit();
						pv.per = treatmentPeriod._period;
						pv.visit = tv._visit;
						pv.tcList = tv._tvCells;
						if (this.isTaskUsed(task, pv))
						{
							// if we are generating a blinded table, then we do not include
							// dosing tasks 
							if ((!_isBlinded) || (!task.isDosingTask()))
							{
								uniqueTasks[task.getObjID()] = task;
							}
						}
					}
				}

				if (anchorTaskIsChunked == false)
				{
					string errorMessage = "The Dosing Task: " + 
						treatmentPeriod._anchorTask.getBriefDescription() + 
						" is not allocated to the \"0\" time slot in any of the visits of Period: " + tpName;
					this.addErrorMessage(errorMessage, errorList);
				}
				
				
				// now populate the Periods's task list with the unique set
				// in the order declared
				foreach (Task task in taskList)
				{
					Task tt = uniqueTasks[task.getObjID()] as Task;
					if (tt != null)
					{
						treatmentPeriod._taskList.Add(tt);
					}
				}
			}
			finally
			{
				if (errorList.Count > 0) 
				{
					_preprocessErrors = true;
				}
#if debug
				if (true && _preprocessErrors == false)
				{
					Word.Range wrkRng = this.inoutRng_.Duplicate;

					wrkRng.InsertAfter("Treatment Period: " + 
						treatmentPeriod._period.getBriefDescription() + 
						" has " + treatmentPeriod._treatmentVisits.Count + " qualifying  visits");
					wrkRng.InsertParagraphAfter();
					wrkRng.InsertAfter("AnchorTask: " + treatmentPeriod._anchorTask.getBriefDescription());
					wrkRng.InsertParagraphAfter();

					foreach (TreatmentVisit tv in treatmentPeriod._treatmentVisits)
					{
						wrkRng.InsertAfter("Visit: " + tv._visit.getBriefDescription());
						wrkRng.InsertParagraphAfter();
						if (tv._tvCells != null)
						{
							foreach (PfizerUtil.TCWrapper tcw in tv._tvCells)
							{
								string msg = "Chunk: \"" + tcw.Label + 
									"\" minutes from= " + tcw.getStartMinute();
								
								wrkRng.InsertAfter(msg);
								wrkRng.InsertParagraphAfter();
							}
							wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
							
							foreach (PfizerUtil.TCWrapper tcw in tv._tvSpanCells)
							{
								string msg = "Chunk: \"" + tcw.Label + 
									"\" minutes from= " + tcw.getStartMinute();
								msg += " to " + tcw.getEndMinute();
								wrkRng.InsertAfter(msg);
								wrkRng.InsertParagraphAfter();
							}
							wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
						}
					}
					wrkRng.InsertParagraphAfter();
					wrkRng.InsertAfter("Tasks to be considered: ");
					foreach (Task task in treatmentPeriod._taskList)
					{
						wrkRng.InsertAfter(task.getBriefDescription() + ",");
					}
					wrkRng.InsertParagraphAfter();
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				}
#endif
			}
			return true;
		}

		private bool isTaskUsed(Task task, PurdueUtil.PeriodAndVisit pv) 
		{
			// Check if this task is used in a labeled task visit
			foreach (PurdueUtil.TCWrapper tcw in pv.tcList) 
			{
				TimeChunk tc = tcw.tc;

				if (tc.isDefaultLabel()) 
				{
					continue;
				}

				ArrayList tvList = tc.getByTaskID(task.getObjID());
				foreach (TaskVisit tvv in tvList)
				{
					if(!LittleUtilities.isEmpty(tvv.getLabel()))
					{
						return true;
					}
				}
			}

			return false;
		}

		private ClinicalTrialMaterial findCTM(long ctmID) 
		{
			CTMaterialEnumerator ctEnum = bom_.getCTMaterialEnumerator();
			while (ctEnum.MoveNext()) 
			{
				ClinicalTrialMaterial ctm = ctEnum.getCurrent();
				if (ctm.getObjID() == ctmID) 
				{
					return ctm;
				}
			}

			return null;
		}

		public override void display()
		{
			Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;

			pba_.updateProgress(1.0);
			
			string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);
			if (MacroBaseUtilities.isEmpty(elementPath)) 
			{
				macroStatusCode_ = MacroExecutor.MacroRetCd.Failed;
				return;
			}

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

			if (_treatmentPeriods.Count == 0) 
			{
				pba_.updateProgress(70.0);

				wrkRng.InsertAfter("There are no Periods with dosing tasks for the investigational product defined.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
						
				inoutRange.End = wrkRng.End;
				setOutgoingRng(inoutRange);
				wdDoc_.UndoClear();
				return;
			}

			if (_invalidTC.Count != 0) 
			{
				pba_.updateProgress(70.0);

				wrkRng.InsertAfter("There are invalid time points:");
				wrkRng.InsertParagraphAfter();

				foreach (string s in _invalidTC) 
				{
					wrkRng.InsertAfter(s);
					wrkRng.InsertParagraphAfter();
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				}

				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
						
				inoutRange.End = wrkRng.End;
				setOutgoingRng(inoutRange);
				wdDoc_.UndoClear();
				return;
			}

			_footNoter = new Table.PurdueFootnoter(bom_, wdDoc_);

			wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			// merge identical periods as we generate the tables
			for (int startPeriod = 0; startPeriod < _treatmentPeriods.Count; startPeriod++) 
			{
				TreatmentPeriod tp = _treatmentPeriods[startPeriod] as TreatmentPeriod; 
				if (startPeriod == _treatmentPeriods.Count-1)
				{
					// last one.. so no more to merge with
					displayTreatmentPeriod(tp, null, ref wrkRng);
					break;
				}
				else
				{
					int lastMergedIndex = -1;
					// walk the periods looking to merge with the current one.
					// we stop when we've found the first one that is different
					// so print the merged periods (if any) then the singular one.
					for (int endOffset = 1; (startPeriod + endOffset) < _treatmentPeriods.Count; endOffset++)
					{
						if (tp.Equals(_treatmentPeriods[startPeriod + endOffset]))
						{
							lastMergedIndex = startPeriod + endOffset;
							continue;
						}
						if (lastMergedIndex > -1) 
						{
							// print the merged periods
							TreatmentPeriod lastMerged = _treatmentPeriods[lastMergedIndex] as TreatmentPeriod;
							displayTreatmentPeriod(tp, lastMerged, ref wrkRng);

							// print the next one
							TreatmentPeriod next = _treatmentPeriods[lastMergedIndex+1] as TreatmentPeriod;
							displayTreatmentPeriod(next, null, ref wrkRng);
							startPeriod = startPeriod + endOffset;
							lastMergedIndex = -1;
						}
						else // no merges this time
						{
							displayTreatmentPeriod(tp, null, ref wrkRng);
						}
						break;
					}
					// the last period is merged with its predecessors
					if (lastMergedIndex > -1) 
					{
						TreatmentPeriod lastMerged = _treatmentPeriods[lastMergedIndex] as TreatmentPeriod;
						displayTreatmentPeriod(tp, lastMerged, ref wrkRng);
						startPeriod = lastMergedIndex;
					}
				}
			}

			if (_footNoter.hasFootnotes()) 
			{
				Word.Table fnTbl = createFootnoteTable(wrkRng, _footNoter.getFootnotes().Count);
				fillFootnoteTable(fnTbl, _footNoter.getFootnotes());
			}

			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}

		private void displayTreatmentPeriod(TreatmentPeriod tp, TreatmentPeriod endTP, ref Word.Range wrkRng)
		{
			#region more pre-processing

			ArrayList errorList = new ArrayList();
			ArrayList usedTasks = tp._taskList;

			ArrayList tcList = new ArrayList();
			ArrayList tcSpannedList = new ArrayList();

			// create column chunks across all visits of this period
			foreach (TreatmentVisit tv in tp._treatmentVisits)
			{
				tcSpannedList.AddRange(tv._tvSpanCells);
				tcList.AddRange(tv._tvCells);
			}

			// Add missing labels for spanned times whose boundaries do not live on an already created
			// column... (watch for spanning beyond existing columns too:
			// columns   8 10 12, with a chunk thats 8-11.  It should create an '11' column
			ArrayList addList =  new ArrayList();
			foreach (PurdueUtil.TCWrapper tcwSpan in tcSpannedList) 
			{
				pba_.updateProgress(1.0);

				bool foundStart = false;
				bool foundEnd = false;

				// Match spanned start/end+unit against start+unit
				foreach (PurdueUtil.TCWrapper tcw in tcList) 
				{
					if (tcwSpan.getStartMinute() == tcw.getStartMinute())
					{
						foundStart = true;
					}
				
					if (tcwSpan.getEndMinute() == tcw.getStartMinute()) 
					{
						foundEnd = true;
					}

					if (foundStart && foundEnd) 
					{
						break;
					}
				}

				if (!foundStart) 
				{
					bool doAdd = true;
					// we may have already added it in the outer loop
					foreach (PurdueUtil.TCWrapper tcwa in addList)
					{
						if(tcwa.getStartMinute() == tcwSpan.getStartMinute())
						{
							doAdd = false;
							break;
						}
					}
					// add a column to align with the start of a span if the column does not exist
					if (doAdd) 
					{
						addList.Add(new PurdueUtil.TCWrapper(null, tcwSpan.Start, "", tcwSpan.Unit));
					}
				}

				if (!foundEnd) 
				{
					bool doAdd = true;
					// we may have already added it in the outer loop
					foreach (PurdueUtil.TCWrapper tcwa in addList)
					{
						if(tcwa.getStartMinute() == tcwSpan.getEndMinute())
						{
							doAdd = false;
							break;
						}
					}
					// add a column to align with the end of a span if the column does not exist
					// but Don't add endpoint label of 24 hrs
					if (doAdd == true && !(tcwSpan.End.Equals("24") && tcwSpan.Unit.Equals(PurdueUtil.TimeUnit.sHOURS)))
					{
						addList.Add(new PurdueUtil.TCWrapper(null, tcwSpan.End, "", tcwSpan.Unit));
					}
				}
			}

			// Add the fixup columns now
			tcList.AddRange(addList);
			

			// None?
			if (tcList.Count == 0)
			{
				return;
			}

			// Sort both column lists by start times
			tcList.Sort(new PurdueUtil.TCWrapperComparer());
			tcSpannedList.Sort(new PurdueUtil.TCWrapperComparer());
#if debug
			Word.Range dbgRange = wrkRng.Duplicate;
			dbgRange.Collapse(ref WordHelper.COLLAPSE_END);
			object o = this.tspdDoc_.getStyleHelper().getNormalStyle();
			dbgRange.set_Style(ref o);
			dbgRange.InsertAfter("Spans in order");
			dbgRange.InsertParagraphAfter();
			foreach(PfizerUtil.TCWrapper tcw in tcSpannedList)
			{
				dbgRange.InsertAfter("Span:" + tcw.Label + " (" + tcw.getStartMinute() + "," + 
					tcw.getEndMinute() + ") start:" + tcw.Start + " end: " + tcw.End);
				dbgRange.InsertParagraphAfter();
			}
			dbgRange.InsertAfter("All Displayed Columns in Order");
			dbgRange.InsertParagraphAfter();
			foreach(PfizerUtil.TCWrapper tcw in tcList)
			{
				dbgRange.InsertAfter("Column:" + tcw.Label + " (" + tcw.getStartMinute() + 
					") start:" + tcw.Start + " end: " + tcw.End);
				dbgRange.InsertParagraphAfter();
			}
			dbgRange.Collapse(ref WordHelper.COLLAPSE_END);
			wrkRng.End = dbgRange.End;
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
#endif 

			// Collect task visits into matrix
			TVCell[,] tvCells = new TVCell[usedTasks.Count, tcList.Count];

			// for each used task
			for (int taskRow = 0; taskRow < usedTasks.Count; taskRow++)
			{
				pba_.updateProgress(1.0);

				Task t1 = usedTasks[taskRow] as Task;

				// for each  unspanned column
				for (int cc = 0; cc < tcList.Count; cc++) 
				{
					TVCell tvCell = new TVCell();
					// insert cell into the matrix
					tvCells[taskRow, cc] = tvCell;

					PurdueUtil.TCWrapper tcw = tcList[cc] as PurdueUtil.TCWrapper;
					// get the inner time chunk from the wrapper
					TimeChunk tc = tcw.tc;

					// Collect task visits
					if (tc != null) 
					{
						// get any task visits that of the given task that should be done in 
						// this chunk.  Normally only one... but you never know.
						ArrayList a = tc.getByTaskID(t1.getObjID());
						foreach(TaskVisit tvv in a)
						{
							// filter now because this chunker picks up
							// unlabeled items following the last labeled tv
							// and adds them to the chunk.  But we don't want them here
							if (!LittleUtilities.isEmpty(tvv.getLabel()))
							{
								tvCell.tvList.AddRange(a);
							}
						}
					}
				}
			}

			// Find spanning.  In this process we try to match up each span with the columns
			// in matrix we've created. 
			for (int taskRow = 0; taskRow < usedTasks.Count; taskRow++)
			{
				pba_.updateProgress(1.0);

				Task t1 = usedTasks[taskRow] as Task;

				string sErrTask = "\r\n" + t1.getBriefDescription();

				foreach (PurdueUtil.TCWrapper tcwSpan in tcSpannedList) 
				{
					// get the task visits for the spanned column for the given task
					// there should be one since it was a task visit that
					// created the span in the first place.
					ArrayList tvList = tcwSpan.tc.getByTaskID(t1.getObjID());
					if (tvList.Count == 0) 
					{
						continue;
					}

					// each spanned cell in a row can only refer to one instance of that task
					// unless that task has a multi visit span
					if (tvList.Count != 1 && !tcwSpan.isMultiDaySpan()) 
					{
						addErrorMessage(sErrTask, errorList);

						string serr = tcwSpan.Label + " overlap (1) with: " + tcwSpan.Label;
						addErrorMessage(serr, errorList);
					}

					bool startEmpty = true;
					bool foundStart = false;
					// the number of empty cells encountered
					// where this span can be placed
					int  emptyCount = 0;

					TVCell startCell = null;
					// for each unspanned column
					for (int cc = 0; cc < tcList.Count; cc++) 
					{
						PurdueUtil.TCWrapper tcw = tcList[cc] as PurdueUtil.TCWrapper;
						TVCell tvCell = tvCells[taskRow, cc];

						// Match spanned start against start
						if (tcwSpan.getStartMinute() == tcw.getStartMinute())
						{
							// Found first proper column so remember the cell
							// this span begins at
							foundStart = true;
							startCell = tvCell;

							if (tvCell.tvList.Count == 0) 
							{
								emptyCount++;
							}
							else
							{
								startEmpty = false;
							}

							// If the cell already has a span, error
							if (!MacroBaseUtilities.isEmpty(tvCell.spanLabel)) 
							{
								if (tcwSpan.isMultiDaySpan())
									continue;

								addErrorMessage(sErrTask, errorList);
								string serr = tcwSpan.Label + " overlap (2) with: " + tvCell.spanLabel;
								addErrorMessage(serr, errorList);
								break;
							}

							// copy the the span's label into the cell's label
							tvCell.spanLabel = tcwSpan.Label;
						}


						// end - done with this span
						if (tcwSpan.getEndMinute() <= tcw.getStartMinute()) 
						{
							break;
						}
						// if the span has begun, but we are one or more cells along
						if (foundStart && tvCell != startCell) 
						{
							// If the start cell is empty and we have an X, error
							// I think this because we only match the TV with the first cell on the
							// left of a group of cells we intend to merge.

							if (startEmpty && tvCell.tvList.Count != 0) 
							{
								if (tcwSpan.isMultiDaySpan())
									continue;

								addErrorMessage(sErrTask, errorList);

								string serr = tcwSpan.Label + " overlap (3) with: " + tcw.Label;
								addErrorMessage(serr, errorList);
								break;
							}

							// If the cell already has a span, error
							if (!MacroBaseUtilities.isEmpty(tvCell.spanLabel)) 
							{
								if (tcwSpan.isMultiDaySpan())
									continue;

								addErrorMessage(sErrTask, errorList);

								string serr = tcwSpan.Label + " overlap (4) with: " + tvCell.spanLabel;
								addErrorMessage(serr, errorList);
								break;
							}
							
							if (tvCell.tvList.Count == 0) 
							{
								emptyCount++;
							}
							// every cell in the span gets the same label
							tvCell.spanLabel = tcwSpan.Label;
						}
					}

					if (foundStart && emptyCount == 0) 
					{
						addErrorMessage(sErrTask, errorList);

						string serr = tcwSpan.Label + " no open cell ";
						addErrorMessage(serr, errorList);
					}
				}
			}

			if (errorList.Count != 0) 
			{
				emitErrors(errorList);
				return;
			}

			#region dumptable
			
			if (false) 
			{
				for (int taskRow = 0; taskRow < usedTasks.Count; taskRow++)
				{
					Task t1 = usedTasks[taskRow] as Task;

					for (int cc = 0; cc < tcList.Count; cc++) 
					{
						TVCell tvCell = tvCells[taskRow, cc];

						string msg = taskRow.ToString() + ", " + cc.ToString() + ": " +
							t1.getBriefDescription() + ", tvCount: " + 
							tvCell.tvList.Count.ToString() + ", spanLabel: " +
							tvCell.spanLabel;
					
						wrkRng.InsertAfter(msg);
						wrkRng.InsertParagraphAfter();
						wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
					}

					wrkRng.InsertParagraphAfter();
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				}
			}

			#endregion

			#endregion

			#region build table

			// Build out the table
			int tableRows = usedTasks.Count + 3;
			int tableCols = tcList.Count + 1;

			// Column/Row
			int PERIOD_VISIT_ROW = 1;
			int PROCEDURE_ROW = 2;
			int HOURS_HEADER_ROW = 2;
			int TIMEPOINT_HEADER_ROW = 3;
			int TASK_ROW = 4;

			int TABLE_COL_OFFSET = 2;

			#region table body

			if (tableCols > 63) 
			{
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				wrkRng.InsertAfter("Error: table for period: " + tp._period.getBriefDescription() + 
					" has more columns that Microsoft Word allows (63)");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				return;
			}

			Word.Table tbl = createTable(ref wrkRng, tableRows, tableCols);


			SOATableFormat tblFmt = tspdDoc_.getSOATblFormat(_currentSOA.getObjID());
			if (tblFmt == null)
			{
				// create a fake one to run thru defaults...
				TrialDocument.SOATableFormatCV cv = new TrialDocument.SOATableFormatCV();
				tblFmt = cv.newSOATableFormat();
			}

			Word.Font targetFont;

			object oStyle = tspdDoc_.getStyleHelper().setNamedStyle(PurdueUtil.PFIZER_STYLE_TABLETEXT_10, tbl.Range);

			tbl.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

			// Set Table Body font
			targetFont = tbl.Range.Font;
			setTableBodyFont(ref targetFont, tblFmt);

			Word.Row row = null;
			Word.Cell cell = null;

			// Period/Visit header
			// instert   <Period>,<Visitx>-<Visity>
			row = tbl.Rows[PERIOD_VISIT_ROW];
			cell = row.Cells[1];

			string startPeriodName = (tp != null && tp._period != null) 
				? tp._period.getBriefDescription() 
				: null;

			string endPeriodName = (endTP != null && endTP._period != null) 
				? endTP._period.getBriefDescription() 
				: null;


			string periodRange = makePeriodRangeString(tp._period.getBriefDescription(), endPeriodName);
			cell.Range.Text = periodRange;
			
			Word.Range clRng = cell.Range.Duplicate;
			clRng.Collapse(ref WordHelper.COLLAPSE_END);
			MacroBaseUtilities.putAfterElemRef(", ", tspdDoc_, clRng);
			
			// Here's the logic
			// you may have a single period with a single visit
			// you may have a set of visits in one period
			// you may have a series of contiguous periods each with 1 or more visits
			// Note that visits that are not of type dosing-treatment, or which do not use task 
			//    sequence labels will be ignored

			// P1, V1
			// P1, V1-VN
			// P1-PN, V1(P1)-VN(PN)
		
			int visitCount = tp._treatmentVisits.Count;
			ProtocolEvent ev1 = ((TreatmentVisit)(tp._treatmentVisits[0]))._visit;
			ProtocolEvent ev2 = ((TreatmentVisit)(tp._treatmentVisits[visitCount-1]))._visit;
			if (endTP != null)
			{
				visitCount = endTP._treatmentVisits.Count;
				ev2 = ((TreatmentVisit)(endTP._treatmentVisits[visitCount-1]))._visit;
			}
			

			if (ev1.getObjID() == ev2.getObjID())
			{
				MacroBaseUtilities.putElemRefInCell(tspdDoc_, cell, ev1, ProtocolEvent.BRIEF_DESCRIPTION, false, macroEntry_);
			}

			else
			{
				MacroBaseUtilities.putElemRefInCell(tspdDoc_, cell, ev1, ProtocolEvent.BRIEF_DESCRIPTION, false, macroEntry_);
				// need extra space
				clRng = cell.Range.Duplicate;
				clRng.Collapse(ref WordHelper.COLLAPSE_END);
				MacroBaseUtilities.putAfterElemRef("- ", tspdDoc_, clRng);
				MacroBaseUtilities.putElemRefInCell(tspdDoc_, cell, ev2, ProtocolEvent.BRIEF_DESCRIPTION, false, macroEntry_);
			}

			targetFont = cell.Range.Font;
			setPeriodFont(ref targetFont, tblFmt);

			// Procedure header
			row = tbl.Rows[PROCEDURE_ROW];
			cell = row.Cells[1];
			cell.Range.Text = "Procedure";
			cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
			targetFont = cell.Range.Font;
			setTaskHeaderFont(ref targetFont, tblFmt);

			// Hours header
			row = tbl.Rows[HOURS_HEADER_ROW];

			ArrayList hoursHeader = new ArrayList();

			for (int cc = 0; cc < tcList.Count; cc++) 
			{
				pba_.updateProgress(1.0);

				PurdueUtil.TCWrapper tcw = tcList[cc] as PurdueUtil.TCWrapper;

				cell = row.Cells[cc + TABLE_COL_OFFSET];
				cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

				double start1 = double.Parse(tcw.Start);

				string sUnit = "";
				string sTime = "";

				if (tcw.Unit.Equals(PurdueUtil.TimeUnit.sMINUTES)) 
				{
					sUnit = "(min)";
				}
				else if (tcw.Unit.Equals(PurdueUtil.TimeUnit.sHOURS)) 
				{
					sUnit = "(h)";
				}

				if (start1 < 0) 
				{
					sTime = "Predose " + sUnit;
				}
				else if (start1 > 0)
				{
					sTime = "Postdose " + sUnit;
				}
				// no label if 0

				hoursHeader.Add(sTime);
				cell.Range.Text = sTime;


				targetFont = cell.Range.Font;
				setVisitFont(ref targetFont, tblFmt);
			}

			// Timechunk label header
			row = tbl.Rows[TIMEPOINT_HEADER_ROW];

			// Set heading row
			row.HeadingFormat = VBAHelper.iTRUE;

			for (int cc = 0; cc < tcList.Count; cc++) 
			{
				PurdueUtil.TCWrapper tcw = tcList[cc] as PurdueUtil.TCWrapper;

				cell = row.Cells[cc + TABLE_COL_OFFSET];

				cell.Range.Text = tcw.Label;

				cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

				targetFont = cell.Range.Font;
				setVisitFont(ref targetFont, tblFmt);
			}

			// Tasks header
			for (int taskRow = 0; taskRow < usedTasks.Count; taskRow++)
			{
				pba_.updateProgress(1.0);

				Task t1 = usedTasks[taskRow] as Task;

				row = tbl.Rows[taskRow + TASK_ROW];
				cell = row.Cells[1];
				cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

				MacroBaseUtilities.putElemRefInCell(tspdDoc_, cell,	t1, Task.BRIEF_DESCRIPTION, true, macroEntry_);
				targetFont = cell.Range.Font;
				setTaskHeaderFont(ref targetFont, tblFmt);

				putFootnote(cell.Range, t1);
				
				// Fill in check marks
				for (int cc = 0; cc < tcList.Count; cc++) 
				{
					TVCell tvCell = tvCells[taskRow, cc];

					cell = row.Cells[cc + TABLE_COL_OFFSET];
					cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

					if (tvCell.tvList.Count != 0) 
					{
						// Add a check for each task visit
						foreach (TaskVisit tv in tvCell.tvList)
						{
							clRng = cell.Range.Duplicate;
							clRng.End--;
							clRng.Collapse(ref WordHelper.COLLAPSE_END);
							clRng.InsertAfter("X");
							clRng.Font.Superscript = VBAHelper.iFALSE;

							putFootnote(cell.Range, tv);
						}
					}
					else if (!MacroBaseUtilities.isEmpty(tvCell.spanLabel)) 
					{
						clRng = cell.Range.Duplicate;
						clRng.End--;
						clRng.Collapse(ref WordHelper.COLLAPSE_END);
						clRng.InsertAfter(tvCell.spanLabel);
					}
				}
			}

			// Turn off autofit now
			tbl.AllowAutoFit = false;

			// Make first column bigger
			foreach (Word.Row rw in tbl.Rows)
			{
				try 
				{
					if (rw.Cells.Count > 1 && rw.Cells.Count < 10) 
					{
						Word.Cell cl = rw.Cells[1];
						cl.SetWidth(tbl.Application.InchesToPoints(1.25f), Word.WdRulerStyle.wdAdjustProportional);
					}
				}
				catch (Exception ex) 
				{
					Log.exception(ex, "Error setting column 1 width");
				}
			}


			#endregion table body

			#region Cell Merging

			// Cell merging
#if true
			// Period/Visit header
			row = tbl.Rows[PERIOD_VISIT_ROW];
			Word.Cell c1 = row.Cells[1];
			Word.Cell c2 = row.Cells[tableCols];
			c1.Merge(c2);

			// Hours header
			row = tbl.Rows[HOURS_HEADER_ROW];

			int firstcol = 0;
			ArrayList mergeList = new ArrayList();

			// Scan cells for same text, merge
			for (int cc = 0; cc < tcList.Count-1; cc++) 
			{
				int nextcol = cc + 1;

				string curtxt = hoursHeader[cc] as string;
				string nexttxt = hoursHeader[nextcol] as string;

				if (!curtxt.Equals(nexttxt)) 
				{
					if (firstcol != cc) 
					{
						mergeList.Add(new PurdueUtil.MergePair(curtxt, HOURS_HEADER_ROW, 
							firstcol + TABLE_COL_OFFSET, cc + TABLE_COL_OFFSET));
					}

					firstcol = nextcol;
				}
				else if (nextcol == tcList.Count-1)
				{
					if (firstcol != nextcol) 
					{
						mergeList.Add(new PurdueUtil.MergePair(curtxt, HOURS_HEADER_ROW, 
							firstcol + TABLE_COL_OFFSET, nextcol + TABLE_COL_OFFSET));
					}
				}
			}

			// Merge
			PurdueUtil.MergePair.merge(tbl, mergeList);

			// Merge spans
			for (int taskRow = 0; taskRow < usedTasks.Count; taskRow++)
			{
				pba_.updateProgress(1.0);

				row = tbl.Rows[taskRow + TASK_ROW];

				firstcol = -1;

				// Scan cells for same text, merge
				for (int cc = 0; cc < tcList.Count-1; cc++) 
				{
					int nextcol = cc + 1;

					TVCell tvCell = tvCells[taskRow, cc];
					TVCell tvCellNext = tvCells[taskRow, nextcol];

					string curtxt = tvCell.spanLabel;
					string nexttxt = tvCellNext.spanLabel;

					// If the cell has tv or is empty skip, reset
					if (tvCell.tvList.Count != 0 ||
						MacroBaseUtilities.isEmpty(tvCell.spanLabel)) 
					{
						firstcol = -1;
						continue;
					}

					// first col?
					if (firstcol == -1) 
					{
						firstcol = cc;
					}

					// cell text changed, merge
					if (!curtxt.Equals(nexttxt)) 
					{
						if (firstcol != cc) 
						{
							mergeList.Add(new PurdueUtil.MergePair(curtxt, taskRow + TASK_ROW, 
								firstcol + TABLE_COL_OFFSET, cc + TABLE_COL_OFFSET));
						}

						firstcol = nextcol;
					}
					else if (nextcol == tcList.Count-1)
					{
						if (firstcol != nextcol) 
						{
							mergeList.Add(new PurdueUtil.MergePair(curtxt, taskRow + TASK_ROW, 
								firstcol + TABLE_COL_OFFSET, nextcol + TABLE_COL_OFFSET));
						}
					}
				}
			}

			// Merge
			PurdueUtil.MergePair.merge(tbl, mergeList);


			// !!!!!  
			// Vertical merging must be done last, 
			// because after any vertical merge word can't access table rows

#if false
			// Procedure Header, vertical merge !!
			row = tbl.Rows[PROCEDURE_ROW);
			c1 = row.Cells[1);
			row = tbl.Rows[TASK_ROW - 1);
			c2 = row.Cells[1);
			
			try 
			{
				c1.Merge(c2);
			} 
			catch (Exception ex) {}
#else

			// Procedure Header, hide border simulate merge
			row = tbl.Rows[PROCEDURE_ROW];
			c1 = row.Cells[1];
			row = tbl.Rows[PROCEDURE_ROW + 1];
			c2 = row.Cells[1];

			c1.Borders[Word.WdBorderType.wdBorderBottom].Visible = false;
			c2.Borders[Word.WdBorderType.wdBorderTop].Visible = false;
#endif

#endif
			#endregion Cell Merging

			#endregion  build table

			// Done with merging
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			wrkRng.InsertParagraphAfter();
			oStyle = tspdDoc_.getStyleHelper().setNamedStyle(PurdueUtil.PFIZER_STYLE_TABLETEXT_10, wrkRng);
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			
			wdDoc_.UndoClear();
		}

		private string makePeriodRangeString(string string1, string string2)
		{
			if (string1 == null) 
			{
				return "Missing Period Name";
			}

			if (string2 == null)
			{
				return string1;
			}

			int endSegment = string2.LastIndexOf(' ');
			int startSegment = string1.LastIndexOf(' ');
			if (endSegment == -1 || startSegment == -1 || (endSegment != startSegment))
				return  string1 + " - " + string2;
			if (!string1.Substring(0, startSegment).Equals(string2.Substring(0, startSegment)))
				return  string1 + " - " + string2;
			return string1.Substring(0, startSegment) + " " 
				+ string1.Substring(startSegment+1)
				+ " - "
				+ string2.Substring(endSegment+1);
		}

		/// <summary>
		/// helper method used by merger to bypass errormessages when we encounter
		/// a second or subsequent task visit and we have already passed by and populated
		/// the first spanned columns.
		/// </summary>
		/// <param name="tvCell"></param>
		/// <param name="tcwSpan"></param>
		/// <returns></returns>
		private bool passingOurselves(TVCell tvCell, PurdueUtil.TCWrapper tcwSpan)
		{
			return (tcwSpan.isMultiDaySpan() && tcwSpan.Label.Equals(tvCell.spanLabel));
		}

		private void emitErrors(ArrayList errorList)
		{
			Word.Range wrkRng = startAtBeginningOfParagraph();
			int start = inoutRng_.Start;
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			for (int i = 0; i < errorList.Count; i++) 
			{
				string err = errorList[i] as String;

				wrkRng.InsertAfter(err);

				if (i < errorList.Count-1) 
				{
					wrkRng.InsertAfter(", ");
				}
			}
			wrkRng.InsertParagraphAfter();
			setOutgoingRng(start, wrkRng.End);
		}

		private void addErrorMessage(string serr, ArrayList errorList) 
		{
			if (!errorList.Contains(serr)) 
			{
				errorList.Add(serr);
			}
		}

		private void setTableBodyFont(ref Word.Font targetFont, SOATableFormat tblFmt) 
		{
			if (tblFmt.getDocTableBodyFontName() != null)
			{
				targetFont.Name = tblFmt.getDocTableBodyFontName();
			}

			if (tblFmt.getDocTableBodyFontSize() != null)
			{
				targetFont.Size = tblFmt.getDocTableBodyFontSize();
			}

			targetFont.Bold = (tblFmt.getDocTableBodyFontBold() ? -1 : 0);	
			targetFont.Italic = (tblFmt.getDocTableBodyFontItalics() ? -1 : 0);
		}

		private void setPeriodFont(ref Word.Font targetFont, SOATableFormat tblFmt) 
		{
			if (tblFmt.getDocPeriodFontName() != null)
			{
				targetFont.Name = tblFmt.getDocPeriodFontName();
			}

			if (tblFmt.getDocPeriodFontSize() != null)
			{
				targetFont.Size = tblFmt.getDocPeriodFontSize();
			}

			targetFont.Bold = (tblFmt.getDocPeriodFontBold() ? -1 : 0);	
			targetFont.Italic = (tblFmt.getDocPeriodFontItalics() ? -1 : 0);
		}

		private void setVisitFont(ref Word.Font targetFont, SOATableFormat tblFmt) 
		{
			if (tblFmt.getDocVisitFontName() != null)
			{
				targetFont.Name = tblFmt.getDocVisitFontName();
			}

			if (tblFmt.getDocVisitFontSize() != null)
			{
				targetFont.Size = tblFmt.getDocVisitFontSize();
			}

			targetFont.Bold = (tblFmt.getDocVisitFontBold() ? -1 : 0);	
			targetFont.Italic = (tblFmt.getDocVisitFontItalics() ? -1 : 0);
		}

		private void setTaskHeaderFont(ref Word.Font targetFont, SOATableFormat tblFmt) 
		{
			if (tblFmt.getDocTaskFontName() != null)
			{
				targetFont.Name = tblFmt.getDocTaskFontName();
			}

			if (tblFmt.getDocTaskFontSize() != null)
			{
				targetFont.Size = tblFmt.getDocTaskFontSize();
			}

			targetFont.Bold = (tblFmt.getDocTaskFontBold() ? -1 : 0);	
			targetFont.Italic = (tblFmt.getDocTaskFontItalics() ? -1 : 0);
		}

		private void putFootnote(Word.Range r, SOAObject sobj) 
		{
			Word.Range fnRng = r.Duplicate;
			fnRng.Collapse(ref WordHelper.COLLAPSE_END);
			fnRng.End--;
			_footNoter.putAtRng(sobj, fnRng);
		}

		private Word.Table createTable(ref Word.Range wrkRng, int rows, int cols) 
		{
			// Turn off auto caption for Word tables.
			Word.AutoCaption ac = wdApp_.AutoCaptions.get_Item(ref WordHelper.AUTO_CAPTION_WORD_TABLE);
			bool oldState = ac.AutoInsert;
			ac.AutoInsert = false;

			if (false) 
			{
				object tableCaption = "\ttable caption";
				wrkRng.InsertAfter(" "); // single space which'll get shifted after the caption.
				Word.Range captionRng = wrkRng.Duplicate;
				captionRng.Collapse(ref WordHelper.COLLAPSE_START);

				// put the caption in if the TableView gives you one...
				if (tableCaption != null)
				{
					captionRng.InsertCaption(
						ref WordHelper.CAPTION_LABEL_TABLE, ref tableCaption,
                        ref VBAHelper.OPT_MISSING, ref WordHelper.CAPTION_POSITION_ABOVE, ref VBAHelper.OPT_MISSING);
				}

				// Mark the added single space and replace it with a paragraph mark.
				wrkRng.Start = wrkRng.End - 1;
				wrkRng.InsertParagraph();

				// Collapse the range to the end of the paragraph mark so that the table can be added
				// after it.
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}

			// Insert the table. Note, that the table is inserted starting at but after the range.
			// So viewRng isn't increased.

			Word.Table tbl = wdDoc_.Tables.Add(
				wrkRng, rows, cols,
				ref WordHelper.WORD8_TABLE_BEHAVIOR, ref VBAHelper.OPT_MISSING);

			// Reinstate auto caption for Word tables.
			ac.AutoInsert = oldState;


			// Autofit and table sizing are problematic, use the monkey
			tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);

			tbl.Borders.Enable = VBAHelper.iTRUE;
			tbl.TopPadding = 0;
			tbl.BottomPadding = 0;
			tbl.LeftPadding = 0;
			tbl.RightPadding = 0;
			tbl.Spacing = 0;

			tbl.Rows.LeftIndent = tbl.Application.InchesToPoints(0f);

			tbl.Rows.AllowBreakAcrossPages = VBAHelper.iFALSE;

			tbl.Borders.OutsideLineWidth = MacroBaseUtilities.LINE_WIDTH_NORMAL;
			tbl.Borders.InsideLineWidth = MacroBaseUtilities.LINE_WIDTH_NORMAL;

			wrkRng.End = tbl.Range.End;
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			wdDoc_.UndoClear();

			return tbl;
		}

		private Word.Table createFootnoteTable(Word.Range viewRng, int numFootnotes) 
		{
			// Turn off auto caption for Word tables.
			Word.AutoCaption ac =
				wdApp_.AutoCaptions.get_Item(ref WordHelper.AUTO_CAPTION_WORD_TABLE);
			bool oldState = ac.AutoInsert;
			ac.AutoInsert = false;

			Word.Range wrkRng = viewRng.Duplicate;

			// Collapse the range to the end of the paragraph mark so that the table can be added
			// after it.
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			wrkRng.InsertParagraphAfter();
			object oStyle = tspdDoc_.getStyleHelper().setNamedStyle(PurdueUtil.PFIZER_STYLE_TABLETEXT_10, wrkRng);
			//wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			//let the table eat the paragraph mark er created

			int nbrCols = 2;
			int nbrRows = numFootnotes;

			// Insert the table. Note, that the table is inserted starting at but after the range.
			// So viewRng isn't increased.
			Word.Table tbl =
				wdDoc_.Tables.Add(
				wrkRng, nbrRows, nbrCols,
				ref WordHelper.WORD8_TABLE_BEHAVIOR, ref VBAHelper.OPT_MISSING);
			

			oStyle = tspdDoc_.getStyleHelper().setNamedStyle(PurdueUtil.PFIZER_STYLE_TABLETEXT_10, tbl.Range);

			// Reinstate auto caption for Word tables.
			ac.AutoInsert = oldState;

			tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);

			tbl.Borders.Enable = VBAHelper.iFALSE;
			tbl.LeftPadding = 0;
			tbl.RightPadding = 0;
			tbl.Spacing = 0;
			tbl.Rows.LeftIndent = tbl.Application.InchesToPoints(0f);

			tbl.Rows.AllowBreakAcrossPages = VBAHelper.iFALSE;

			// Increase viewRng to include the table.
			viewRng.End = tbl.Range.End;

			wdDoc_.UndoClear();

			return tbl;
		}

		private class FNComparer : IComparer
		{
			public int Compare(object x, object y)
			{
				FootNoteWrapper fn1 = x as FootNoteWrapper;
				FootNoteWrapper fn2 = y as FootNoteWrapper;

				return fn1.footNoteNumber.CompareTo(fn2.footNoteNumber);
			}
		}

		private void fillFootnoteTable(Word.Table tbl, Hashtable footNotes) 
		{
			Word.Range wrk;
			FNComparer fnComparer = new FNComparer();

			int i = 0;
			ArrayList al = new ArrayList(footNotes.Values);
			al.Sort(fnComparer);

			foreach (FootNoteWrapper fnw in al)
			{
				Word.Row row = tbl.Rows[i + 1];
				// Number column
				wrk = row.Cells[1].Range;
				wrk.Paragraphs.KeepTogether = VBAHelper.iTRUE;
				wrk.Paragraphs.KeepWithNext = VBAHelper.iTRUE;
				wrk.End--;
				
				string s = fnw.footNoteNumberString;
				wrk.Text = s;

				// Footnote text
				wrk = row.Cells[2].Range;
				wrk.Paragraphs.KeepTogether = VBAHelper.iTRUE;
				wrk.Paragraphs.KeepWithNext = VBAHelper.iTRUE;
				wrk.End--;

				wrk.Text = fnw.footNote.getFootNoteText();

				wdDoc_.UndoClear();

				i++;
			}

			SOATableFormat tblFmt = tspdDoc_.getSOATblFormat(_currentSOA.getObjID());

			Word.Font targetFont;
			if (tblFmt == null)
			{
				// create a fake one to run thru defaults...
				TrialDocument.SOATableFormatCV cv = new TrialDocument.SOATableFormatCV();
				tblFmt = cv.newSOATableFormat();
			}

			targetFont = tbl.Range.Font;
			setTableBodyFont(ref targetFont, tblFmt);

			// Turn off autofit now
			// tbl.AllowAutoFit = false;

			// Make first col have a good size
			foreach (Word.Row rw in tbl.Rows)
			{
				try 
				{
					Word.Cell cl = rw.Cells[1];
					cl.SetWidth(tbl.Application.InchesToPoints(0.25f), Word.WdRulerStyle.wdAdjustProportional);
				}
				catch (Exception ex) 
				{
					Log.exception(ex, "Error setting column 1 width");
				}
			}

			wdDoc_.UndoClear();
		}

		public override void postProcess()
		{
			// Clean up memory
			_currentSOA = null;
			_currentArm = ArmRule.ALL_ARMS;
			_taskList.Clear();
			_invalidTC.Clear();
			_treatmentPeriods.Clear();
			_footNoter = null;
		}
	}
}
