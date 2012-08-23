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
	internal sealed class ScheduleOfProceduresMacro
	{
		private static readonly string header_ = @"$Header: ScheduleOfProceduresMacro.cs, 1, 18-Aug-09 12:05:45, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for ScheduleOfProceduresMacro.
	/// </summary>
	public class ScheduleOfProceduresMacro : AbstractMacroImpl
	{
		public class TVCell
		{
			public ArrayList tvList = new ArrayList();
			public string spanLabel = "";
		}

		SOA _currentSOA = null;
		long _currentArm = ArmRule.ALL_ARMS;
		ArrayList _periodVisitList = new ArrayList();
		ArrayList _taskList = new ArrayList();
		ArrayList _invalidTC = new ArrayList();
		IFootnoter _footNoter = null;

		public static readonly string CTMROLE_INVESTIGATIONAL_PRODUCT = "investigationalProduct";

		public ScheduleOfProceduresMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region ScheduleOfProcedures
		public static MacroExecutor.MacroRetCd ScheduleOfProcedures (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.ScheduleOfProceduresMacro.ScheduleOfProcedures,ProtocolDTs.dll" elementLabel="Schedule of Procedures" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="Prints table of Schedule of Procedures for a given Schedule of Assessments dosing days." shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
#endif
			try 
			{
				mp.pba_.setOperation("Schedule Of Procedures Macro", "Generating information...");
				
				ScheduleOfProceduresMacro macro = null;
				macro = new ScheduleOfProceduresMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Schedule Of Procedures Macro"); 
				mp.inoutRng_.Text = "Schedule Of Procedures Macro: " + e.Message;
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


			// Collect the ordered visit list
			ArrayList orderedTopLevelEvents = new ArrayList();
			_currentSOA.getTopLevelActivityList(_currentArm, null, orderedTopLevelEvents);
			
			ArrayList pvList = PurdueUtil.getPeriodVisitList(_currentSOA , _currentArm, orderedTopLevelEvents);

			// Collect the ordered tasks
			ArrayList includedTasks = new ArrayList();
			TaskEnumerator taskEnum = _currentSOA.getTaskEnumerator();
			while (taskEnum.MoveNext()) 
			{
				Task task = taskEnum.getCurrent();

				_taskList.Add(task);

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

			// Go through the pv list and find all task visits for each visit, include if one of the tasks is the 
			// in the includedTasks list
			foreach (PurdueUtil.PeriodAndVisit pv in pvList) 
			{
				ArrayList tcList = PurdueUtil.getTimeChunks(_currentSOA, pv.visit);

				// Accumulation of time chunks in this pv
				ArrayList usedTcList = new ArrayList();

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
						_invalidTC.Add("Period: " + pv.per.getBriefDescription() + ", Visit: " +
							pv.visit.getBriefDescription() + ", Time: " + tc.Label + ": " + serr);
						continue;
					}

					// Add time chunks
					usedTcList.Add(new PurdueUtil.TCWrapper(tc, startTime, endTime, unit));

					// If the included task is used by this time chunk
					// point the pv.tcList at all timechunks in this pv
					foreach (Task task in includedTasks) 
					{
						ArrayList tvList = tc.getByTaskID(task.getObjID());
						if (tvList.Count != 0) 
						{
							if (!_periodVisitList.Contains(pv)) 
							{
								pv.tcList = usedTcList;
								_periodVisitList.Add(pv);
							}
						}
					}
				}
			}
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
				if (tvList.Count != 0) 
				{
					return true;
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

			if (_periodVisitList.Count == 0) 
			{
				pba_.updateProgress(70.0);

				wrkRng.InsertAfter("There are no visits with dosing tasks for the investigational product defined.");
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

			foreach (PurdueUtil.PeriodAndVisit pv in _periodVisitList) 
			{
				displayVisitTables(pv, ref wrkRng);
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

		private void displayVisitTables(PurdueUtil.PeriodAndVisit pv, ref Word.Range wrkRng) 
		{
			#region more pre processing

			ArrayList errorList = new ArrayList();
			ArrayList usedTasks = new ArrayList();

			foreach (Task t in _taskList) 
			{
				if (isTaskUsed(t, pv))
				{
					usedTasks.Add(t);
				}
			}

			if (usedTasks.Count == 0) 
			{
				return;
			}

			if (pv.tcList.Count == 0)
			{
				return;
			}

			// Split out spanned times
			ArrayList tcList = new ArrayList();
			ArrayList tcSpannedList = new ArrayList();

			foreach (PurdueUtil.TCWrapper tcw in pv.tcList) 
			{
				if (tcw.isSpan()) 
				{
					// spans live in columns time-aligned with their start
					// e.g.  chunk 8-12 lives in the 8 column
					tcSpannedList.Add(tcw);
				}
				else
				{
					// These are the time chunks that will become columns in the table
					tcList.Add(tcw);
				}
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
					// add a column to align with the start of a span if the column does not exist
					addList.Add(new PurdueUtil.TCWrapper(null, tcwSpan.Start, "", tcwSpan.Unit));
				}

				if (!foundEnd) 
				{
					// add a column to align with the end of a span if the column does not exist
					// but Don't add endpoint label of 24 hrs
					if (!(tcwSpan.End.Equals("24") && tcwSpan.Unit.Equals(PurdueUtil.TimeUnit.sHOURS)))
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
						tvCell.tvList.AddRange(a);
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
					if (tvList.Count != 1) 
					{
						addErrorMessage(sErrTask, errorList);

						string serr = tcwSpan.Label + " overlap with: " + tcwSpan.Label;
						addErrorMessage(serr, errorList);
					}

					bool startEmpty = true;
					bool foundStart = false;
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

							// If the cell already has a span, error
							if (!MacroBaseUtilities.isEmpty(tvCell.spanLabel)) 
							{
								addErrorMessage(sErrTask, errorList);

								string serr = tcwSpan.Label + " overlap with: " + tvCell.spanLabel;
								addErrorMessage(serr, errorList);
								break;
							}

							// not sure what;s happening here
							if (tvCell.tvList.Count == 0) 
							{
								emptyCount++;
							}
							else
							{
								startEmpty = false;
							}

							// copy the the span's label into the cell's label
							tvCell.spanLabel = tcwSpan.Label;
						}


						// end
						if (tcwSpan.getEndMinute() == tcw.getStartMinute()) 
						{
							break;
						}
						// if the span has begun, but we are one or more cells along
						if (foundStart && tvCell != startCell) 
						{
							// If the start cell is empty and we have an X, error
							if (startEmpty && tvCell.tvList.Count != 0) 
							{
								addErrorMessage(sErrTask, errorList);

								string serr = tcwSpan.Label + " overlap with: " + tcw.Label;
								addErrorMessage(serr, errorList);
								break;
							}

							// If the cell already has a span, error
							if (!MacroBaseUtilities.isEmpty(tvCell.spanLabel)) 
							{
								addErrorMessage(sErrTask, errorList);

								string serr = tcwSpan.Label + " overlap with: " + tvCell.spanLabel;
								addErrorMessage(serr, errorList);
								break;
							}
							
							if (tvCell.tvList.Count == 0) 
							{
								emptyCount++;
							}
							tvCell.spanLabel = tcwSpan.Label;
						}
					}

					if (foundStart && emptyCount == 0) 
					{
						addErrorMessage(sErrTask, errorList);

						string serr = tcwSpan.Label + " no open cell";
						addErrorMessage(serr, errorList);
					}
				}
			}

			if (errorList.Count != 0) 
			{
				wrkRng.InsertAfter("Period: " + pv.per.getBriefDescription() + 
									", Visit: " + pv.visit.getBriefDescription() +
									": Invalid time point: ");

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
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

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
			row = tbl.Rows[PERIOD_VISIT_ROW];
			cell = row.Cells[1];
			MacroBaseUtilities.putElemRefInCell(tspdDoc_, cell, pv.per, Period.BRIEF_DESCRIPTION, true, macroEntry_);
			
			Word.Range clRng = cell.Range.Duplicate;
			clRng.End--;
			clRng.Collapse(ref WordHelper.COLLAPSE_END);
			MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, clRng);

			MacroBaseUtilities.putElemRefInCell(tspdDoc_, cell, pv.visit, ProtocolEvent.BRIEF_DESCRIPTION, true, macroEntry_);
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
					if (rw.Cells.Count > 1) 
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
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

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
			tbl.AllowAutoFit = false;

			// Make first col have a good size
			foreach (Word.Row rw in tbl.Rows)
			{
				try 
				{
					Word.Cell cl = rw.Cells[1];
					cl.SetWidth(tbl.Application.InchesToPoints(0.5f), Word.WdRulerStyle.wdAdjustProportional);
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
			_periodVisitList.Clear();
			_taskList.Clear();
			_invalidTC.Clear();
			_footNoter = null;
		}
	}
}
