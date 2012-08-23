using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.MacroBase;
using Tspd.MacroBase.Table;
using Tspd.MacroBase.BaseImpl;
using Tspd.Macros;
using Tspd.Tspddoc;
using Tspd.Utilities;
using Word = Microsoft.Office.Interop.Word;

namespace TspdCfg.Purdue.DynTmplts.Table
{
	class PfizerPeriodRowCellRenderer : DefSOAPeriodRowHeaderCell
	{
		public override string getDisplayName()
		{
			return "Protocol Activity";
		}

		public override void display(TspdDocument doc, Word.Cell currCell, object obj)
		{
			Period period = obj as Period;
			if (period != null)
			{
				if (period is PeriodDivider)
				{
					MacroBaseUtilities.putElemRefInCell(doc, currCell, period, Period.BRIEF_DESCRIPTION, true);
				}
				else
				{
					MacroBaseUtilities.putElemRefInCell(doc, currCell, period, Period.BRIEF_DESCRIPTION, true);
				}
			}
		}
		
	}

	class PfizerSubPeriodRowCellRenderer : DefSOASubPeriodRowHeaderCell
	{
		public override string getDisplayName()
		{
			return "";
		}

		public override void display(TspdDocument doc, Word.Cell currCell, object obj)
		{
			SubPeriodWrapper spw = obj as SubPeriodWrapper;
			if (spw != null && spw.isRealSubPeriod())
			{
				Period subPeriod = spw.getWrapped() as Period;
				if (subPeriod != null)
				{
					MacroBaseUtilities.putElemRefInCell(doc, currCell, subPeriod, Period.BRIEF_DESCRIPTION, true);
				}
			}
		}
	}

	class PfizerSOAVisitStudyDaysRowHeaderCell : DefSOAVisitStudyDaysRowHeaderCell 
	{
		public override string getDisplayName()
		{
			return "";
		}

		public override void addFootNotes(IFootnoter footnoter, Word.Cell currCell, object dataObj)
		{
			VisitStudyDays studyDays = dataObj as VisitStudyDays;
			if (studyDays != null)
			{
				ProtocolEvent visit = studyDays.getWrapped() as ProtocolEvent;
				if (visit != null)
				{
					MacroBaseUtilities.addFootNotesToCell(footnoter, currCell, visit);
				}
			}
		}
	}

	class PfizerVisitRowCellRenderer : DefSOAVisitRowHeaderCell 
	{
		  public override string getDisplayName()
		{
			
			return "";
		}

		public override void display(TspdDocument doc, Word.Cell currCell, object obj)
		{
			ProtocolEvent visit = obj as ProtocolEvent;
			if(visit != null)
			{
				MacroBaseUtilities.putElemRefInCell(doc,currCell, visit, ProtocolEvent.BRIEF_DESCRIPTION, true);

			}
		}

		
//		public override void addFootNotes(IFootnoter footnoter, Word.Cell currCell, object dataObj)
//		{
//			// No footnote on Visits
//		}
	}

	class PfizerVisitWindowRowHeaderCell : DefSOAVisitWindowRowHeaderCell
	{
		public override string getDisplayName()
		{
			return "Window";
		}
	}

	class NoValueRenderer: DefNoValueSOATableCell
	{
		public NoValueRenderer(Word.WdColor cellColor) : base(cellColor) {}

		public override void display(TspdDocument doc, Word.Cell currCell, object obj)
		{
			//nada... don't do the expensive stuff frozen in core.
		}
	}

	class PfizerTaskColumnHeaderCell : DefSOAColumnHeaderCell
	{
	
		DefSOATableView _tableView = null;

		public PfizerTaskColumnHeaderCell(DefSOATableView tableView) 
		{
			_tableView = tableView;
		}

		public override void display(TspdDocument doc, Word.Cell currCell, object obj)
		{
			TaskWrapper tw = obj as TaskWrapper;
			if (tw != null)
			{
				Task task = tw.getWrapped() as Task;
				if (task != null)
				{
					MacroBaseUtilities.putElemRefInCell(doc, currCell, task, Task.DISPLAYNAME, true);

					SOA soa = _tableView.DataModel.getSOA();
					TaskDivider tdParent = soa.getParentTaskDivider(task);

					if (tdParent != null) 
					{
						Word.Range cr = currCell.Range.Duplicate;
						cr.Collapse(ref WordHelper.COLLAPSE_START);
						cr.InsertAfter("    ");
					}
				}
			}			
		}
	}
}
