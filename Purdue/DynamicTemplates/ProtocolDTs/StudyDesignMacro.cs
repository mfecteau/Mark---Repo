using System;
using System.Collections;
using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;
using MSXML2;

using TspdCfg.FastTrack.DynTmplts;

using Word = Microsoft.Office.Interop.Word;

namespace VersionControl 
{
	internal sealed class StudyDesignMacro
	{
		private static readonly string header_ = @"$Header: StudyDesignMacro.cs, 1, 18-Aug-09 12:05:48, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for StudyDesignMacro.
	/// </summary>
	public class StudyDesignMacro : AbstractMacroImpl
	{
		public class PeriodAndVisit
		{
			public Period per = null;
			public ProtocolEvent fv = null;
			public ProtocolEvent lv = null;
		}

		SOA _currentSOA = null;
		long _currentArm = ArmRule.ALL_ARMS;
		ArrayList _periodVisitList = new ArrayList();
		

		public new static bool canRun(BaseProtocolObject bpo)
		{
			return true;
		}

		public StudyDesignMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region StudyDesignMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd StudyDesign (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.StudyDesignMacro.StudyDesign,ProtocolDTs.dll" elementLabel="Study Design" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="Creates text for Study Design - Overview of Study Design Section." shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
#endif
			try 
			{
				mp.pba_.setOperation("Study Design Macro", "Generating information...");
				
				StudyDesignMacro macro = null;
				macro = new StudyDesignMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Study Design Macro"); 
				mp.inoutRng_.Text = "Study Design Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		#endregion

		#endregion

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

			displayPeriods(wrkRng);

			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}

		private void displayPeriods(Word.Range wrkRng) 
		{
			pba_.updateProgress(2.0);

			if (_periodVisitList.Count == 0)
			{
				wrkRng.InsertAfter("There are no periods defined in this schedule.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

				return;
			}

			//wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
			//wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			WordListHelper.ListTemplate wlt = WordListHelper.getNumberedListTemplate(wdApp_);

			foreach (PeriodAndVisit pv in _periodVisitList) 
			{
				pba_.updateProgress(2.0);

				wlt.BeginListItem(ref wrkRng);

				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
					pv.per, Period.BRIEF_DESCRIPTION, wrkRng, macroEntry_);

				if (!MacroBaseUtilities.isEmpty(pv.per.getFullDescription())) 
				{
					wrkRng.InsertAfter(" (");

					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
						pv.per, Period.FULL_DESCRIPTION, wrkRng, macroEntry_);

					wrkRng.InsertAfter(")");
				}

				wrkRng.InsertAfter(": ");

				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
					pv.per, Period.DURATION, wrkRng, macroEntry_);

				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
					pv.per, Period.DURATION_UNIT, wrkRng, macroEntry_);


				if (pv.fv == null || pv.lv == null) 
				{
					wrkRng.InsertAfter(" (No visits defined)");
				}
				else 
				{
					wrkRng.InsertAfter(" ");

					// Test
					if (false) 
					{
						wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_,pv.fv, ProtocolEvent.LABEL, wrkRng, macroEntry_);
						wrkRng.InsertAfter(" ");

						wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, pv.lv, ProtocolEvent.LABEL, wrkRng, macroEntry_);
						wrkRng.InsertAfter(" ");
					}
					// Test

					wrkRng.InsertAfter("(Days ");

					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
						pv.fv, ProtocolEvent.EFFECTIVE_TIME, wrkRng, macroEntry_);

					wrkRng.InsertAfter("to ");
					
					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
						pv.lv, ProtocolEvent.EFFECTIVE_TIME, wrkRng, macroEntry_);

					wrkRng.InsertAfter(")");
				}

				wlt.EndListItem(ref wrkRng);

				wdDoc_.UndoClear();
			}

			// wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);

			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			wdDoc_.UndoClear();
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

			if (_currentSOA != null) 
			{
				ArrayList orderedTopLevelEvents = new ArrayList();
				_currentSOA.getTopLevelActivityList(_currentArm, null, orderedTopLevelEvents);
				foreach (EventScheduleBase obj in orderedTopLevelEvents)
				{
					Period per = obj as Period;
					if (per == null) 
					{
						continue;
					}

					pba_.updateProgress(2.0);

					PeriodAndVisit pv = new PeriodAndVisit();
					pv.per = per;

					ArrayList visits = PurdueUtil.getVisits(_currentSOA, _currentArm, per, EventType.EventSubType.Scheduled);

					if (visits.Count != 0) 
					{
						pv.fv = visits[0] as ProtocolEvent;
						pv.lv = visits[visits.Count - 1] as ProtocolEvent;
					}

					_periodVisitList.Add(pv);
				}


				//if (pe.getEventType().getSubtype() == EventType.EventSubType.Scheduled) 
			}
		}

		public override void postProcess()
		{
			// Clean up memory
			_currentSOA = null;
			_periodVisitList.Clear();
			_currentArm = ArmRule.ALL_ARMS;
		}
	}
}
