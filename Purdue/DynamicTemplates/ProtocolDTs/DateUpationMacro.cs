using System;
using System.Collections;
using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;
using MSXML2;

using TspdCfg.SalesDemo.DynTmplts;

namespace VersionControl 
{
	internal sealed class DateUpdationMacro
	{
		private static readonly string header_ = @"$Header: DateUpationMacro.cs, 1, 18-Aug-09 12:03:35, Pinal Patel$";
	}
}

namespace TspdCfg.SalesDemo.DynTmplts
{
	/// <summary>
	/// Summary description for StudyDesignMacro.
	/// </summary>
	public class DateUpdationMacro : AbstractMacroImpl
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

		public DateUpdationMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region DateUpdationMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd DateUpdation (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.SalesDemo.DynTmplts.StudyDesignMacro.StudyDesign,ProtocolDTs.dll" elementLabel="Study Design" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="Creates text for Study Design - Overview of Study Design Section." shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
#endif
			try 
			{
				mp.pba_.setOperation("Study Design Macro", "Generating information...");
				
				DateUpdationMacro macro = null;
				macro = new DateUpdationMacro(mp);
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
			Word.Selection sel_ = null;

			TspdTrial trial = tspdDoc_.getTspdTrial();
			FTDateTime dtCreated = trial.getCreateDate();
				
			string sCreated = dtCreated.getDateTime().ToLongDateString();
			WordHelper.setVariableValue(tspdDoc_.getActiveWordDocument(),"tspd.trial.createdate",sCreated);

			Word.Document myDoc = tspdDoc_.getActiveWordDocument();
			pba_.updateProgress(1.0);

			

			float defSpacing = wrkRng.ParagraphFormat.SpaceAfter;
			wrkRng.ParagraphFormat.SpaceAfter = 6F;
		
			//sel_.Font.Name = HeaderFontName;
		    //sel_.Font.Size = TitleFontSize;
			//enterHeaderFooter(Word.WdSeekView.wdSeekCurrentPageHeader);	
			

			if (!(tspdDoc_.getActiveWordDocument().ActiveWindow.View.SplitSpecial == Word.WdSpecialPane.wdPaneNone)) 
			{
				tspdDoc_.getActiveWordDocument().ActiveWindow.Panes.Item(2).Close();
			}
			Word.View view = tspdDoc_.getActiveWordDocument().ActiveWindow.ActivePane.View;
			if(view.Type == Word.WdViewType.wdNormalView
				|| view.Type == Word.WdViewType.wdOutlineView 
				|| view.Type == Word.WdViewType.wdMasterView )
			{
				view.Type = Word.WdViewType.wdPrintView;
			}

			view.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader; 

			string crDate = WordHelper.getVariableValue(tspdDoc_.getActiveWordDocument(),"tspd.trial.createdate").Value;

			wrkRng.InsertAfter(crDate);		



				wrkRng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
			

			
			
			//	exitHeaderFooter();	
			tspdDoc_.getActiveWordDocument().ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
			tspdDoc_.getActiveWordDocument().ActiveWindow.DocumentMap = false;
			wrkRng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
			wrkRng.ParagraphFormat.SpaceAfter = defSpacing;
			


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
				
//			while (soaEnum.MoveNext())
//			{
//				pba_.updateProgress(2.0);
//
//				SOA soa = soaEnum.getCurrent();
//				if (soa.getElementPath().Equals(elementPath)) 
//				{
//					_currentSOA = soa;
//					break;
//				}
//			}

//			if (_currentSOA != null) 
//			{
//				ArrayList orderedTopLevelEvents = new ArrayList();
//				_currentSOA.getTopLevelActivityList(_currentArm, null, orderedTopLevelEvents);
//				foreach (EventScheduleBase obj in orderedTopLevelEvents)
//				{
//					Period per = obj as Period;
//					if (per == null) 
//					{
//						continue;
//					}
//
//					pba_.updateProgress(2.0);
//
//					PeriodAndVisit pv = new PeriodAndVisit();
//					pv.per = per;
//
//					ArrayList visits = PfizerUtil.getVisits(_currentSOA, _currentArm, per, EventType.EventSubType.Scheduled);
//
//					if (visits.Count != 0) 
//					{
//						pv.fv = visits[0] as ProtocolEvent;
//						pv.lv = visits[visits.Count - 1] as ProtocolEvent;
//					}
//
//					_periodVisitList.Add(pv);
//				}
//
//
//				//if (pe.getEventType().getSubtype() == EventType.EventSubType.Scheduled) 
//			}
		}

		public override void postProcess()
		{
			// Clean up memory
//			_currentSOA = null;
//			_periodVisitList.Clear();
//			_currentArm = ArmRule.ALL_ARMS;
		}
	}
}
