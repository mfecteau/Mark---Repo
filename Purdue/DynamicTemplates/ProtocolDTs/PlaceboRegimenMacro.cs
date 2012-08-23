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
	internal sealed class PlaceboRegimenMacro
	{
		private static readonly string header_ = @"$Header: PlaceboRegimenMacro.cs, 1, 18-Aug-09 12:05:08, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for PlaceboRegimenMacro.
	/// </summary>
	public class PlaceboRegimenMacro : AbstractMacroImpl
	{
		SOA _currentSOA = null;
		long _currentArm = ArmRule.ALL_ARMS;
		IList _placeboList = null;
		public static readonly string CTMROLE_PLACEBO = "placebo";

		public PlaceboRegimenMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region PlaceboRegimen

		public static MacroExecutor.MacroRetCd PlaceboRegimen (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.PlaceboRegimenMacro.PlaceboRegimen,ProtocolDTs.dll" elementLabel="Placebo Dosing Regiment" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.test Article" autogenerates="true" toolTip="Dosing Regiment for the placebo test article" shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
#endif
			try 
			{
				mp.pba_.setOperation("PlaceboRegimen Macro", "Generating information...");
				mp.pba_.updateProgress(1.0);
				
				PlaceboRegimenMacro macro = null;
				macro = new PlaceboRegimenMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in PlaceboRegimen Macro"); 
				mp.inoutRng_.Text = "PlaceboRegimen Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		#endregion

		#endregion

		// If this is a macro based on a fly out menu, check if valid
		public static new bool canRun(BaseProtocolObject bpo)
		{
			SOA soa = bpo as SOA;
			if (soa == null)
			{
				return false;
			}

			// Example of further restriction
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


			_placeboList = PfizerUtil.findCTMsByType(bom_, CTMROLE_PLACEBO);
			pba_.updateProgress(10.0);
		}

		private bool outputPlaceboTask(Word.Range wrkRng, ClinicalTrialMaterial placebo, DosingTask task)
		{
			long studyDuration = 0;
			string durationUnit = null;
			bool isBadTime = true;
			bool isBadTimeUnit = true;
			bool noError = true;
					
			//loop thru the task visits break at the first sign of trouble
			IEnumerator ie = _currentSOA.getTaskVisitForTaskID(task.getObjID());
			bool found = false;
			while (ie.MoveNext()) 
			{
				found = true;
				pba_.updateProgress(2.0);
				IXMLDOMNode node = (IXMLDOMNode)ie.Current;
				TaskVisit tv = new TaskVisit(node);
						
				long vID = tv.getAssociatedVisitID();
						
				ProtocolEvent ev = _currentSOA.getProtocolEventByID(vID);
				PfizerUtil.addTimeUnit(ref studyDuration, ev.getDuration(), out isBadTime);

				if(isBadTime)
				{
					Period p = _currentSOA.getGrandPeriodOfScheduleEvent(ev);
					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, task, Task.BRIEF_DESCRIPTION, wrkRng, macroEntry_);					
					wrkRng.End = MacroBaseUtilities.putAfterElemRef(": a duration has not been defined for the event", tspdDoc_, wrkRng);
					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, p, Period.BRIEF_DESCRIPTION, wrkRng, macroEntry_);
					wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);
					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, ev, ProtocolEvent.BRIEF_DESCRIPTION, wrkRng, macroEntry_);
					wrkRng.End = MacroBaseUtilities.putAfterElemRef(".", tspdDoc_, wrkRng);
					wrkRng.InsertParagraphAfter();
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

					noError = false;
				}
//				if(isBadTimeUnit)
//				{
//					Period p = _currentSOA.getGrandPeriodOfScheduleEvent(ev);
//					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, task, Task.BRIEF_DESCRIPTION, wrkRng, macroEntry_);
//					wrkRng.End = MacroBaseUtilities.putAfterElemRef(": a duration unit has not been defined for the event", tspdDoc_, wrkRng);
//					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, p, Period.BRIEF_DESCRIPTION, wrkRng, macroEntry_);
//					wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);
//					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, ev, ProtocolEvent.BRIEF_DESCRIPTION, wrkRng, macroEntry_);
//					wrkRng.End = MacroBaseUtilities.putAfterElemRef(".", tspdDoc_, wrkRng);
//					wrkRng.InsertParagraphAfter();
//					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
//							
//					noError = false;
//				}
				if(MacroBaseUtilities.isEmpty(durationUnit))
				{
					durationUnit = ev.getDurationTimeUnit();
				}
			}
			if(found == false)
			{
				wrkRng.InsertAfter("A dosing event for the ");
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, task, Task.BRIEF_DESCRIPTION, wrkRng, macroEntry_);
				wrkRng.End = MacroBaseUtilities.putAfterElemRef(" does not exist.", tspdDoc_, wrkRng);
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				noError = false;
			}
			if(noError == true)
			{
				string s = PfizerUtil.getDisplayTime(studyDuration, durationUnit);							
						
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, placebo, ClinicalTrialMaterial.CTMATERIAL_NAME, wrkRng, macroEntry_);
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, placebo, ClinicalTrialMaterial.FORMULATION, wrkRng, macroEntry_);
				wrkRng.End = MacroBaseUtilities.putAfterElemRef(" REGIMEN for" + s + ".", tspdDoc_, wrkRng);
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}

			return found;
		}

		public override void display()
		{
			Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;

			pba_.updateProgress(1.0);

			string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);
			int count = _placeboList.Count;

			if (_currentSOA == null)
			{
				pba_.updateProgress(70.0);
				wrkRng.InsertAfter("This schedule that this macro refers to was removed, delete this macro.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				
			}
			else if(count == 0)
			{
				pba_.updateProgress(70.0);
				wrkRng.InsertAfter("A " + CTMROLE_PLACEBO + " has not been defined.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
			else
			{
				IList dosingList = PfizerUtil.getDosingTaskskByTAType(_currentSOA, CTMROLE_PLACEBO, bom_, icdSchemaMgr_);
				int taskCount = dosingList.Count;	
				
				for (int i = 0; i < count; i++)
				{ 
					if(i > 0)
					{
						wrkRng.InsertParagraphAfter();
					}
					pba_.updateProgress(20.0);
					ClinicalTrialMaterial placebo = (ClinicalTrialMaterial)_placeboList[i];					

					int j = 0;
					
					for(; j < taskCount; j++)
					{
						DosingTask dt = (DosingTask)dosingList[j];
						if(dt.getctMaterialID() == placebo.getObjID())
						{	
							outputPlaceboTask(wrkRng, placebo, dt);
							//whatever the outcome, break.
							break;
						}
					}
					if(j == taskCount)
					{
						wrkRng.InsertAfter("A dosing event for the ");
						wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, placebo, ClinicalTrialMaterial.CTMATERIAL_NAME, wrkRng, macroEntry_);
						wrkRng.End = MacroBaseUtilities.putAfterElemRef(" does not exists.", tspdDoc_, wrkRng);
						wrkRng.InsertParagraphAfter();
						wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
					}
				}
			}

			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}

		public override void postProcess()
		{
			// Clean up memory
			_currentSOA = null;
			_currentArm = ArmRule.ALL_ARMS;
			_placeboList.Clear();
			_placeboList = null;
		}
	}
}
