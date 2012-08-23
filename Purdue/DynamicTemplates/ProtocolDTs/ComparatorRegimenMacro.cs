using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;
using System.Xml;

using Word = Microsoft.Office.Interop.Word;

namespace VersionControl 
{
	internal sealed class ComparatorRegimenMacro
	{
		private static readonly string header_ = @"$Header: ComparatorRegimenMacro.cs, 1, 18-Aug-09 12:03:23, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for ComparatorRegimenMacro.
	/// </summary>
	public class ComparatorRegimenMacro : AbstractMacroImpl
	{
		SOA _currentSOA;
		long _currentArm;
		List<PurdueUtil.TreatmentComponentAndTestArticle> _placeboList;
        List<PurdueUtil.TreatmentComponentAndTestArticle> _comparatorList;

		public const string CTMROLE_COMPARATOR = "comparator";		
		public const string CTMROLE_PLACEBO = "placebo";

		public ComparatorRegimenMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
            _currentArm = ArmRule.ALL_ARMS;
        }
		
		#region Dynamic Tmplt Methods
		
		#region ComparatorRegimen

		public static MacroExecutor.MacroRetCd ComparatorRegimen (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.ComparatorRegimenMacro.ComparatorRegimen,ProtocolDTs.dll" elementLabel="Comparator Regimen" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Test Article" autogenerates="true" toolTip="Dosing Regimen for the comparator test article" shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
#endif
			try 
			{
				mp.pba_.setOperation("ComparatorRegimen Macro", "Generating information...");
				
				ComparatorRegimenMacro macro = null;
				macro = new ComparatorRegimenMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in ComparatorRegimen Macro"); 
				mp.inoutRng_.Text = "ComparatorRegimen Macro: " + e.Message;
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

			_placeboList = PurdueUtil.findTreatmentsByRole(bom_, CTMROLE_PLACEBO);
			pba_.updateProgress(5.0);

            _comparatorList = PurdueUtil.findTreatmentsByRole(bom_, CTMROLE_COMPARATOR);
			pba_.updateProgress(5.0);
		}

		private bool outputComparatorTask(Word.Range wrkRng, Treatment comparatorTreatment, Component comparatorComponent,
            Treatment placeboTreatment, Component placeboComponent, DosingTask task)
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
                XmlNode node = (XmlNode)ie.Current;
				TaskVisit tv = new TaskVisit(node);
						
				long vID = tv.getAssociatedVisitID();
						
				ProtocolEvent ev = _currentSOA.getProtocolEventByID(vID);
				PurdueUtil.addTimeUnit(ref studyDuration, ev.getDuration(), out isBadTime);

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
				if(isBadTimeUnit)
				{
					Period p = _currentSOA.getGrandPeriodOfScheduleEvent(ev);
					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, task, Task.BRIEF_DESCRIPTION, wrkRng, macroEntry_);
					wrkRng.End = MacroBaseUtilities.putAfterElemRef(": a duration unit has not been defined for the event", tspdDoc_, wrkRng);
					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, p, Period.BRIEF_DESCRIPTION, wrkRng, macroEntry_);
					wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);
					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, ev, ProtocolEvent.BRIEF_DESCRIPTION, wrkRng, macroEntry_);
					wrkRng.End = MacroBaseUtilities.putAfterElemRef(".", tspdDoc_, wrkRng);
					wrkRng.InsertParagraphAfter();
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
							
					noError = false;
				}
				if(MacroBaseUtilities.isEmpty(durationUnit))
				{
					durationUnit = ev.getDurationTimeUnit();
				}
			}
			if(found == false)
			{
				wrkRng.InsertAfter("A dosing event for ");
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, task, Task.BRIEF_DESCRIPTION, wrkRng, macroEntry_);
				wrkRng.End = MacroBaseUtilities.putAfterElemRef(" does not exist.", tspdDoc_, wrkRng);
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				noError = false;
			}
			
			if(noError == true)
			{
				string s = PurdueUtil.getDisplayTime(studyDuration, durationUnit);

                wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, comparatorTreatment, Treatment.NAME, wrkRng, macroEntry_);
                wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, comparatorComponent, Component.FORMULATION, wrkRng, macroEntry_);
				wrkRng.InsertAfter("(");
                wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, comparatorTreatment, Treatment.DOSE, wrkRng, macroEntry_);
				wrkRng.End = MacroBaseUtilities.putAfterElemRef(") on the last day of a", tspdDoc_, wrkRng);
                wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, placeboTreatment, Treatment.NAME, wrkRng, macroEntry_);
                wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, placeboComponent, Component.FORMULATION, wrkRng, macroEntry_);
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
			int comparatorCount = _comparatorList.Count;
			int placeboCount = _placeboList.Count;
				
			if (_currentSOA == null)
			{
				pba_.updateProgress(70.0);
				wrkRng.InsertAfter("This schedule that this macro refers to was removed, delete this macro.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				
			}
			else if(comparatorCount == 0)
			{
				pba_.updateProgress(70.0);
				wrkRng.InsertAfter("A " + CTMROLE_COMPARATOR + " has not been defined.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
			else
			{
				IList dosingList = PurdueUtil.getDosingTaskskByTAType(_currentSOA, CTMROLE_COMPARATOR, bom_, icdSchemaMgr_);
				int taskCount = dosingList.Count;

				for (int i = 0; i < comparatorCount; i++)
				{ 
					if(i > 0)
					{
						wrkRng.InsertParagraphAfter();
					}
					pba_.updateProgress(20.0);
					PurdueUtil.TreatmentComponentAndTestArticle comparator = _comparatorList[i];
					string cname = comparator.MatchingTreatment.Name;

                    PurdueUtil.TreatmentComponentAndTestArticle ctmPlacebo = null;
					for (int k = 0; k < placeboCount; k++)
					{
                        PurdueUtil.TreatmentComponentAndTestArticle placebo = _placeboList[k];
                        string pname = placebo.MatchingTreatment.Name;
						if(pname.StartsWith(cname))
						{
							ctmPlacebo = placebo;
							break;
						}
					}
					if(ctmPlacebo == null)
					{
						wrkRng.InsertAfter("A Placebo has not been defined for ");
						wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, comparator.MatchingTreatment, Treatment.NAME, wrkRng, macroEntry_);
						wrkRng.End = MacroBaseUtilities.putAfterElemRef(".", tspdDoc_, wrkRng);	
						wrkRng.InsertParagraphAfter();
						wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
					}
					else
					{
                        /* Dosing tasks do not exist in TSPD 3.1 - LAP
						int j = 0;					
						for(; j < taskCount; j++)
						{
							DosingTask dt = (DosingTask)dosingList[j];
							if(dt.getctMaterialID() == ctm.getObjID())
							{	
								outputComparatorTask(wrkRng, ctm, ctmPlacebo, dt);
								//whatever the outcome, break.
								break;
							}
						}
						if(j == taskCount)
						{
							wrkRng.InsertAfter("A dosing event for the ");
							wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, comparator.MatchingTreatment, Treatment.NAME, wrkRng, macroEntry_);
							wrkRng.End = MacroBaseUtilities.putAfterElemRef(" does not exists.", tspdDoc_, wrkRng);
							wrkRng.InsertParagraphAfter();
							wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
						}
                        */
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
			_comparatorList.Clear();
			_comparatorList = null;
		}
	}
}
