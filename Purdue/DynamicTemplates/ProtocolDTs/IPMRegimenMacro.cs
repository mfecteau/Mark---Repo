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
	internal sealed class IPMRegimenMacro
	{
		private static readonly string header_ = @"$Header: IPMRegimenMacro.cs, 1, 18-Aug-09 12:04:19, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for IPMRegimenMacro.
	/// </summary>
	public class IPMRegimenMacro : AbstractMacroImpl
	{
		
		SOA _currentSOA = null;
		long _currentArm = ArmRule.ALL_ARMS;
		IList _placeboList = null;
		IList _ipmList = null;
		public static readonly string CTMROLE_INVESTIGATIONAL_PRODUCT = "investigationalProduct";		
		public static readonly string CTMROLE_PLACEBO = "placebo";

		public IPMRegimenMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region IPMRegimen

		public static MacroExecutor.MacroRetCd IPMRegimen (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.IPMRegimenMacro.IPMRegimen,ProtocolDTs.dll" elementLabel="IPM Regimen" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Test Article" autogenerates="true" toolTip="Dosing Regimen for the investigational product" shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
#endif
			try 
			{
				mp.pba_.setOperation("IPMRegimen Macro", "Generating information...");
				
				IPMRegimenMacro macro = null;
				macro = new IPMRegimenMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in IPMRegimen Macro"); 
				mp.inoutRng_.Text = "IPMRegimen Macro: " + e.Message;
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


			_placeboList = new ArrayList();
			_ipmList = new ArrayList();
			CTMaterialEnumerator ctEnum = bom_.getCTMaterialEnumerator();
			while (ctEnum.MoveNext()) 
			{				
				pba_.updateProgress(2.0);

				ClinicalTrialMaterial ctm = ctEnum.getCurrent();
				string primaryRole = ctm.getPrimaryRole();
					
				if (!MacroBaseUtilities.isEmpty(primaryRole))
				{
					if(primaryRole.Equals("placebo"))
					{
						_placeboList.Add(ctm);
					}
					else if(primaryRole.Equals("investigationalProduct"))
					{
						//do not enter parents....
						BusinessObjectFactory bof = ctm.getChildrenLikeParent();
						if(bof.getList().Count == 0)
						{
							_ipmList.Add(ctm);
						}
					}
				}
			}
		}

		private bool nameOutput1(Word.Range wrkRng, ClinicalTrialMaterial ipm)
		{
			ClinicalTrialMaterial parent = (ClinicalTrialMaterial)ipm.getParentLikeChild();
			if(MacroBaseUtilities.isEmpty(parent))
			{
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, ipm, ClinicalTrialMaterial.CTMATERIAL_NAME, wrkRng, macroEntry_);
				return false;
			}
			else
			{
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, parent, ClinicalTrialMaterial.CTMATERIAL_NAME, wrkRng, macroEntry_);
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, ipm, ClinicalTrialMaterial.CTMATERIAL_NAME, wrkRng, macroEntry_);
				return true;
			}
		}

		private bool nameOutput2(Word.Range wrkRng, ClinicalTrialMaterial ipm)
		{
			ClinicalTrialMaterial parent = (ClinicalTrialMaterial)ipm.getParentLikeChild();
			if(MacroBaseUtilities.isEmpty(parent))
			{
				wrkRng.InsertAfter(" REGIMEN");
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, ipm, ClinicalTrialMaterial.CTMATERIAL_NAME, wrkRng, macroEntry_);				
				return false;
			}
			else
			{
				wrkRng.InsertAfter("(");
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, ipm, ClinicalTrialMaterial.CTMATERIAL_NAME, wrkRng, macroEntry_);
				wrkRng.End = MacroBaseUtilities.putAfterElemRef(") REGIMEN", tspdDoc_, wrkRng);
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, parent, ClinicalTrialMaterial.CTMATERIAL_NAME, wrkRng, macroEntry_);				
				return true;
			}
		}

		private bool outputIPTask(Word.Range wrkRng, ClinicalTrialMaterial ipm, ClinicalTrialMaterial placebo, DosingTask task)
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
				string s = PfizerUtil.getDisplayTime(studyDuration, durationUnit);							
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, ipm, ClinicalTrialMaterial.DOSE, wrkRng, macroEntry_);
				nameOutput2(wrkRng, ipm);
				wrkRng.End = MacroBaseUtilities.putAfterElemRef(" plus a", tspdDoc_, wrkRng);
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, placebo, ClinicalTrialMaterial.CTMATERIAL_NAME, wrkRng, macroEntry_);
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, placebo, ClinicalTrialMaterial.FORMULATION, wrkRng, macroEntry_);
				wrkRng.End = MacroBaseUtilities.putAfterElemRef(" for" + s + ".", tspdDoc_, wrkRng);												
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
			int ipmCount = _ipmList.Count;
			int placeboCount = _placeboList.Count;
				
			if (_currentSOA == null)
			{
				pba_.updateProgress(70.0);
				wrkRng.InsertAfter("This schedule that this macro refers to was removed, delete this macro.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				
			}
			else if(ipmCount == 0)
			{
				pba_.updateProgress(70.0);
				wrkRng.InsertAfter("A " + CTMROLE_INVESTIGATIONAL_PRODUCT + " has not been defined.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
			else
			{
				IList dosingList = PfizerUtil.getDosingTaskskByTAType(_currentSOA, CTMROLE_INVESTIGATIONAL_PRODUCT, bom_, icdSchemaMgr_);
				int taskCount = dosingList.Count;

				for (int i = 0; i < ipmCount; i++)
				{ 
					if(i > 0)
					{
						wrkRng.InsertParagraphAfter();
					}
					
					pba_.updateProgress(20.0);
					ClinicalTrialMaterial ctm = (ClinicalTrialMaterial)_ipmList[i];	
					if(ctm.getChildrenLikeParent().getList().Count > 0)
					{
						continue;
						//handle children individually
					}
					
					string ctname = ctm.getMaterialName();
					ClinicalTrialMaterial ip2Compare = (ClinicalTrialMaterial)ctm.getParentLikeChild();
					if(!MacroBaseUtilities.isEmpty(ip2Compare))
					{
						ctname = ip2Compare.getMaterialName();;
					}

				    ClinicalTrialMaterial ctmPlacebo = null;
					for (int k = 0; k < placeboCount; k++)
					{
						ctmPlacebo = null;
						ClinicalTrialMaterial placebo = (ClinicalTrialMaterial)_placeboList[k];
						string pname = placebo.getMaterialName();
						if(pname.StartsWith(ctname))
						{
							ctmPlacebo = placebo;
							break;
						}
					}
					if(ctmPlacebo == null)
					{
						wrkRng.InsertAfter("A Placebo has not been defined for ");
						nameOutput1(wrkRng, ctm);
						wrkRng.End = MacroBaseUtilities.putAfterElemRef(".", tspdDoc_, wrkRng);	
						wrkRng.InsertParagraphAfter();
						wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
					}
					else
					{

						int j = 0;					
						for(; j < taskCount; j++)
						{
							DosingTask dt = (DosingTask)dosingList[j];
							if(dt.getctMaterialID() == ctm.getObjID())
							{								
								outputIPTask(wrkRng, ctm, ctmPlacebo, dt);
								break;
							}
						}
						if(j == taskCount)
						{
							wrkRng.InsertAfter("A dosing event for the ");
							nameOutput1(wrkRng, ctm);
							wrkRng.End = MacroBaseUtilities.putAfterElemRef(" does not exists.", tspdDoc_, wrkRng);
							wrkRng.InsertParagraphAfter();
							wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
						}
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
			_ipmList.Clear();
			_ipmList = null;
		}
	}
}
