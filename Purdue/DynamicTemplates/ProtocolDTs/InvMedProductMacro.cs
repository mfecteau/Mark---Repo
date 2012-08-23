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
	internal sealed class InvMedProductMacro
	{
		private static readonly string header_ = @"$Header: InvMedProductMacro.cs, 1, 18-Aug-09 12:04:39, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for InvMedProductMacro.
	/// </summary>
	public class InvMedProductMacro : AbstractMacroImpl
	{
		SOA _currentSOA = null;
		long _currentArm = ArmRule.ALL_ARMS;
		ArrayList _orderedVisits = new ArrayList();
		ClinicalTrialMaterial _ta1 = null;
		ClinicalTrialMaterial _ta2 = null;

		Hashtable _tv1 = new Hashtable();
		Hashtable _tv2 = new Hashtable();

		string _sTreatmentTimeUnit = "";

		public InvMedProductMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region InvMedProductMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd InvMedProduct (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.InvMedProductMacro.InvMedProduct,ProtocolDTs.dll" elementLabel="Investigational Medicial Product" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="Description of administration of Test Articles." shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
#endif
			try 
			{
				mp.pba_.setOperation("Invesigational Medicinal Product Macro", "Generating information...");
				
				InvMedProductMacro macro = null;
				macro = new InvMedProductMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Invesigational Medicinal Product Macro"); 
				mp.inoutRng_.Text = "Invesigational Medicinal Product Macro: " + e.Message;
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

			// Find our soa
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

			// No soa, done
			if (_currentSOA == null) return;

			_sTreatmentTimeUnit = "";

			// Find ta1 and ta2
			int ctmCount = 0;
			CTMaterialEnumerator ctEnum = bom_.getCTMaterialEnumerator();
			while (ctEnum.MoveNext() && ctmCount < 2) 
			{
				ClinicalTrialMaterial ctm = ctEnum.getCurrent();
				ctmCount++;

				switch (ctmCount) 
				{
					case 1: _ta1 = ctm; break;
					case 2: _ta2 = ctm; break;
				}
			}

			// Missing a ta, done
			if (_ta1 == null || _ta2 == null) 
			{
				return;
			}

			_orderedVisits = PfizerUtil.getAllPlannedVisitsByArm(_currentSOA, _currentArm);

			// Get task visits for ta1, and ta2
			TaskEnumerator taskEnum = _currentSOA.getTaskEnumerator();
			while (taskEnum.MoveNext()) 
			{
				Task task = taskEnum.getCurrent();

				if (task.isDosingTask()) 
				{
					DosingTask dosingTask = new DosingTask(
						task.getObjectRoot(), icdSchemaMgr_.getTemplateByClass(typeof(DosingTask)));

					// get tv for ta1
					if (_ta1.getObjID() == dosingTask.getctMaterialID())
					{
						addTaskVisits(task.getObjID(), _tv1); 
					}

					// get tv for ta2
					if (_ta2.getObjID() == dosingTask.getctMaterialID())
					{
						addTaskVisits(task.getObjID(), _tv2); 
					}
				}
			}

			if (_sTreatmentTimeUnit.Length == 0) 
			{
				_sTreatmentTimeUnit = PfizerUtil.TimeUnit.getMax().getSystemName();
			}
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

			// Do we have all inputs?
			if (!readyToRumble(ref wrkRng)) 
			{
				inoutRange.End = wrkRng.End;
				setOutgoingRng(inoutRange);
				wdDoc_.UndoClear();

				return;
			}

			displayTa1(ref wrkRng);

			displayTa2(ref wrkRng);

			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}

		private void displayTa1(ref Word.Range wrkRng) 
		{
			ArrayList aAlone = new ArrayList();
			ArrayList aBoth = new ArrayList();
			
			// Fill Alone/Both buckets
			foreach (ProtocolEvent pe in _orderedVisits)
			{
				long visitID = pe.getObjID();

				TaskVisit tv1 = _tv1[visitID] as TaskVisit;
				TaskVisit tv2 = _tv2[visitID] as TaskVisit;

				if (tv1 != null && tv2 != null) 
				{
					aBoth.Add(pe);
				}
				else if (tv1 != null && tv2 == null) 
				{
					aAlone.Add(pe);
				}
				else
				{
				}
			}

			// Start writing
			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
				_ta1, ClinicalTrialMaterial.CTMATERIAL_NAME , wrkRng, macroEntry_);

			wrkRng.End = MacroBaseUtilities.putAfterElemRef(":", tspdDoc_, wrkRng);

			wrkRng.InsertParagraphAfter();
			//wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
				_ta1, ClinicalTrialMaterial.CTMATERIAL_NAME , wrkRng, macroEntry_);

			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
				_ta1, ClinicalTrialMaterial.FORMULATION , wrkRng, macroEntry_);

			wrkRng.InsertAfter("will be administered ");

			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
				_ta1, ClinicalTrialMaterial.ROUTE_OF_ADMIN , wrkRng, macroEntry_);

			wrkRng.InsertAfter("as a single dose of ");

			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
				_ta1, ClinicalTrialMaterial.DOSE , wrkRng, macroEntry_);

			if (aAlone.Count == 0) 
			{
				wrkRng.InsertAfter("on no events. ");
			}
			else
			{
				wrkRng.InsertAfter("on ");

				displayVisits(aAlone, ref wrkRng);
				
                wrkRng.InsertAfter("(alone)");
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}

			if (aBoth.Count == 0) 
			{
				if (aAlone.Count != 0) 
				{
					wrkRng.InsertAfter(". ");
				}
			}
			else
			{
				wrkRng.InsertAfter(" and on ");

				displayVisits(aBoth, ref wrkRng);
				
				wrkRng.InsertAfter("(in combination with ");

				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
					_ta2, ClinicalTrialMaterial.CTMATERIAL_NAME, wrkRng, macroEntry_);

				wrkRng.End = MacroBaseUtilities.putAfterElemRef(")", tspdDoc_, wrkRng);

				wrkRng.InsertAfter(". ");
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}

			wrkRng.InsertParagraphAfter();
			//wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			wdDoc_.UndoClear();
		}

		private void displayTa2(ref Word.Range wrkRng) 
		{
			ArrayList pvList = new ArrayList();

			// Fill bucket
			foreach (ProtocolEvent pe in _orderedVisits)
			{
				long visitID = pe.getObjID();

				TaskVisit tv2 = _tv2[visitID] as TaskVisit;

				if (tv2 != null) 
				{
					Period per = bom_.getParentOfScheduleItem(pe);
					PfizerUtil.PeriodAndVisit pv = new PfizerUtil.PeriodAndVisit();
					pv.per = per;
					pv.visit = pe;
					pvList.Add(pv);
				}
			}

			// Start writing
			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
				_ta2, ClinicalTrialMaterial.CTMATERIAL_NAME , wrkRng, macroEntry_);

			wrkRng.End = MacroBaseUtilities.putAfterElemRef(":", tspdDoc_, wrkRng);

			//wrkRng.InsertParagraphAfter();
			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);


			if (pvList.Count == 0) 
			{
				wrkRng.InsertAfter("There are no task/events defined.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				return;
			}

			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
				_ta2, ClinicalTrialMaterial.CTMATERIAL_NAME , wrkRng, macroEntry_);

			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
				_ta2, ClinicalTrialMaterial.FORMULATION , wrkRng, macroEntry_);

			wrkRng.InsertAfter("will be administered ");

			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
				_ta2, ClinicalTrialMaterial.ROUTE_OF_ADMIN , wrkRng, macroEntry_);

			wrkRng.InsertAfter("at ");

			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
				_ta2, ClinicalTrialMaterial.DOSE , wrkRng, macroEntry_);

			wrkRng.InsertAfter("for ");

			// Sort by visit order
			pvList.Sort(new PfizerUtil.PeriodAndVisitComparer());

			Period lastPer = null;
			foreach (PfizerUtil.PeriodAndVisit pv in pvList) 
			{
				Period curPer = pv.per;
				if (lastPer == null || curPer.getObjID() != lastPer.getObjID()) 
				{
					if (lastPer != null) 
					{
						wrkRng.InsertAfter(": ");
					}

					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, curPer, Period.BRIEF_DESCRIPTION, wrkRng, macroEntry_);
					wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);
				}
				else if (lastPer != null) 
				{
					wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);
				}

				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, pv.visit, ProtocolEvent.STUDY_DAYTIME, wrkRng, macroEntry_);

				lastPer = curPer;
			}


			// string totalDurationDisplay = PfizerUtil.getDisplayTime(totalDurationMinutes, _sTreatmentTimeUnit);
			// wrkRng.InsertAfter(totalDurationDisplay + ".");

			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			wdDoc_.UndoClear();
		}

		private void displayBadVisits(ArrayList al, ref Word.Range wrkRng, string message) 
		{
			if (al.Count == 0) return;

			wrkRng.InsertAfter(message);

			int i = 0;
			foreach (ProtocolEvent pe in al) 
			{
				i++;
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
					pe, ProtocolEvent.BRIEF_DESCRIPTION, wrkRng, macroEntry_);

				if (i < al.Count) 
				{
					wrkRng.End = MacroBaseUtilities.putAfterElemRef(", ", tspdDoc_, wrkRng);
				}
				else
				{
					wrkRng.End = MacroBaseUtilities.putAfterElemRef(".", tspdDoc_, wrkRng);
				}
				
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}

			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			wdDoc_.UndoClear();

		}
		private void displayVisits(ArrayList al, ref Word.Range wrkRng) 
		{
			int i = 0;
			Period lastPer = null;
			foreach (ProtocolEvent pe in al)
			{
				i++;
				Period per = _currentSOA.getGrandPeriodOfScheduleEvent(pe);

				if (lastPer != null && per.getObjID() != lastPer.getObjID()) 
				{
					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
						lastPer, Period.BRIEF_DESCRIPTION, wrkRng, macroEntry_);

					wrkRng.End = MacroBaseUtilities.putAfterElemRef("; ", tspdDoc_, wrkRng);
				}

				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
					pe, ProtocolEvent.BRIEF_DESCRIPTION, wrkRng, macroEntry_);

				wrkRng.End = MacroBaseUtilities.putAfterElemRef(", ", tspdDoc_, wrkRng);
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

				lastPer = per;
			}

			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
				lastPer, Period.BRIEF_DESCRIPTION, wrkRng, macroEntry_);

			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			wdDoc_.UndoClear();
		}

		public override void postProcess()
		{
			// Clean up memory
			_currentSOA = null;

			_ta1 = null;
			_ta2 = null;

			_tv1.Clear();
			_tv2.Clear();

			_orderedVisits.Clear();
		}

		private bool readyToRumble(ref Word.Range wrkRng) 
		{
			// Not found SOA
			if (_currentSOA == null)
			{
				pba_.updateProgress(70.0);

				wrkRng.InsertAfter("This schedule that this macro refers to was removed, delete this macro.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
						
				return false;
			}

			// Not defined ta1
			if (_ta1 == null)
			{
				pba_.updateProgress(70.0);

				wrkRng.InsertAfter("Test Article 1 is not defined.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
						
				return false;
			}

			// Not defined ta2
			if (_ta2 == null)
			{
				pba_.updateProgress(70.0);

				wrkRng.InsertAfter("Test Article 2 is not defined.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
						
				return false;
			}

			return true;
		}
		
		private void addTaskVisits(long taskID, Hashtable ht) 
		{
			foreach (ProtocolEvent pe in _orderedVisits)
			{
				long visitID = pe.getObjID();
				TaskVisit tv = _currentSOA.getOrCreateTaskVisit(taskID, visitID, false);

				if (tv == null) continue;

				Period per = _currentSOA.getGrandPeriodOfScheduleEvent(pe);
				string stype = per.getScheduleItemType();

				// Add if grandaddy is treatment
				if (!MacroBaseUtilities.isEmpty(stype) && stype.Equals("treatment"))
				{
					ht[tv.getAssociatedVisitID()] = tv;
				}
			
			}
		}

	}
}
