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
using MSXML2;

using Word = Microsoft.Office.Interop.Word;

namespace VersionControl 
{
	internal sealed class StudyDesignOverviewMacro
	{
		private static readonly string header_ = @"$Header: StudyDesignOverviewMacro.cs, 1, 18-Aug-09 12:05:49, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for StudyDesignOverviewMacro.
	/// </summary>
	public class StudyDesignOverviewMacro : AbstractMacroImpl
	{
		SOA _currentSOA;
		long _currentArm;
		Treatment _ipTreatment;
        Component _ipComponent;
        TestArticle _ipTestArticle;
		Period _washoutPeriod;

		public static readonly string sWASHOUT = "Washout";

		public StudyDesignOverviewMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
            _currentArm = ArmRule.ALL_ARMS;
		}
		
		#region Dynamic Tmplt Methods
		
		#region StudyDesignOverview

		public static MacroExecutor.MacroRetCd StudyDesignOverview (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.StudyDesignOverviewMacro.StudyDesignOverview,ProtocolDTs.dll" elementLabel="Study Design Overview" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="Summary of the Study Design" shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
#endif
			try 
			{
				mp.pba_.setOperation("Study Design Overview Macro", "Generating information...");
				
				StudyDesignOverviewMacro macro = null;
				macro = new StudyDesignOverviewMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Study Design Overview Macro"); 
				mp.inoutRng_.Text = "Study Design Overview Macro: " + e.Message;
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

            foreach (Treatment treatment in this.bom_.getTreatments().Enumerable.OfType<Treatment>().OrderBy(tr => tr.getSequence()))
            {
                foreach (Component component in bom_.getAssociatedComponents(treatment).Enumerable)
                {
                    TestArticle testArticle = bom_.getTestArticle(component.AssociatedTestArticleID);
                    if ((testArticle != null) && testArticle.PrimaryRole.Equals("investigationalProduct", StringComparison.InvariantCultureIgnoreCase))
                    {
                        _ipTreatment = treatment;
                        _ipComponent = component;
                        _ipTestArticle = testArticle;
                    }

                    pba_.updateProgress(2.0);
                }
            }

			if (_ipTreatment == null) 
                return;

			// Collect the ordered visit list
			ArrayList orderedTopLevelEvents = new ArrayList();
			_currentSOA.getTopLevelActivityList(_currentArm, null, orderedTopLevelEvents);

			_washoutPeriod = null;

			// Find the first washout period
			foreach (EventScheduleBase obj in orderedTopLevelEvents)
			{
				Period per = obj as Period;
				if (per == null) 
				{
					continue;
				}

				pba_.updateProgress(2.0);

				string stype = per.getScheduleItemType();
				if (!MacroBaseUtilities.isEmpty(stype) && stype.Equals(sWASHOUT)) 
				{
					_washoutPeriod = per;
					break;
				}
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

			if (_ipTreatment == null)
			{
				pba_.updateProgress(70.0);

				wrkRng.InsertAfter("There is no test article defined with a primary role of investigational product.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
						
				inoutRange.End = wrkRng.End;
				setOutgoingRng(inoutRange);
				wdDoc_.UndoClear();
				return;
			}

			// Paths to variables/custom variables
			string secondaryControlTypePath = "/FTICP/StudyDesign/Design/secondarycontrolType";
			string tertiaryControlTypePath = "/FTICP/StudyDesign/Design/tertiarycontrolType";

			string numberOfDosesTypePath = "/FTICP/StudyDesign/Design/NumberOfDoses";

			bool isOther;
			string secondaryContolType = icpInstMgr_.getTypedDisplayValue(secondaryControlTypePath, out isOther);
			string tertiaryContolType = icpInstMgr_.getTypedDisplayValue(tertiaryControlTypePath, out isOther);

			// Start writing
			wrkRng.InsertAfter("This is a ");

			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, AdminDefines.IsMultiCenteredType, wrkRng, macroEntry_);
			wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);

			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, numberOfDosesTypePath, wrkRng, macroEntry_);
			wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);

//			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, _ipCTM, ClinicalTrialMaterial.DOSE, wrkRng, macroEntry_);
//			wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);

			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, DesignDefines.MethodOfAllocationType, wrkRng, macroEntry_);
			wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);

			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, DesignDefines.MaskingType, wrkRng, macroEntry_);
			wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);

			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, DesignDefines.ControlType, wrkRng, macroEntry_);
			wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);

			// Write if we have 2nd
			if (!MacroBaseUtilities.isEmpty(secondaryContolType)) 
			{
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, secondaryControlTypePath, wrkRng, macroEntry_);
				wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);
			}

			// Write if we have 3rd
			if (!MacroBaseUtilities.isEmpty(tertiaryContolType)) 
			{
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, tertiaryControlTypePath, wrkRng, macroEntry_);
				wrkRng.End = MacroBaseUtilities.putAfterElemRef(",", tspdDoc_, wrkRng);
			}

			wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, DesignDefines.StudyConfigurationType, wrkRng, macroEntry_);
			
			// We have washout
			if (_washoutPeriod != null) 
			{
				wrkRng.InsertAfter("(with a washout period of ");

				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, _washoutPeriod, Period.DURATION, wrkRng, macroEntry_);
				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, _washoutPeriod, Period.DURATION_UNIT, wrkRng, macroEntry_);

				wrkRng.InsertAfter("between treatments) trial.");

				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
			else
			{
				wrkRng.InsertAfter("trial.");

				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
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
			_ipTreatment = null;
            _ipComponent = null;
            _ipTestArticle = null;
			_washoutPeriod = null;
		}
	}
}
