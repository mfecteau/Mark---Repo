#define xUSE_TVTIMES

using System;
using System.Collections;

using System.Windows.Forms;

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
	internal sealed class LengthOfStudyMacro
	{
		private static readonly string header_ = @"$Header: LengthOfStudyMacro.cs, 1, 18-Aug-09 12:04:41, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for LengthOfStudyMacro.
	/// </summary>
	public class LengthOfStudyMacro : AbstractMacroImpl
	{
		public static string Pfizer_LOS_BULLETED_LIST = "TSPDLOSBulletList";

		public static readonly string sSCREENING = "screening";
		public static readonly string sTREATMENT = "treatment";
		public static readonly string sWASHOUT = "Washout";
		public static readonly string sFOLLOWUP = "followUp";

		public static string _sDurationTimeUnit = "";

		private int _nTPs = 0;
		private int _nTPsWithCycles = 0;
		private int _nWOs = 0;

		SOA _currentSOA = null;
		long _currentArm = ArmRule.ALL_ARMS;

		ArrayList _invalidPer = new ArrayList();

		Hashtable _htPeriodsByType = new Hashtable();

		long _studyDuration = 0;

		public LengthOfStudyMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region LengthOfStudyMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd LengthOfStudy (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.LengthOfStudyMacro.LengthOfStudy,ProtocolDTs.dll" elementLabel="Length Of Study" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="Narrative Study Schedule Study Length." shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
#endif
			try 
			{
				mp.pba_.setOperation("Length Of Study Macro", "Generating information...");
				
				LengthOfStudyMacro macro = null;
				macro = new LengthOfStudyMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Length Of Study Macro"); 
				mp.inoutRng_.Text = "Length Of Study Macro: " + e.Message;
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

			ChooserEntry perTemplate = icdSchemaMgr_.getTemplateByClass(typeof(Period));
			IChooserEntry perDurationMeta = perTemplate.getMetaData(Period.DURATION_UNIT);
			ArrayList durationEnumPairs = icpSchemaMgr_.getEnumPairs(perDurationMeta.getDropdownListName());

			bool bValidParms = true;

			// Get stored parameters
			string sTimeUnit = execParms_.getParm(MacroExecutor.MacroExecParameters.PARM_1);
			
			if (MacroBaseUtilities.isEmpty(sTimeUnit))
			{
				bValidParms = false;
			}
			else 
			{
				bool found = false;
				foreach (EnumPair ep in durationEnumPairs) 
				{
					if (ep.getSystemName().Equals(sTimeUnit)) 
					{
						found = true;
						break;
					}
				}

				if (!found) 
				{
					bValidParms = false;
				}
			}
			
			if (!bValidParms) 
			{
				DurationSelect durSelect = new DurationSelect();
				durSelect.loadDurations(durationEnumPairs);

				DialogResult res = durSelect.ShowDialog();
				if (res == DialogResult.OK) 
				{
					EnumPair ep = durationEnumPairs[durSelect.SelectedDuration] as EnumPair;
					sTimeUnit = ep.getSystemName();

	 				execParms_.setParm(MacroExecutor.MacroExecParameters.PARM_1, sTimeUnit);
				}
			}

			// We have a good value now
			_sDurationTimeUnit = sTimeUnit;

			// Nope
			if (MacroBaseUtilities.isEmpty(_sDurationTimeUnit)) 
			{
				return;
			}

			

			// Check period duration/duration time unit
			ArrayList orderedTopLevelEvents = new ArrayList();
			_currentSOA.getTopLevelActivityList(_currentArm, null, orderedTopLevelEvents);
			foreach (EventScheduleBase obj in orderedTopLevelEvents)
			{
				Period per = obj as Period;
				if (per == null) 
				{
					continue;
				}

				string stype = per.getScheduleItemType();
				if (MacroBaseUtilities.isEmpty(stype))
				{
					_invalidPer.Add(per);
					continue;
				}

				Hashtable hPer = _htPeriodsByType[stype] as Hashtable;
				if (hPer == null) 
				{
					hPer = new Hashtable();
					_htPeriodsByType[stype] = hPer; 
				}
				
				hPer[per.getObjID()] = per;

				bool isBadTime = false;
				
				int nCycles = 0;
				if (stype.Equals(sWASHOUT)) 
				{
					_nWOs++;
				}
				else if (stype.Equals(sTREATMENT)) 
				{
					_nTPs++;
					nCycles = getCycleCount(per);

					if (nCycles != 0) 
					{
						_nTPsWithCycles++;
					}
				}
								
				for (int i = 0; i <= nCycles; i++) 
				{
					PfizerUtil.addTimeUnit(ref _studyDuration, per.getDuration(), out isBadTime);

					if (isBadTime) 
					{
						break;
					}
				}
				
				if (isBadTime) 
				{
					_invalidPer.Add(per);						
					continue;
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

			if (MacroBaseUtilities.isEmpty(_sDurationTimeUnit)) 
			{
				wrkRng.InsertAfter("You must a duration unit.");
				wrkRng.InsertParagraphAfter();

				inoutRange.End = wrkRng.End;
				setOutgoingRng(inoutRange);
				wdDoc_.UndoClear();
				return;
			}

			Hashtable htScreen = _htPeriodsByType[sSCREENING] as Hashtable;
			Hashtable htTreatment = _htPeriodsByType[sTREATMENT] as Hashtable;
			Hashtable htWashout = _htPeriodsByType[sWASHOUT] as Hashtable;
			Hashtable htFU = _htPeriodsByType[sFOLLOWUP] as Hashtable;

			#region ReadyToRumble

			if (htScreen == null || htFU == null) 
			{
				wrkRng.InsertAfter("You must define a <Screening> and <Follow Up> Period in your ");
				wrkRng.InsertAfter("Schedule of Activities.");
				wrkRng.InsertParagraphAfter();

				inoutRange.End = wrkRng.End;
				setOutgoingRng(inoutRange);
				wdDoc_.UndoClear();
				return;
			}

			if (htTreatment == null) 
			{
				wrkRng.InsertAfter("You must define one or more <Treatment> in your ");
				wrkRng.InsertAfter("Schedule of Activities.");
				wrkRng.InsertParagraphAfter();

				inoutRange.End = wrkRng.End;
				setOutgoingRng(inoutRange);
				wdDoc_.UndoClear();
				return;
			}

			bool haveInvalid = false;
			if (_invalidPer.Count != 0) 
			{
				haveInvalid = true;
				foreach (Period p in _invalidPer) 
				{
					wrkRng.InsertAfter("You must first specify a duration for Period: ");
					wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
						p, Period.BRIEF_DESCRIPTION, wrkRng, macroEntry_);
					wrkRng.InsertAfter("and the duration unit, as well as the treatment epoch.");
					wrkRng.InsertParagraphAfter();
				}
			}


			// Round up to nearest greater timeunit, fail if 0
			long timeMult = PfizerUtil.TimeUnit.find(_sDurationTimeUnit).getMultiplier();
			long durationInt = _studyDuration / timeMult;
			long durationRem = _studyDuration % timeMult;

			if (durationInt == 0) 
			{
				haveInvalid = true;
				wrkRng.InsertAfter("The duration unit must be greater than " + _sDurationTimeUnit + ".");
				wrkRng.InsertParagraphAfter();
			}

			if (haveInvalid) 
			{
				inoutRange.End = wrkRng.End;
				setOutgoingRng(inoutRange);
				wdDoc_.UndoClear();
				return;
			}

			#endregion

			bool isBadTime;
			long timeSeconds;

			// Total study duration, round up to nearest greater timeunit
			if (durationRem != 0) 
			{
				durationInt++;
			}
			string s = PfizerUtil.getDisplayTime(durationInt * timeMult, _sDurationTimeUnit);

			wrkRng.InsertAfter("Up to");
			wrkRng.InsertAfter(s);
			wrkRng.InsertAfter(" (from screening through study completion) for each enrolled subject as follows:");
			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);


			WordListHelper.ListTemplate wlt = WordListHelper.getBulletListTemplate(wdApp_);
			wlt.BeginListItem(ref wrkRng);

			// Summarize Screening periods
			timeSeconds = 0;
			foreach (Period p in htScreen.Values) 
			{
				PfizerUtil.addTimeUnit(ref timeSeconds, p.getDuration(), out isBadTime);
			}

			s = PfizerUtil.getDisplayTime(timeSeconds, _sDurationTimeUnit);

			wrkRng.InsertAfter("Screening: up to");
			wrkRng.InsertAfter(s);
							
			wlt.EndListItem(ref wrkRng);
		
			// New Item
			wlt.BeginListItem(ref wrkRng);

			if (_nTPs == 1 && _nTPsWithCycles == 1 && _nWOs == 1) 
			{
				// If only one Treatment period with cycles
				IEnumerator en = htTreatment.Values.GetEnumerator();
				en.MoveNext();
				Period p = en.Current as Period;

				int nCycles = getCycleCount(p);

				timeSeconds = 0;
				PfizerUtil.addTimeUnit(ref timeSeconds, p.getDuration(), out isBadTime);

				s = PfizerUtil.getDisplayTime(timeSeconds, _sDurationTimeUnit);
				
				int nTotal = nCycles + 1;
				wrkRng.InsertAfter("Treatment periods: " + nTotal.ToString() + " periods, each" + s);
			}
			else
			{
				// Summarize Treatment periods
				timeSeconds = 0;
				foreach (Period p in htTreatment.Values) 
				{
					int nCycles = getCycleCount(p);

					for (int i = 0; i <= nCycles; i++) 
					{
						PfizerUtil.addTimeUnit(ref timeSeconds, p.getDuration(), out isBadTime);
					}
				}

				s = PfizerUtil.getDisplayTime(timeSeconds, _sDurationTimeUnit);

				wrkRng.InsertAfter("Treatment periods: In total" + s);
			}

			// Summarize Washout periods
			if (htWashout != null) 
			{
				timeSeconds = 0;
				foreach (Period p in htWashout.Values) 
				{
					PfizerUtil.addTimeUnit(ref timeSeconds, p.getDuration(), out isBadTime);
				}

				s = PfizerUtil.getDisplayTime(timeSeconds, _sDurationTimeUnit);

				wrkRng.InsertAfter(" with washout of" + s);
			}
			
			wlt.EndListItem(ref wrkRng);


			// List Follow Up periods
			foreach (Period p in htFU.Values) 
			{
				wlt.BeginListItem(ref wrkRng);

				timeSeconds = 0;
				PfizerUtil.addTimeUnit(ref timeSeconds, p.getDuration(), out isBadTime);
				s = PfizerUtil.getDisplayTime(timeSeconds, _sDurationTimeUnit);

				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
					p, Period.BRIEF_DESCRIPTION, wrkRng, macroEntry_);

				wrkRng.End = MacroBaseUtilities.putAfterElemRef(":", tspdDoc_, wrkRng);

				wrkRng.InsertAfter(s);

				wrkRng.InsertAfter(" after last dosing (study completion)");
				
				wlt.EndListItem(ref wrkRng);
			}

			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}

		private int getCycleCount(Period p) 
		{
			int nCycles = 0;

			LinkingRuleMgr lrm = _currentSOA.getLinkManager();
			IList ats = lrm.repeatRuleWalker(_currentArm, p, LinkingRuleMgr.Motion.Forward, null, true);

			if (ats.Count != 0) 
			{
				LinkingRuleMgr.ActivityTarget at = ats[0] as LinkingRuleMgr.ActivityTarget;
				CycleRule cr = at.ActivityRule as CycleRule;
				nCycles = cr.getLimit();
			}


			return nCycles;
		}

		public override void postProcess()
		{
			// Clean up memory
			_currentSOA = null;

			_htPeriodsByType.Clear();

			_invalidPer.Clear();

			_sDurationTimeUnit = "";
		}
	}
}
