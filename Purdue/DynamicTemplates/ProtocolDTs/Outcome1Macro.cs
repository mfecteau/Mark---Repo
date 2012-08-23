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
	internal sealed class Outcome1Macro
	{
		private static readonly string header_ = @"$Header: Outcome1Macro.cs, 1, 18-Aug-09 12:05:02, Pinal Patel$";
	}
}


namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for ObjectiveMacro.
	/// </summary>
	public class Outcome1Macro : AbstractMacroImpl
	{
		public static readonly string PRIMARY = "Primary";
		public static readonly string SECONDARY = "Secondary";

		private string outcomeType;
		private ArrayList outcomes = new ArrayList();

		public Outcome1Macro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region PrimaryOutcome
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd PrimaryOutcome (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.OutcomeMacro.PrimaryOutcome,ProtocolDTs.dll" elementLabel="Primary Outcome" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Objectives" autogenerates="true" toolTip="Lists primary outcomes." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Primary Outcome Macro", "Generating information...");
				
				Outcome1Macro macro = null;
				macro = new Outcome1Macro(mp);
				
				macro.outcomeType = PRIMARY;

				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Primary Efficacy Macro"); 
				mp.inoutRng_.Text = "Primary Efficacy Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		#endregion

		#region SecondaryOutcome
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd SecondaryOutcome (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.OutcomeMacro.SecondaryObjective,ProtocolDTs.dll" elementLabel="Secondary Outcome" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Objectives" autogenerates="true" toolTip="Lists secondary outcomes." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Secondary Efficacy Macro", "Generating information...");
				
				Outcome1Macro macro = null;
				macro = new Outcome1Macro(mp);
				
				macro.outcomeType = SECONDARY;

				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Secondary Outcome Macro"); 
				mp.inoutRng_.Text = "Secondary Outcome Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		#endregion
		
		#endregion
		public override void preProcess() 
		{
			try 
			{
				outcomes.Clear();

				/*****/


				OutcomeEnumerator oe = bom_.getOutcomes();
				int count = bom_.getOutcomes().getList().Count;

				double progInc = 30.0 / (double)count;

				while(oe.MoveNext()) 
					{
						Outcome outcome1 = (Outcome)oe.Current;
						ObjectiveEnumerator objEnum =  bom_.getAssociatedObjectives(outcome1);

								while(objEnum.MoveNext())
								{
									Objective PriObj = (Objective)objEnum.Current;

									if (PriObj.getObjectiveType().Equals(outcomeType))
									{
										outcomes.Add(outcome1);
									}

								}


////					outcomes.Add(outcome1);
////						// check the outcome rank here
////
////					if (outcome1.getOutcomeRank().Equals(Outcome.Rank.Primary))
////						{
////							outcomes.Add(outcome1);
////						}
					pba_.updateProgress(progInc);

					}

////
////				// Loop through all the outcomes and group the associated outcomes by outcome
////				// type, sorta.
//				IEnumerator ie = bom_.getOutcomes();
//				int count = bom_.getOutcomes().getList().Count;
//				if (count > 0) 
//				{
//					double progInc = 30.0 / (double)count;
//					while(ie.MoveNext()) 
//					{
//						Outcome obj = (Outcome)ie.Current;
//						
//						if (outcomeType.Equals(obj.getOutcomeType())) 
//						{
//							outcomes.Add(obj);
//						}
//
//						pba_.updateProgress(progInc);
//					}
//				}
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Problem in preprocess()");
				throw e;
			}
		}


		public override void display() 
		{
			Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;

			pba_.updateProgress(1.0);

			string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);

			if (outcomeType == PRIMARY) 
			{
				displayPrimary(ref wrkRng);
			}
			else
			{
				displaySecondary(ref wrkRng);
			}

			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}

		private void displayPrimary(ref Word.Range wrkRng)
		{
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			if (outcomes.Count == 0) 
			{
				wrkRng.InsertAfter("No primary outcome(s) have been defined.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				return;
			}

			wrkRng.InsertAfter("Primary Outcome:");
			Word.Range rngPrimary = wrkRng.Duplicate;
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			rngPrimary.Font.Bold = VBAHelper.iTRUE;

			double progInc = 20.0 / (double)outcomes.Count;

			bool numberList = (outcomes.Count > 1);

			WordListHelper.ListTemplate wlt = WordListHelper.getNumberedListTemplate(wdApp_);

			foreach (Outcome obj1 in outcomes)
			{

				pba_.updateProgress(progInc);

				wlt.BeginListItem(ref wrkRng, numberList);

				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
					obj1, Outcome.BRIEF_DESCRIPTION , wrkRng, macroEntry_);
				
				wlt.EndListItem(ref wrkRng);
			}

			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			wdDoc_.UndoClear();
		}

		
		private void displaySecondary(ref Word.Range wrkRng)
		{
			
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			if (outcomes.Count == 0) 
			{
				wrkRng.InsertAfter("No secondary outcome(s) have been defined.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				return;
			}

			wrkRng.InsertAfter("Secondary Endpoint:");
			Word.Range rngPrimary = wrkRng.Duplicate;
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			rngPrimary.Font.Bold = VBAHelper.iTRUE;

			double progInc = 20.0 / (double)outcomes.Count;

			bool numberList = (outcomes.Count > 1);

			WordListHelper.ListTemplate wlt = WordListHelper.getNumberedListTemplate(wdApp_);

			foreach (Outcome obj1 in outcomes)
			{

				pba_.updateProgress(progInc);

				wlt.BeginListItem(ref wrkRng, numberList);

				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
					obj1, Outcome.BRIEF_DESCRIPTION , wrkRng, macroEntry_);
				
				wlt.EndListItem(ref wrkRng);
			}

			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			wdDoc_.UndoClear();



		}

		
		public override void postProcess()
		{
			// Clean up memory
			outcomes.Clear();
		}

	
	}
}
