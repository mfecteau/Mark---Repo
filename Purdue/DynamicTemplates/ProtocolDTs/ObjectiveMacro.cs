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
	internal sealed class ObjectiveMacro
	{
		private static readonly string header_ = @"$Header: ObjectiveMacro.cs, 1, 18-Aug-09 12:04:55, Pinal Patel$";
	}
}


namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for ObjectiveMacro.
	/// </summary>
	public class ObjectiveMacro : AbstractMacroImpl
	{
		public static readonly string PRIMARY = "Primary";
		public static readonly string SECONDARY = "Secondary";

		private string objectiveType;
		private ArrayList objectives = new ArrayList();

		public ObjectiveMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region PrimaryObjective
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd PrimaryObjective (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.ObjectiveMacro.PrimaryObjective,ProtocolDTs.dll" elementLabel="Primary Objective" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Objectives" autogenerates="true" toolTip="Lists primary objective." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Primary Objective Macro", "Generating information...");
				
				ObjectiveMacro macro = null;
				macro = new ObjectiveMacro(mp);
				
				macro.objectiveType = PRIMARY;

				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Primary Objective Macro"); 
				mp.inoutRng_.Text = "Primary Objective Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		#endregion

		#region SecondaryObjective
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd SecondaryObjective (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.ObjectiveMacro.SecondaryObjective,ProtocolDTs.dll" elementLabel="Secondary Objectives" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Objectives" autogenerates="true" toolTip="Lists secondary objectives." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Secondary Objective Macro", "Generating information...");
				
				ObjectiveMacro macro = null;
				macro = new ObjectiveMacro(mp);
				
				macro.objectiveType = SECONDARY;

				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Secondary Objective Macro"); 
				mp.inoutRng_.Text = "Secondary Objective Macro: " + e.Message;
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

			if (objectiveType == PRIMARY) 
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

			if (objectives.Count == 0) 
			{
				wrkRng.InsertAfter("No primary objectives have been defined.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				return;
			}

			wrkRng.InsertAfter("Primary:");
			Word.Range rngPrimary = wrkRng.Duplicate;
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			rngPrimary.Font.Bold = VBAHelper.iTRUE;

			double progInc = 20.0 / (double)objectives.Count;

			bool numberList = (objectives.Count > 1);

			WordListHelper.ListTemplate wlt = WordListHelper.getNumberedListTemplate(wdApp_);

			foreach (Objective obj1 in objectives)
			{
				pba_.updateProgress(progInc);

				wlt.BeginListItem(ref wrkRng, numberList);

				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
					obj1, Objective.FULL_DESCRIPTION, wrkRng, macroEntry_);

				wlt.EndListItem(ref wrkRng);
			}
		}

		/// <summary>
		/// Rules:
		/// There must be a pharmacodynamic secondary objective.  This is tested
		/// by establishing that an objective is of type secondary and the string 
		/// "pd" or "pharmacodynamic" is present in the short description of the Objective
		/// 
		/// The pharmacodynamic objective is NOT printed in the output list if 'NA' is the
		/// text value of the table cell to the right of  the PHARMACODYNAMICS title cell in the
		/// synopsis table, otherwise that Objective is printed in the list of Objectives
		/// 
		/// If there is only one item to print, then turn off numbering
		/// </summary>
		/// <param name="wrkRng"></param>
		private void displaySecondary(ref Word.Range wrkRng)
		{
			string pdStatement = findPDStatement().ToLower();

			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			if (MacroBaseUtilities.isEmpty(pdStatement)) 
			{
				wrkRng.InsertAfter("Unable to locate Synopsis Pharmacodynamics section value.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				return;
			}

			bool omitPDObjective = pdStatement.Equals("na");
			ArrayList objectiveList = new ArrayList();

			foreach (Objective obj1 in objectives)
			{
				if (!obj1.getObjectiveType().Equals(SECONDARY))
				{
					continue;
				}
			
				if (omitPDObjective) 
				{
					string lowerName = obj1.getBriefDescription().ToLower();

					if (lowerName.IndexOf(" pd") != -1 ||
						lowerName.IndexOf("(pd") != -1 ||
						lowerName.IndexOf("pd)") != -1 ||
						lowerName.IndexOf("pharmacodynamic") != -1)
					{
						continue;
					}
				}

				objectiveList.Add(obj1);
			}

			// There are no objectives which meet the critera
			if (objectiveList.Count == 0) 
			{
				wrkRng.InsertAfter("There is no applicable secondary objective.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				return;
			}

			// If more than one number it
			bool numberList = (objectiveList.Count > 1);

			wrkRng.InsertAfter("Secondary:");
			Word.Range rngPrimary = wrkRng.Duplicate;
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			rngPrimary.Font.Bold = VBAHelper.iTRUE;

			double progInc = 20.0 / (double)objectives.Count;

			WordListHelper.ListTemplate wlt = WordListHelper.getNumberedListTemplate(wdApp_);

			foreach (Objective obj1 in objectiveList)
			{
				pba_.updateProgress(progInc);

				wlt.BeginListItem(ref wrkRng, numberList);

				wrkRng.End = MacroBaseUtilities.putElemRef(tspdDoc_, 
					obj1, Objective.FULL_DESCRIPTION, wrkRng, macroEntry_);
				
				wlt.EndListItem(ref wrkRng);

				wdDoc_.UndoClear();
			}

			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			wdDoc_.UndoClear();
		}

		public override void preProcess() 
		{
			try 
			{
				objectives.Clear();

				// Loop through all the objectives and group the associated outcomes by outcome
				// type, sorta.
				IEnumerator ie = bom_.getObjectives();
				int count = bom_.getObjectives().getList().Count;
				if (count > 0) 
				{
					double progInc = 30.0 / (double)count;
					while(ie.MoveNext()) 
					{
						Objective obj = (Objective)ie.Current;
						if (objectiveType.Equals(obj.getObjectiveType())) 
						{
							objectives.Add(obj);
						}

						pba_.updateProgress(progInc);
					}
				}
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Problem in preprocess()");
				throw e;
			}
		}

		public override void postProcess()
		{
			// Clean up memory
			objectives.Clear();
		}

		public string findPDStatement() 
		{
			string pdObjectiveTitle = "PHARMACODYNAMICS\r\a";
			string ra = "\r\a";

			for (int i = 1; i <= wdDoc_.Tables.Count; i++) 
			{
				try 
				{
					Word.Table tbl = wdDoc_.Tables[i];

					// Get table text, look for the title
					string tblText = tbl.Range.Text;
					int pdObj = tblText.IndexOf(pdObjectiveTitle);
					
					if (pdObj == -1) continue;

					// Now pick out the text
					string sPdObj = tblText.Substring(pdObj + pdObjectiveTitle.Length);
					int endPos = sPdObj.IndexOf(ra);

					if (endPos == -1) continue;

					// found it, trim down
					sPdObj = sPdObj.Substring(0, endPos);
					sPdObj = sPdObj.Trim();
					
					return sPdObj;
				} 
				catch (Exception ex) 
				{
					string s = ex.Message;
				}
			}

			return "";
		}
	}
}
