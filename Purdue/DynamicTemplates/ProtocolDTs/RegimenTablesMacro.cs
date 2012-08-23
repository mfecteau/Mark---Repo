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
	internal sealed class RegimenTablesMacro
	{
		private static readonly string header_ = @"$Header: RegimenTablesMacro.cs, 1, 18-Aug-09 12:05:36, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for RegimenTablesMacro.
	/// </summary>
	public class RegimenTablesMacro : AbstractMacroImpl
	{
		private class ArmCTMPair
		{
			public Arm arm;
			public CTMaterialToArm ctta;
			public ClinicalTrialMaterial ctm;
		}

		public static string CTMaterialPriRole_IP = "investigationalProduct";
		public static string CTMaterialPriRole_Placebo = "placebo";
		public static string CTMaterialPriRole_Comparator = "comparator";

		private string _currentCTMaterialPriRole = "";
		private int _foundCTMs = 0;
		private int _foundArms = 0;

		ArrayList _armCtmList = new ArrayList();

		public RegimenTablesMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		
		#region Dynamic Template Methods

		#region IPRegimenTable
		/// <summary>
		/// Displays information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd IPRegimenTable (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.RegimenTablesMacro.IPRegimenTable,ProtocolDTs.dll" elementLabel="IP Dose Route Regimen" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Test Article" autogenerates="true" toolTip="Table of Dosing per Arm." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("IP Regimen Table Macro", "Generating information...");
				
				RegimenTablesMacro macro = null;
				macro = new RegimenTablesMacro(mp);
				
				macro._currentCTMaterialPriRole = CTMaterialPriRole_IP;

				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in IP Regimen Table Macro"); 
				mp.inoutRng_.Text = "IP Regimen Table Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}
		#endregion

		#region PlaceboRegimenTable
		/// <summary>
		/// Displays information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd PlaceboRegimenTable (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.RegimenTablesMacro.PlaceboRegimenTable,ProtocolDTs.dll" elementLabel="Placebo Dose Route Regimen" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Test Article" autogenerates="true" toolTip="Table of Dosing per Arm." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Placebo Regimen Table Macro", "Generating information...");
				
				RegimenTablesMacro macro = null;
				macro = new RegimenTablesMacro(mp);
				
				macro._currentCTMaterialPriRole = CTMaterialPriRole_Placebo;

				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Placebo Regimen Table Macro"); 
				mp.inoutRng_.Text = "Placebo Regimen Table Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}
		#endregion

		#region ComparatorRegimenTable
		/// <summary>
		/// Displays information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd ComparatorRegimenTable (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.RegimenTablesMacro.ComparatorRegimenTable,ProtocolDTs.dll" elementLabel="Comparator Dose Route Regimen" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Test Article" autogenerates="true" toolTip="Table of Dosing per Arm." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Comparator Regimen Table Macro", "Generating information...");
				
				RegimenTablesMacro macro = null;
				macro = new RegimenTablesMacro(mp);
				
				macro._currentCTMaterialPriRole = CTMaterialPriRole_Comparator;

				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Comparator Regimen Table Macro"); 
				mp.inoutRng_.Text = "Comparator Regimen Table Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}
		#endregion


		#endregion

		public override void preProcess()
		{
			CTMaterialEnumerator ctmEnum = bom_.getCTMaterialEnumerator();
			_foundCTMs = ctmEnum.getList().Count;

			ArmEnumerator armEnum = bom_.getArmEnumerator();
			_foundArms = armEnum.getList().Count;

			while (armEnum.MoveNext()) 
			{
				pba_.updateProgress(1.0);

				Arm arm = armEnum.getCurrent();

				CTToArmEnumerator cttaEnum = bom_.getCTMaterialToArmForArm(arm);
				while (cttaEnum.MoveNext()) 
				{
					pba_.updateProgress(1.0);

					CTMaterialToArm ctta = cttaEnum.getCurrent();
					long ctmObjID = ctta.getAssociatedMaterialID();
					ClinicalTrialMaterial ctm = findCTM(ctmObjID);

					if (ctm != null) 
					{
						if (ctta.getOfficialRole() == _currentCTMaterialPriRole) 
						{
							ArmCTMPair actp = new ArmCTMPair();
							actp.arm = arm;
							actp.ctm = ctm;
							actp.ctta = ctta;

							_armCtmList.Add(actp);
						}
					}
				}
			}
		}

		public override void display()
		{
			Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;

			pba_.updateProgress(1.0);

			string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);

			if (_foundCTMs == 0) 
			{
				wrkRng.InsertAfter("No Test Articles defined.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
			else if (_foundArms == 0) 
			{
				wrkRng.InsertAfter("No Study Arms defined.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
			else
			{
				bool macroContext = tmpIsInMacroContext(macroEntry_);

				if (macroContext) 
				{
					wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				}

				displayRegimenTable(wrkRng);

				if (macroContext) 
				{
					wrkRng.End = MacroBaseUtilities.insertSquashedParagraphAfter(wrkRng);
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
			_foundCTMs = 0;
			_foundArms = 0;

			_armCtmList.Clear();
		}

		private bool tmpIsInMacroContext(ChooserEntry macroEntry) 
		{
			if (macroEntry != null) 
			{
				bool isProtected = IcpDefines.ConvertToBoolean(
					macroEntry.getValueForNode(Tspd.Businessobject.MacroEntry.PROTECTED), true);
				bool autoGenerates = IcpDefines.ConvertToBoolean(
					macroEntry.getValueForNode(Tspd.Businessobject.MacroEntry.AUTOGENERATES), true);

				if (!isProtected && !autoGenerates) 
				{
					return false;
				}
			}

			return true;
		}
		
		private void displayRegimenTable(Word.Range wrkRng) 
		{
			wdDoc_.UndoClear();

			if (_armCtmList.Count == 0) 
			{
				wrkRng.InsertAfter("No Test Article to Study Arm association defined.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

				return;
			}

			Word.Table tbl = createTable(wrkRng, _armCtmList.Count, 3);
			int currentRow = 0;

			foreach (ArmCTMPair actp in _armCtmList) 
			{
				currentRow++;
				Word.Row tableRow = tbl.Rows[currentRow];

				// Cell 1
				Word.Cell tableCell = tableRow.Cells[1];

				Word.Range cellRange = tableCell.Range.Duplicate;
				cellRange.Collapse(ref WordHelper.COLLAPSE_END);
				cellRange.End--;
				cellRange.InsertAfter("Dose Group ");

				MacroBaseUtilities.putElemRefInCell(tspdDoc_, tableCell, 
						actp.arm, Arm.BRIEF_DESCRIPTION, 
						true, macroEntry_);

				// Cell 2
				tableCell = tableRow.Cells[2];
				MacroBaseUtilities.putElemRefInCell(tspdDoc_, tableCell, 
					actp.ctm, ClinicalTrialMaterial.DOSE, 
					true, macroEntry_);

				// Cell 3
				tableCell = tableRow.Cells[3];
				MacroBaseUtilities.putElemRefInCell(tspdDoc_, tableCell, 
					actp.ctm, "DosingRegimen", 
					true, macroEntry_);
			}

			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			wdDoc_.UndoClear();
		}

		private ClinicalTrialMaterial findCTM(long ctmObjID) 
		{
			CTMaterialEnumerator ctmEnum = bom_.getCTMaterialEnumerator();
			while (ctmEnum.MoveNext()) 
			{
				ClinicalTrialMaterial ctm = ctmEnum.getCurrent();
				if (ctm.getObjID() == ctmObjID) 
				{
					return ctm;
				}
			}

			return null;
		}

		public virtual Word.Table createTable(Word.Range viewRng, int rows, int cols) 
		{

			// Turn off auto caption for Word tables.
			Word.AutoCaption ac = wdApp_.AutoCaptions.get_Item(ref WordHelper.AUTO_CAPTION_WORD_TABLE);
			bool oldState = ac.AutoInsert;
			ac.AutoInsert = false;

			Word.Range wrkRng = viewRng.Duplicate;


			if (false) 
			{
				object tableCaption = "\ttable caption";
				wrkRng.InsertAfter(" "); // single space which'll get shifted after the caption.
				Word.Range captionRng = wrkRng.Duplicate;
				captionRng.Collapse(ref WordHelper.COLLAPSE_START);

				// put the caption in if the TableView gives you one...
				if (tableCaption != null)
				{
					captionRng.InsertCaption(
						ref WordHelper.CAPTION_LABEL_TABLE, ref tableCaption,
						ref VBAHelper.OPT_MISSING, ref WordHelper.CAPTION_POSITION_ABOVE, ref VBAHelper.OPT_MISSING);
				}

				// Mark the added single space and replace it with a paragraph mark.
				wrkRng.Start = wrkRng.End - 1;
				wrkRng.InsertParagraph();
				// Collapse the range to the end of the paragraph mark so that the table can be added
				// after it.
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}

			// Insert the table. Note, that the table is inserted starting at but after the range.
			// So viewRng isn't increased.
			Word.Table tbl = wdDoc_.Tables.Add(
				wrkRng, rows, cols,
				ref WordHelper.WORD8_TABLE_BEHAVIOR, ref VBAHelper.OPT_MISSING);


			tbl.Borders.Enable = VBAHelper.iTRUE;
			tbl.Borders.InsideLineWidth = MacroBaseUtilities.LINE_WIDTH_NORMAL;

			// Reinstate auto caption for Word tables.
			ac.AutoInsert = oldState;

			// Increase viewRng to include the table.
			viewRng.End = tbl.Range.End;

			viewRng.Collapse(ref WordHelper.COLLAPSE_END);

			return tbl;
		}
	}
}
