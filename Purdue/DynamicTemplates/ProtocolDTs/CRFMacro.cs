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
	internal sealed class CRFMacro
	{
		private static readonly string header_ = @"$Header: CRFMacro.cs, 1, 18-Aug-09 12:03:20, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for CRFMacro.
	/// </summary>
	public class CRFMacro : AbstractMacroImpl
	{
		public CRFMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region CRFMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd CRF (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.CRFMacro.CRF,ProtocolDTs.dll" elementLabel="Case Report Form" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Identification" autogenerates="true" toolTip="Creates text based on use of EDC." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("CRF Macro", "Generating information...");
				
				CRFMacro macro = null;
				macro = new CRFMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in CRF Macro"); 
				mp.inoutRng_.Text = "CRF Macro: " + e.Message;
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

			string EDCStudyPath = "/FTICP/Administrative/ProtocolSkeleton/EDCStudy";

			bool isOther;
			string EDCStudyType = icpInstMgr_.getTypedDisplayValue(EDCStudyPath, out isOther);

			if (MacroBaseUtilities.isEmpty(EDCStudyType)) 
			{
				wrkRng.InsertAfter("Please select a value for EDC Study in the Administration custom element area.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
			else if (EDCStudyType == "true") 
			{
				tspdDoc_.insertLibraryItemByName("DT_EDCStudy", wrkRng);
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
			else
			{
				tspdDoc_.insertLibraryItemByName("DT_NonEDCStudy", wrkRng);
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
		}
	}
}
