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
	internal sealed class BlindingUnblindingMacro
	{
		private static readonly string header_ = @"$Header: BlindingUnblindingMacro.cs, 1, 18-Aug-09 12:02:45, Pinal Patel$";
	}
}


namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for BlindingUnblindingMacro.
	/// </summary>
	public class BlindingUnblindingMacro : AbstractMacroImpl
	{
		public BlindingUnblindingMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Blinding Unblinding Methods
		
		#region BlindingUnblindingMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd BlindingUnblinding (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.BlindingUnblindingMacro.BlindingUnblinding,ProtocolDTs.dll" elementLabel="Blinding and Unblinding" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Concept" autogenerates="true" toolTip="Blinding and Unblinding." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Blinding Unblinding Macro", "Generating information...");
				
				BlindingUnblindingMacro macro = null;
				macro = new BlindingUnblindingMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Blinding Unblinding Macro");
				mp.inoutRng_.Text = "Blinding Unblinding Macro: " + e.Message;
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

			bool isOther = false;

			string blindingType = icpInstMgr_.getTypedDisplayValue(DesignDefines.MaskingType, out isOther);
			string blindingOtherValue = null;
			if (isOther) 
			{
				blindingOtherValue = icpInstMgr_.getTypedOtherDisplayValue(DesignDefines.MaskingType, out isOther);
			}

			if (MacroBaseUtilities.isEmpty(blindingType)) 
			{
				wrkRng.InsertAfter("Study Blinding not defined.");
				wrkRng.InsertParagraphAfter();
			}
			else if (blindingType.Equals("Open-label")) 
			{
				wrkRng.InsertAfter("Not applicable, study is open label.");
				wrkRng.InsertParagraphAfter();
			}
			else
			{
				tspdDoc_.insertLibraryItemByName("DT_BlindingandUnblinding", wrkRng);
			}

			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

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
