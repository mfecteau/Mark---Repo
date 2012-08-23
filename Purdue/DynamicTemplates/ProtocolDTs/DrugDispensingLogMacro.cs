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
	internal sealed class DrugDispensingLogMacro
	{
		private static readonly string header_ = @"$Header: DrugDispensingLogMacro.cs, 1, 18-Aug-09 12:03:42, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for DrugDispensingLogMacro.
	/// </summary>
	public class DrugDispensingLogMacro : AbstractMacroImpl
	{
		public DrugDispensingLogMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region DrugDispensingLogMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd DrugDispensingLog (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.DrugDispensingLogMacro.DrugDispensingLog,ProtocolDTs.dll" elementLabel="Drug Dispensing Log" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Test Article" autogenerates="true" toolTip="Drug dispensing log." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("DrugDispensingLog Macro", "Generating information...");
				
				DrugDispensingLogMacro macro = null;
				macro = new DrugDispensingLogMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in DrugDispensingLog Macro"); 
				mp.inoutRng_.Text = "DrugDispensingLog Macro: " + e.Message;
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

			string aicAttribute = "AdministeredinClinic";

			bool haveEmptyValue = false;
			bool haveAdministeredInClinic = false;
			int ctmCount = 0;
			CTMaterialEnumerator ctEnum = bom_.getCTMaterialEnumerator();
			while (ctEnum.MoveNext()) 
			{
				ClinicalTrialMaterial ctm = ctEnum.getCurrent();
				ctmCount++;;

				string aicValue = (string )ctm.getValueForNode(aicAttribute);

				if (MacroBaseUtilities.isEmpty(aicValue)) 
				{
					haveEmptyValue = true;
				}
				else if (aicValue.Equals("true")) 
				{
					haveAdministeredInClinic = true;
				}
			}

			if (ctmCount == 0) 
			{
				wrkRng.InsertAfter("There are no study drugs defined.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			} 
			else if (haveEmptyValue)
			{
				wrkRng.InsertAfter("Please specify a value for the 'Drug Administered in Clinic' field for the study drug.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
			else if (haveAdministeredInClinic)
			{
				tspdDoc_.insertLibraryItemByName("DT_DrugDispensedClinic", wrkRng);
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
			else
			{
				tspdDoc_.insertLibraryItemByName("DT_DrugDispensedNoClinic", wrkRng);
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
