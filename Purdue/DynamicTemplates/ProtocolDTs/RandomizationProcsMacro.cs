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
	internal sealed class RandomizationProcsMacro
	{
		private static readonly string header_ = @"$Header: RandomizationProcsMacro.cs, 1, 18-Aug-09 12:05:35, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for RandomizationProcsMacro.
	/// </summary>
	public class RandomizationProcsMacro : AbstractMacroImpl
	{
		public RandomizationProcsMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Template Methods
		
		#region RandomizationProcsMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd RandomizationProcs (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.RandomizationProcsMacro.RandomizationProcs,ProtocolDTs.dll" elementLabel="Randomization Procedures" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Concept" autogenerates="true" toolTip="Creates randomization text based on study blind." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Randomization Procedures Macro", "Generating information...");
				
				RandomizationProcsMacro macro = null;
				macro = new RandomizationProcsMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Randomization Procedures Macro");
				mp.inoutRng_.Text = "Randomization Procedures Macro: " + e.Message;
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

			
			bool isOther;

			string blindingType = icpInstMgr_.getTypedDisplayValue(DesignDefines.MaskingType, out isOther);
			string blindingOtherValue = null;
			if (isOther) 
			{
				blindingOtherValue = icpInstMgr_.getTypedOtherDisplayValue(DesignDefines.MaskingType, out isOther);
			}
	
			string randomizationType = icpInstMgr_.getTypedDisplayValue(DesignDefines.MethodOfAllocationType, out isOther);
			string randomizationOtherValue = null;
			if (isOther) 
			{
				randomizationOtherValue = icpInstMgr_.getTypedOtherDisplayValue(DesignDefines.MethodOfAllocationType, out isOther);
			}

			if (MacroBaseUtilities.isEmpty(randomizationType) )
			{
				wrkRng.InsertAfter("Study Randomization not defined.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
			else
			{
				if (randomizationType.Equals("Randomized")) 
				{
					wrkRng.InsertAfter("Subjects will be randomized into treatment groups.  ");
					wrkRng.InsertAfter("The subject randomization numbers will be generated ");
					wrkRng.InsertAfter("by Ogn Pharmaceutical or its designee and incorporated into ");

					if (MacroBaseUtilities.isEmpty(blindingType)) 
					{
						wrkRng.InsertAfter("Study Blinding not defined.");
					}
					else if (blindingType.Equals("Double-blinded")) 
					{
						wrkRng.InsertAfter("double-blind labeling.");
					}
					else if (blindingType.Equals("Single-blinded") ||
						blindingType.Equals("Open-label")) 
					{
						wrkRng.InsertAfter("a set of subject numbers and associated ");
						wrkRng.InsertAfter("treatment[s] which will be sent to the investigator.");
					}
					
					wrkRng.InsertParagraphAfter();
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				}
				else if (randomizationType.Equals("PhoneIn")) 
				{
					wrkRng.InsertAfter("Subjects will be randomized to treatment groups. ");
					wrkRng.InsertAfter("The subject randomization numbers will be generated ");
					wrkRng.InsertAfter("by Ogn Pharmaceutical or its designee and incorporated into a ");
					wrkRng.InsertAfter("set of subject numbers and associated treatment[s] ");
					wrkRng.InsertAfter("which will be given to the investigator over the ");
					wrkRng.InsertAfter("telephone at the time of individual subject ");
					wrkRng.InsertAfter("enrollment.");
				}
				else if (randomizationType.Equals("Unrandomized")) 
				{
					wrkRng.InsertAfter("Subjects will be assigned to treatment in ");
					wrkRng.InsertAfter("consultation with the sponsor.");
				}
				//else if (randomizationType.Equals("other")) 
				//{
				//}
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
