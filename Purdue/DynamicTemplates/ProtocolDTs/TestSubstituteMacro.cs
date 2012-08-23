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
	internal sealed class TestSubstituteMacro
	{
		private static readonly string header_ = @"$Header: TestSubstituteMacro.cs, 1, 18-Aug-09 12:06:00, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for TestSubstituteMacro.
	/// </summary>
	public class TestSubstituteMacro : AbstractMacroImpl
	{
		public TestSubstituteMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region TestSubstituteMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd TestSubstitute (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.TestSubstituteMacro.TestSubstitute,ProtocolDTs.dll" elementLabel="TestSubstitute" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="" autogenerates="true" toolTip="TestSubstitute." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("TestSubstitute Macro", "Generating information...");
				
				TestSubstituteMacro macro = null;
				macro = new TestSubstituteMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in TestSubstitute Macro"); 
				mp.inoutRng_.Text = "TestSubstitute Macro: " + e.Message;
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

			wrkRng.InsertAfter("Hello from TestSubstitute Macro");
			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			ChooserEntry template = IcdSchemaMgr.getTreatmentTemplate();
			IEnumerator ien = template.getComplexChildren();
			while (ien.MoveNext())
			{
				IChooserEntry entry = (IChooserEntry)ien.Current;
				string ep = entry.getElementPath();

				IChooserEntry mData = template.getMetaData(ep);
				if (mData != null && mData.isClientVariable() && ep.StartsWith("TMPLT")) 
				{
				}
			}


			wdDoc_.UndoClear();
		}

		public override void postProcess()
		{
			// Clean up memory
		}
	}
}
