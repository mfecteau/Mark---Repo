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
	internal sealed class TemplateMacro
	{
		private static readonly string header_ = @"$Header: TemplateMacro.cs, 1, 18-Aug-09 12:05:58, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for TemplateMacro.
	/// </summary>
	public class TemplateMacro : AbstractMacroImpl
	{
		public TemplateMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region Template

		public static MacroExecutor.MacroRetCd Template (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.TemplateMacro.Template,ProtocolDTs.dll" elementLabel="Template" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="" autogenerates="true" toolTip="Template." shouldRun="true"/>

or

<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.TemplateMacro.Template,ProtocolDTs.dll" elementLabel="Template" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="Template" shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>
#endif
			try 
			{
				mp.pba_.setOperation("Template Macro", "Generating information...");
				
				TemplateMacro macro = null;
				macro = new TemplateMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Template Macro"); 
				mp.inoutRng_.Text = "Template Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		#endregion

		#endregion

		// If this is a macro based on a fly out menu, check if valid
		public static new bool canRun(BaseProtocolObject bpo)
		{
			// Example for SOA
			/*
			SOA soa = bpo as SOA;
			if (soa == null)
			{
				return false;
			}

			// Example of further restriction
			if (soa.isSchemaDesignMode()) 
			{
				return false;
			}
			*/

			return true;
		}

		public override void preProcess()
		{
		}

		public override void display()
		{
			Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;

			pba_.updateProgress(1.0);

			string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);

			wrkRng.InsertAfter("Hello from Template Macro");
			wrkRng.InsertParagraphAfter();
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
