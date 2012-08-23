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

using Word = Microsoft.Office.Interop.Word;

namespace VersionControl 
{
	internal sealed class LinkViewerMacro
	{
		private static readonly string header_ = @"$Header: LinkViewerMacro.cs, 1, 18-Aug-09 12:04:43, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for PregnancyMacro.
	/// </summary>
	public class LinkViewerMacro : AbstractMacroImpl
	{
		public LinkViewerMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region LinkViewerMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd LinkViewer(
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.LinkViewerMacro.LinkViewer,ProtocolDTs.dll" elementLabel="Library Item" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.LibraryItem" autogenerates="true" toolTip="Library Item." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Link Viewer Macro", "Generating information...");
				
				LinkViewerMacro macro = null;
				macro = new LinkViewerMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in LinkViewer Macro");
				mp.inoutRng_.Text = "LinkViewer Macro: " + e.Message;
			}
			return  MacroExecutor.MacroRetCd.Failed;
		}

		#endregion

		#endregion

		public override void display()
		{
			string str="";
			Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;

			//pba_.updateProgress(1.0);

			string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);


			bool isOther;

			// Get stored parameters
			string sParms = execParms_.getParm(MacroExecutor.MacroExecParameters.PARM_1);
			string aParms = null;
				
				frmLinkageViewer  lItem = new frmLinkageViewer();
				lItem.Load_Data(tspdDoc_);


				macroStatusCode_ = MacroExecutor.MacroRetCd.Failed;
			return;


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
