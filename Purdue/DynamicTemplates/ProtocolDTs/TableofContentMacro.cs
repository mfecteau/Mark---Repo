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
	internal sealed class TableofContentMacro
    {
		private static readonly string header_ = @"$Header: TableofContentMacro.cs, 1, 18-Aug-09 12:05:56, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for PregnancyMacro.
	/// </summary>
	public class TableofContentMacro: AbstractMacroImpl
	{
		public TableofContentMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region TableofContentMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd TOCUpdate (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.TableofContentMacro.TOCUpdate,ProtocolDTs.dll" elementLabel="Update TOC" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.LibraryItem" autogenerates="true" toolTip="Library Item." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Study Conduct Macro", "Generating information...");
				
				TableofContentMacro macro = null;
				macro = new TableofContentMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Table of Content Macro");
				mp.inoutRng_.Text = "Table of Content Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		#endregion

		#endregion

		public override void display()
		{
			string str="";
			Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;

			pba_.updateProgress(1.0);

			string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);

			Word.Field TOC = null;

		
			IEnumerator allTOC = tspdDoc_.getActiveWordDocument().TablesOfContents.GetEnumerator();

			while (allTOC.MoveNext())
			{
				Word.TableOfContents TOC_ = (Word.TableOfContents)allTOC.Current;
			//	TOC_ = tspdDoc_.getActiveWordDocument().TablesOfContents;
				TOC_.Update();

			}
			
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
