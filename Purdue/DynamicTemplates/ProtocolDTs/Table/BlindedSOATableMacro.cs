using System;
using System.Collections;
using System.Diagnostics;
using System.Text;
using System.Runtime.InteropServices;

using Tspd.Tspddoc;
using Tspd.MacroBase.BaseImpl;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;
using Tspd.FormBase.ProgressBar;
using Tspd.Context;

using System.Windows.Forms;

namespace VersionControl 
{
	internal sealed class BlindedSOATableMacro
	{
		private static readonly string header_ = @"$Header: BlindedSOATableMacro.cs, 1, 18-Aug-09 12:02:44, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts.Table
{
	/// <summary>
	/// Summary description for SOATableMacro.
	/// </summary>
	public class BlindedSOATableMacro
	{
		public BlindedSOATableMacro()
		{
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region BlindedSOATableMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd BlindedSOATable (
			MacroExecutor.MacroParameters mp)
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.Table.BlindedSOATableMacro.BlindedSOATable,ProtocolDTs.dll" elementLabel="Blinded Schedule of Assessments" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="Creates Study Schedule." shouldRun="true">
	<Complex>
<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>

#endif
			try 
			{
				mp.pba_.setOperation("Blinded Study Schedule Macro", "Generating information...");
				
				PurdueSOATableDisplayMgr macro = new PurdueSOATableDisplayMgr(mp);
				macro.BlindedStudy = true;
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.MacroStatusCode;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Blinded Study Schedule Macro"); 
				mp.inoutRng_.Text = "Blinded Study Schedule Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		public static bool canRun(BaseProtocolObject bpo)
		{
			SOA soa = bpo as SOA;

			if (soa != null)
			{
				if (soa.isSchemaDesignMode()) 
				{
					BusinessObjectMgr bom = ContextManager.getInstance().getActiveDocument().getBom();
					IList armList = bom.getArmsForAssociatedSchedule(soa).getList();
					if (armList.Count == 0) 
					{
						return false;
					}
				}

				return true;
			}

			return false;
		}
		
		#endregion

		#endregion


	}
}
