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

namespace VersionControl 
{
	internal sealed class SOATableMacro
	{
		private static readonly string header_ = @"$Header: SOATableMacro.cs, 1, 18-Aug-09 12:05:41, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts.Table
{
	/// <summary>
	/// Summary description for SOATableMacro.
	/// </summary>
	public class SOATableMacro
	{
		public SOATableMacro()
		{
		}
		
		#region Dynamic Tmplt Methods
		
		#region SOATableMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd SOATable (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Roche.DynTmplts.Table.SOATableMacro.SOATable,ProtocolDTs.dll" elementLabel="Study Schedule" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.Schedule" autogenerates="true" toolTip="Creates Study Schedule." shouldRun="true">
	<Complex>
		<ChooserEntry ftElementType="Collection" assocClass="Tspd.Icp.SOA,IcpMgr" elementPath="dummy" elementLabel="dummy" assocChooserPath="/FTICP/StudySchedule/Schedules/Schedule"/>
	</Complex>
</ChooserEntry>

#endif
			try 
			{
				mp.pba_.setOperation("Study Schedule Macro", "Generating information...");
				
				DefSOATableDisplayMgr macro = null;
				macro = new PurdueSOATableDisplayMgr(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.MacroStatusCode;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in Study Schedule Macro"); 
				mp.inoutRng_.Text = "Study Schedule Macro: " + e.Message;
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
