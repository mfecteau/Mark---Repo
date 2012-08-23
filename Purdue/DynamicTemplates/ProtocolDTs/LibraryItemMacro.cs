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
	internal sealed class LibraryItemMacro
	{
		private static readonly string header_ = @"$Header: LibraryItemMacro.cs, 1, 18-Aug-09 12:04:43, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for PregnancyMacro.
	/// </summary>
	public class LibraryItemMacro : AbstractMacroImpl
	{
		public LibraryItemMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region LibraryItemMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd LibraryItem (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.LibraryItemMacro.LibraryItem,ProtocolDTs.dll" elementLabel="Library Item" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.LibraryItem" autogenerates="true" toolTip="Library Item." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Library Item Macro", "Generating information...");
				
				LibraryItemMacro macro = null;
				macro = new LibraryItemMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in LibraryItem Macro");
				mp.inoutRng_.Text = "LibraryItem Macro: " + e.Message;
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

//			if (MacroBaseUtilities.isEmpty(elementPath)) 
//			{
//				return;
//			}


			bool isOther;

			// Get stored parameters
			string sParms = execParms_.getParm(MacroExecutor.MacroExecParameters.PARM_1);
			string aParms = null;

			if (!MacroBaseUtilities.isEmpty(sParms)) 
			{
				aParms = sParms;
			}

			bool parmsValid = false;

			if (aParms != null && aParms.Length >=0)
			{
				parmsValid = true;

				if (!MacroBaseUtilities.isEmpty(aParms)) 
				{
					try 
					{
						str = aParms;
					}
					catch (Exception ex) 
					{
						parmsValid = false;
					}
				}

//				if (!MacroBaseUtilities.isEmpty(aParms[1])) 
//				{
//					try 
//					{
//						_includeScheduledTimes = bool.Parse(aParms[1]);
//					}
//					catch (Exception ex) 
//					{
//						parmsValid = false;
//					}
//				}
			}

			// Ask the user if the parms are missing/invalid
			if (!parmsValid) 
			{
				
				LibraryItem  lItem = new LibraryItem();
				lItem.loadLibraryItems();
				System.Windows.Forms.DialogResult res = lItem.ShowDialog();
			
				if ( res == System.Windows.Forms.DialogResult.OK)
				{
					str = lItem.SelectedItem.ToString();
					if (str.Length<=0)
					{
						wrkRng.InsertAfter("No Library Items Found!");
						wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				
					}
					else
					{
                        Log.trace(str + " Before Insert: " + wrkRng.Start.ToString() + " ---> " + wrkRng.End.ToString());
					//	tspdDoc_.insertLibraryItemByName(str, wrkRng);
                        //Regular Library Item, With NO Placeholders
                        Word.Range tmpRng = tspdDoc_.insertLibraryItemByNameNonInteractive(str, wrkRng);

                        tmpRng.Start = tmpRng.Start - 1;
                        if (tmpRng.Text == "\r")  //Remove extra paragraph markers.
                        {
                            tmpRng.Text = tmpRng.Text.Replace("\r", "");
                            // tmpRng.Start = tmpRng.Start;
                            tmpRng.Collapse(ref WordHelper.COLLAPSE_START);
                        }

                        wrkRng.Start = tmpRng.Start;
                        wrkRng.End = tmpRng.End;

                        wrkRng.Collapse(ref WordHelper.COLLAPSE_END);   //Collapsing range, incase.
                        Log.trace(str + " After Insert: " + wrkRng.Start.ToString() + " ---> " + wrkRng.End.ToString());
						wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
						execParms_.setParm(MacroExecutor.MacroExecParameters.PARM_1, str);
					}
				}
				else

				{
					try
					{
						macroStatusCode_ = MacroExecutor.MacroRetCd.Failed;
						return;
					}
					catch(Exception ex)
					{
						macroStatusCode_ = MacroExecutor.MacroRetCd.Failed;
						return;
					}
				}
				
			}
			else
			{
				//If Paramaeter are present, use it directly in here.
                Log.trace(str + " Before Insert: " + wrkRng.Start.ToString() + " ---> " + wrkRng.End.ToString());

                Word.Range tmpRng = tspdDoc_.insertLibraryItemByNameNonInteractive(str, wrkRng);

                tmpRng.Start = tmpRng.Start - 1;
                if (tmpRng.Text == "\r")  //Remove extra paragraph markers.
                {
                    tmpRng.Text = tmpRng.Text.Replace("\r", "");
                    // tmpRng.Start = tmpRng.Start;
                    tmpRng.Collapse(ref WordHelper.COLLAPSE_START);
                }

                wrkRng.Start = tmpRng.Start;
                wrkRng.End = tmpRng.End;

                //wrkRng.Collapse(ref WordHelper.COLLAPSE_END);   //Collapsing range, incase.

             //   tspdDoc_.insertLibraryItemByName(str, wrkRng);
                Log.trace(str + " After Insert: " + wrkRng.Start.ToString() + " ---> " + wrkRng.End.ToString());
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
 
	
		// Set outgoing range
		inoutRange.End = wrkRng.End;
        Log.trace(str + " IN-OUT RANGE: " + inoutRange.Start.ToString() + " ---> " + inoutRange.End.ToString());
		setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}

       
		public override void postProcess()
		{
			// Clean up memory
		}
	}
}
