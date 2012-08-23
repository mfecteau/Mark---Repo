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
	internal sealed class ARConventionMacro
	{
		private static readonly string header_ = @"$Header: ARConventionMacro.cs, 1, 18-Aug-09 12:02:37, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for A&R Convention Macro.
	/// </summary>
	public class ARConventionMacro : AbstractMacroImpl
	{
		public ARConventionMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		public string  strLibItemCode ="";
		
		#region ARConventionMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd Convention1(
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.ARConventionMacro.Convention1,ProtocolDTs.dll" elementLabel="Library Item" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.LibraryItem" autogenerates="true" toolTip="Library Item." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Library Item Macro", "Generating information...");
				
				ARConventionMacro macro = null;
				macro = new ARConventionMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in A&R Convention Macro");
				mp.inoutRng_.Text = "A&R Convention Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		public static MacroExecutor.MacroRetCd Lab(
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.ARConventionMacro.Lab,ProtocolDTs.dll" elementLabel="Library Item" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.LibraryItem" autogenerates="true" toolTip="Library Item." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Library Item Macro", "Generating information...");
				
				ARConventionMacro macro = null;
				macro = new ARConventionMacro(mp);
				macro.strLibItemCode = "LAB";
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in A&R Convention Macro");
				mp.inoutRng_.Text = "A&R Convention Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		public static MacroExecutor.MacroRetCd Vitals(
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.ARConventionMacro.Vitals,ProtocolDTs.dll" elementLabel="Library Item" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.LibraryItem" autogenerates="true" toolTip="Library Item." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Library Item Macro", "Generating information...");
				
				ARConventionMacro macro = null;
				macro = new ARConventionMacro(mp);
				macro.strLibItemCode = "VITALS";
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in A&R Convention Macro");
				mp.inoutRng_.Text = "A&R Convention Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		public static MacroExecutor.MacroRetCd AE(
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.ARConventionMacro.AE,ProtocolDTs.dll" elementLabel="Library Item" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.LibraryItem" autogenerates="true" toolTip="Library Item." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Library Item Macro", "Generating information...");
				
				ARConventionMacro macro = null;
				macro = new ARConventionMacro(mp);
				macro.preProcess();
				macro.strLibItemCode = "AE";
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in A&R Convention Macro");
				mp.inoutRng_.Text = "A&R Convention Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}


		public static MacroExecutor.MacroRetCd Other(
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.ARConventionMacro.Other,ProtocolDTs.dll" elementLabel="Library Item" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.LibraryItem" autogenerates="true" toolTip="Library Item." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Library Item Macro", "Generating information...");
				
				ARConventionMacro macro = null;
				macro = new ARConventionMacro(mp);
				macro.preProcess();
				macro.strLibItemCode = "OTHER";
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in A&R Convention Macro");
				mp.inoutRng_.Text = "A&R Convention Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		
		public static MacroExecutor.MacroRetCd LibraryItem(
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.ARConventionMacro.LibraryItem,ProtocolDTs.dll" elementLabel="Library Item" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.LibraryItem" autogenerates="true" toolTip="Library Item." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Library Item Macro", "Generating information...");
				
				ARConventionMacro macro = null;
				macro = new ARConventionMacro(mp);
				macro.preProcess();
				macro.strLibItemCode = "Library Items";
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in A&R Convention Macro");
				mp.inoutRng_.Text = "A&R Convention Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}
		#endregion

		#endregion

		public override void display()
		{
			string str="";
			ArrayList selLibItems = new ArrayList();
			Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;

			pba_.updateProgress(1.0);

			string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);

//			if (MacroBaseUtilities.isEmpty(elementPath)) 
//			{
//				return;
//			}


			bool isOther;

			// Ask the user if the parms are missing/invalid
		
				ARConvention  lItem = new ARConvention();
				Word.Bookmarks bk =  tspdDoc_.getActiveWordDocument().Bookmarks;
				lItem.Load_Form(bk,strLibItemCode);

				System.Windows.Forms.DialogResult res = lItem.ShowDialog();
					selLibItems = lItem.selList;    //Get all selected Library Items.
				if ( res == System.Windows.Forms.DialogResult.OK)
				{
					str = "";

					for (int i =0; i < selLibItems.Count;i++)
					{
						Word.Range myRange = wrkRng.Duplicate;;
						myRange.Start = inoutRange.End;
																		
						LibraryElement libElement = (LibraryElement) selLibItems[i];
						str = libElement.getElementName();
						

						//Inserting Library Item.
						tspdDoc_.insertLibraryItemByName(str, wrkRng);
						
						wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

						//execParms_.setParm(MacroExecutor.MacroExecParameters.PARM_1, str);

						myRange.End= wrkRng.End;
						object itemRng = myRange;	
		
						if (str.IndexOf("]") >= 0)
						{
							string bkName =  getbookMark(str);
							Word.Bookmark bm = tspdDoc_.getActiveWordDocument().Bookmarks.Add(bkName,ref itemRng);
						}
						// Set outgoing range
						
						
						inoutRange.End = wrkRng.End;
						//setmyBackColor(inoutRange);
						setmyBackColor(myRange);
                        setOutgoingRng(inoutRange);
						Word.Range pkRange = wrkRng.Duplicate;						
						wrkRng.InsertParagraph();
						pkRange.End = wrkRng.End;
						setmyBackColor1(pkRange);
						wrkRng.Start = pkRange.End;
						inoutRange.End = pkRange.End;
						wdDoc_.UndoClear();
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


		
		public void setmyBackColor(Word.Range mycolorRange)
		{
			mycolorRange.Select();
			mycolorRange.Shading.BackgroundPatternColor= Word.WdColor.wdColorTan;
		}
		
		public void setmyBackColor1(Word.Range mycolorRange)
		{
			mycolorRange.Select();
			mycolorRange.Shading.BackgroundPatternColor= Word.WdColor.wdColorWhite;

		}

		public string getbookMark(string bookmarkName)
		{
			string strName =null;
			string[] bkname;
			if (bookmarkName.IndexOf("-") >= 0)
			{
				 bkname = bookmarkName.Split('-');
				 strName = bkname[0];
			}
			else
			{
				bkname  = bookmarkName.Split(']');   // Assuming that it will have this code [UID: BOokmark Code] Lib_item_name.
				strName = bkname[0];			

			}
			


			string[] bkname2 = strName.Split(':');
			string bookMark = bkname2[1];

			//Check if Bookmark exists in document.
			if (tspdDoc_.getActiveWordDocument().Bookmarks.Exists(bookMark) == true)
			{
			 for(int i =0; i <=250;i++)
				{
					
				string strTemp	= bookMark + "_" + i;
				
					if (tspdDoc_.getActiveWordDocument().Bookmarks.Exists(strTemp) == false)
					{
						return strTemp;
					}
				}
			}

			return bookMark;

		}

		public override void postProcess()
		{
			// Clean up memory
		}
	}
}
