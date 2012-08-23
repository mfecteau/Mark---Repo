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
	internal sealed class StudyConductMacro
	{
		private static readonly string header_ = @"$Header: StudyConductMacro.cs, 1, 18-Aug-09 12:05:47, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for PregnancyMacro.
	/// </summary>
	public class StudyConductMacro : AbstractMacroImpl
	{
		public StudyConductMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region StudyConductMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd StudyConduct(
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.StudyConductMacro.StudyConduct,ProtocolDTs.dll" elementLabel="Study Conduct" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="PDG.LibraryItem" autogenerates="true" toolTip="Library Item." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("Study Conduct Macro", "Generating information...");
				
				StudyConductMacro macro = null;
				macro = new StudyConductMacro(mp);
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
			string bucket = null;
			string chooserEnt = null;
			ArrayList stored_value;
			ArrayList studycondList = new ArrayList();
			int i = 0;
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
			string[] aParms = null;

			if (!MacroBaseUtilities.isEmpty(sParms)) 
			{
				aParms = sParms.Split('|');
			}

			bool parmsValid = false;

			if (aParms != null  && aParms.Length == 3)
			{
				parmsValid = true;

				if (!MacroBaseUtilities.isEmpty(aParms[0])) 
				{
					try 
					{
						bucket = aParms[0];
					}
					catch (Exception ex) 
					{
						parmsValid = false;
					}
				}

				if (!MacroBaseUtilities.isEmpty(aParms[1])) 
				{
					try 
					{
						chooserEnt = aParms[1];
					}
					catch (Exception ex) 
					{
						parmsValid = false;
					}
				}

				if (!MacroBaseUtilities.isEmpty(aParms[2])) 
				{
					try 
					{
						//stored_value= new ArrayList();
						string[] strTemp = aParms[2].Split('~');
						
						for(i = 0;i < strTemp.Length-1; i++)
						{
							studycondList.Add(strTemp[i]);
						}

						if (validation(bucket,chooserEnt,studycondList) == false)
						{
							 parmsValid = false;
						}

					}
					catch (Exception ex) 
					{
						parmsValid = false;
					}
				}
			}

			// Ask the user if the parms are missing/invalid
			if (!parmsValid) 
			{
				string bucketname,entryname;
				string sel_values = null;
				StudyConductSel  SItem = new StudyConductSel();
				SItem.loadStudyconducts();
				System.Windows.Forms.DialogResult res = SItem.ShowDialog();
			
				if ( res == System.Windows.Forms.DialogResult.OK)
				{
					studycondList = SItem.SelectedItems;
					bucketname = SItem.bucket_name;
					entryname =SItem.chooser_name;


					if (studycondList.Count<=0)
					{
						wrkRng.InsertAfter("No Items were selected!");
						wrkRng.Collapse(ref WordHelper.COLLAPSE_END);				
					}
					else
					{
						for(i =0; i < studycondList.Count;i++)
						{
						//	wrkRng.InsertAfter(studycondList[i].ToString());
							wrkRng.InsertAfter(studycondList[i].ToString());
							//MacroBaseUtilities.putElemRef(tspdDoc_,studycondList[i].ToString(),wrkRng,macroEntry_);
							sel_values = sel_values + studycondList[i].ToString() + "~";
							wrkRng.InsertParagraphAfter();
							
						}
						execParms_.setParm(MacroExecutor.MacroExecParameters.PARM_1, bucketname + "|" + entryname + "|" + sel_values);
						wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
					}
				}
				else
				{
					macroStatusCode_ = MacroExecutor.MacroRetCd.Failed;
					return;
				}
				
			}
			else
			{
				//If Paramaeter are present, use it directly in here.
			//	tspdDoc_.insertLibraryItemByName(str, wrkRng);

				for(i =0; i < studycondList.Count;i++)
				{
					wrkRng.InsertAfter(studycondList[i].ToString());
					//MacroBaseUtilities.putElemRef(tspdDoc_,studycondList[i].ToString(),wrkRng,macroEntry_);
					wrkRng.InsertParagraphAfter();
				}
				//wrkRng.InsertAfter(str);
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
 
	
		// Set outgoing range
		inoutRange.End = wrkRng.End;
		setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}
		public bool validation(string bucket,string cEntry,ArrayList strValue)
		{
			BucketEntry Curr_BEntry = null;
			BucketEntry bucketEntry = bom_.getIcd().getBucketBySystemName(ElementType.STUDY_CONDUCT_COMPLIANCE);
			BucketEntry bucketEntry1 = bom_.getIcd().getBucketBySystemName(ElementType.STUDY_CONDUCT_TERMINATION);

			///Compare the bucketlabel 
			if ( bucketEntry.getBucketLabel() == bucket)
			{ 
				Curr_BEntry = bucketEntry;
			}
			if(bucketEntry1.getBucketLabel() == bucket)
			{
				Curr_BEntry = bucketEntry1;
			}

			if (Curr_BEntry == null)
			{
				return false;
			}
 
///GEt the chooserEntry and compare its label to cEntry

			//BucketEntry entry1 = (BucketEntry)bucketname[comboBox2.SelectedIndex];
			IEnumerator Curr_CEntry = bom_.getIcd().getChooserEntriesForBucketEntry(Curr_BEntry);
			ChooserEntry ce = null;
			while (Curr_CEntry.MoveNext())
			{
				ce = (ChooserEntry)Curr_CEntry.Current;
				
				if (ce.getActualDisplayValue()== cEntry)
				{
					break;
				}
				
			}

			if(Curr_CEntry == null)
			{
				return false;
			}
			///get string list and compare it to strvalue.

			StringListHelper sh = bom_.getIcp().getStringList(ce.getElementPath(),tspdDoc_.getDocType());
			ArrayList sh1 =  new ArrayList();
			sh1 = bom_.getIcp().getStringListValues(ce.getElementPath());

			if (sh1.Count == 0)
			{
				return false;
			}

			int i = 0;

			if (strValue.Count == 0  || sh1.Count == 0)
			{
				return false;
			}
			
			for(i = 0; i<strValue.Count;i++)
			{
				if (sh1.IndexOf(strValue[i])< 0)
				{
					return false;
				}
			}

			return true;


		}
		public override void postProcess()
		{
			// Clean up memory
		}
	}
}
