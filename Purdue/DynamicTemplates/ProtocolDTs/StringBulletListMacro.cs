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
	internal sealed class StringBulletListMacro
	{
		private static readonly string header_ = @"$Header: StringBulletListMacro.cs, 1, 18-Aug-09 12:05:46, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for StringBulletListMacro.
	/// </summary>
	public class StringBulletListMacro : AbstractMacroImpl
	{
		public StringBulletListMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}

		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd StringList (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.StringBulletListMacro.StringList,ProtocolDTs.dll" elementLabel="StringList" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="RES_STRINGLIST_EDIT" autogenerates="true" toolTip="" shouldRun="true" hidden="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("String Bullet List Macro", "Generating information...");
				
				StringBulletListMacro macro = null;
				macro = new StringBulletListMacro(mp);
				macro.preProcess();
				macro.display(); 
				//macro.postprocess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				mp.inoutRng_.Text = "StringBulletListMacro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		public override void display()
		{
			Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;

			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			pba_.updateProgress(50.0);

			string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);

			if (MacroBaseUtilities.isEmpty(elementPath)) 
			{
				wrkRng.InsertAfter("StringList elementpath was not passed in.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
			else
			{
				object oStyle = null;
				StringListHelper slist = null;

				string listStyle = execParms_.getParm(MacroExecutor.MacroExecParameters.PARM_1);

				// Check for saved list style
				if (MacroBaseUtilities.isEmpty(listStyle)) 
				{
					listStyle = PurdueUtil.PFIZER_STYLE_TEXT_BULL;

					Word.Style style = wrkRng.get_Style() as Word.Style;

					if (style != null) 
					{
						try 
						{
							string curStyle = style.NameLocal;
							if (curStyle.Equals(PurdueUtil.PFIZER_STYLE_TABLETEXT_10)) 
							{
								listStyle = PurdueUtil.PFIZER_STYLE_TABLETEXT_BULL_10;
							}
						}
						catch (Exception ex) {}
					}

					// Save it
					execParms_.setParm(MacroExecutor.MacroExecParameters.PARM_1, listStyle);
				}
				
				try 
				{
					IXMLDOMNode node = icpInstMgr_.lookupNamedNode(elementPath);
					if (node == null) 
					{
						throw new Exception("stringlist no longer exists");
					}

					slist = bom_.getIcp().getStringList(elementPath, this.tspdDoc_.getDocType());
					ArrayList list = slist.toArray();

					// Display the stringlist elements
					if (list.Count > 0) 
					{
						if (listStyle != null) 
						{
							oStyle = tspdDoc_.getStyleHelper().setNamedStyle(listStyle, wrkRng);
						}

						for (IEnumerator iter = list.GetEnumerator(); iter.MoveNext(); ) 
						{
							string sCurrent = iter.Current.ToString();

							// Convert newlines to returns
							sCurrent = sCurrent.Replace("\n", "\v");

							wrkRng.InsertAfter(sCurrent);
							wrkRng.InsertParagraphAfter();
							wdDoc_.UndoClear();
						}

						wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

						if (listStyle != null) 
						{
							oStyle = tspdDoc_.getStyleHelper().setNamedStyle(PurdueUtil.NORMAL, wrkRng);
						}
					} 
					else 
					{
						wrkRng.InsertAfter("No values");
						wrkRng.InsertParagraphAfter();
					}
					
					//wrkRng.InsertParagraphAfter();
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				}
				catch (Exception ex)
				{
					wrkRng.InsertAfter("The StringList that this macro refers to was removed, delete this macro.");
					wrkRng.InsertParagraphAfter();
					wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				}
			}

			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}
	}
}
