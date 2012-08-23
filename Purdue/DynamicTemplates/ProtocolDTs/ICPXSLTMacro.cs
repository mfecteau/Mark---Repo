using System;
using System.Collections;
using System.Windows.Forms;
using System.IO;

using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;
using MSXML2;
using Tspd.Bridge;

using Word = Microsoft.Office.Interop.Word;

namespace VersionControl 
{
	internal sealed class ICPXSLTMacro
	{
		private static readonly string header_ = @"$Header: ICPXSLTMacro.cs, 1, 18-Aug-09 12:04:18, Pinal Patel$";
	}
}

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for ICPXSLTMacro.
	/// </summary>
	public class ICPXSLTMacro : AbstractMacroImpl
	{
		LibraryElement _xsltElement = null;

		public ICPXSLTMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region ICPXSLT

		public static MacroExecutor.MacroRetCd ICPXSLT (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Purdue.DynTmplts.ICPXSLTMacro.ICPXSLT,ProtocolDTs.dll" elementLabel="ICPXSLT" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="" autogenerates="true" toolTip="ICPXSLT." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("ICPXSLT Macro", "Generating information...");
				
				ICPXSLTMacro macro = null;
				macro = new ICPXSLTMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in ICPXSLT Macro"); 
				mp.inoutRng_.Text = "ICPXSLT Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		#endregion

		#endregion

		// If this is a macro based on a fly out menu, check if valid
		public static new bool canRun(BaseProtocolObject bpo)
		{
			return true;
		}

		public override void preProcess()
		{
			#region parameter check
			// Get stored parameters
			string sParms = execParms_.getParm(MacroExecutor.MacroExecParameters.PARM_1);
			string[] aParms = null;

			// Split strings
			if (!MacroBaseUtilities.isEmpty(sParms)) 
			{
				aParms = sParms.Split('|');
			}

			// Peel apart the parameters
			bool parmsValid = true;
			string sLibraryBucket = null;
			string sLibraryItem = null;

			if (aParms != null && aParms.Length == 2)
			{
				if (!MacroBaseUtilities.isEmpty(aParms[0])) 
				{
					sLibraryBucket = aParms[0];
				}

				if (!MacroBaseUtilities.isEmpty(aParms[1])) 
				{
					sLibraryItem = aParms[1];
				}
			}
			else
			{
				parmsValid = false;
			}

			// We have parms, check if valid library bucket/item and text type
			if (parmsValid) 
			{
				LibraryManager lm = LibraryManager.getInstance();
				IEnumerator bucketEnum = lm.getLibraryBuckets();
				while (bucketEnum.MoveNext()) 
				{
					LibraryBucket bucket = (LibraryBucket )bucketEnum.Current;
					string bucketName = bucket.getBucketName();

					IEnumerator elementEnum = bucket.getElements().iterator();
					while (elementEnum.MoveNext()) 
					{
						LibraryElement libElement = (LibraryElement )elementEnum.Current;

						if (bucketName.Equals(sLibraryBucket) && 
								libElement.getElementName().Equals(sLibraryItem) &&
								libElement.getContentType() == LibraryContentType.TEXT) 
						{
							// found it
							_xsltElement = libElement;
							break;
						}
					}
				}
			}

			if (_xsltElement == null) 
			{
				parmsValid = false;
			}
			#endregion

			// Prompt for parms if they are missing or not valid
			if (!parmsValid) 
			{
				ICPXSLTSelect select = new ICPXSLTSelect();
				DialogResult res = select.ShowDialog();
				if (res == DialogResult.OK) 
				{
					_xsltElement = select.SelectedLibraryElement;
					sLibraryBucket = select.SelectedBucketName;
					sLibraryItem = select.SelectedLibraryItemName;

					// save it for next time so we don't ask
					sParms = sLibraryBucket + "|";
					sParms += sLibraryItem;

					execParms_.setParm(MacroExecutor.MacroExecParameters.PARM_1, sParms);
				}
			}
		}

		public override void display()
		{
			Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;

			pba_.updateProgress(50.0);

			if (_xsltElement == null) 
			{
				wrkRng.InsertAfter("A valid category/item was not selected.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

				// Set outgoing range
				inoutRange.End = wrkRng.End;
				setOutgoingRng(inoutRange);

				wdDoc_.UndoClear();
				return;
			}

			try 
			{
				// Make sure the library element is on disk
				string path = BridgeProxy.getInstance().loadLibraryElement(_xsltElement.getLibraryBucketID(), _xsltElement.getPKValue());

				// Load the fticp
				IXMLDOMDocument doc = icpInstMgr_.getRoot().ownerDocument;

				// Load the xsl
				IXMLDOMDocument xsl = new DOMDocument40Class();
				xsl.load(path);
			
				// Set outfile
				string outFile =  tspdDoc_.getTrialProject().getTrialDirPath() + "\\output.html";

				// Do transform
				string outText = doc.transformNode(xsl);

				// Write out the output
				FileStream stream = new FileStream(outFile, FileMode.Create);
				StreamWriter writer = new StreamWriter(stream);
				writer.Write(outText);
				writer.Flush();
				stream.Close();

				// Strange, needs a space to move the range?
				wrkRng.InsertAfter(" ");
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

				// The place to do the insert at
				Word.Range insertRange = wrkRng.Duplicate;

				// Move our range forward, why?
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

				// Do the insert
				object theRange = System.Reflection.Missing.Value;
				object confirm = false;
				object link = false;
				object attachment = System.Reflection.Missing.Value;
				insertRange.InsertFile(outFile, ref theRange, ref confirm, ref link, ref attachment);
			}
			catch (Exception ex) 
			{
				wrkRng.InsertAfter("Error running macro: " + ex.Message);
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}
		

			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();
		}

		public override void postProcess()
		{
			// Clean up memory
			_xsltElement = null;
		}
	}
}
