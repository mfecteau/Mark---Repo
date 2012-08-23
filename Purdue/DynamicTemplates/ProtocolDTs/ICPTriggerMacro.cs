using System;
using System.IO;
using System.Collections;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;

using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;
using MSXML2;

namespace VersionControl 
{
	internal sealed class ICPTriggerMacro
	{
		private static readonly string header_ = @"$Header: ICPTriggerMacro.cs, 1, 18-Aug-09 12:04:17, Pinal Patel$";
	}
}

namespace TspdCfg.Roche.DynTmplts
{
	/// <summary>
	/// Summary description for ICPTriggerMacro.
	/// </summary>
	public class ICPTriggerMacro : AbstractMacroImpl
	{
		static ICPTriggerMacro()
		{
			// MessageBox.Show(WinApi.getForeGroundWindow(), "static ICPTriggerMacro()");
		}

		static Hashtable triggerCollection = new Hashtable();

		public ICPTriggerMacro(MacroExecutor.MacroParameters mp) : base (mp)
		{
			//
			// TODO: Add constructor logic here
			//
		}
		
		#region Dynamic Tmplt Methods
		
		#region ICPTriggerMacro
		/// <summary>
		/// Displays all inclusion criteria without category information
		/// </summary>
		/// <param name="mp"></param>
		/// <returns></returns>
		public static MacroExecutor.MacroRetCd ICPTrigger (
			MacroExecutor.MacroParameters mp) 
		{
#if false
<ChooserEntry elementPath="TspdCfg.Roche.DynTmplts.ICPTriggerMacro.ICPTrigger,ProtocolDTs.dll" elementLabel="ICPTrigger" ftElementType="Macro" ftMacroType="CSHARP" protected="true" editorClass="" autogenerates="true" toolTip="ICPTrigger." shouldRun="true"/>
#endif
			try 
			{
				mp.pba_.setOperation("ICPTrigger Macro", "Generating information...");
				
				ICPTriggerMacro macro = null;
				macro = new ICPTriggerMacro(mp);
				macro.preProcess();
				macro.display(); 
				macro.postProcess();
				return macro.macroStatusCode_;
			} 
			catch (Exception e) 
			{
				Log.exception(e, "Error in ICPTrigger Macro"); 
				mp.inoutRng_.Text = "ICPTrigger Macro: " + e.Message;
			}
			return MacroExecutor.MacroRetCd.Failed;
		}

		#endregion

		#endregion

		
		bool DUMP_INFO;
		ArrayList triggerPaths;
		string macroText;
		bool confirmReplace;
		bool peristValues;
		string targetNodes;

		public override void display()
		{
			Word.Range inoutRange = this.startAtBeginningOfParagraph();
			Word.Range wrkRng = inoutRange.Duplicate;

			// Load configuration
			loadConfiguration();

			pba_.updateProgress(3.0);

			string elementPath = execParms_.getParm(MacroExecutor.MacroExecParameters.ELEMENT_PATH);

			wrkRng.InsertAfter(macroText);
			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			// Load up the values for the trigger paths, and determine what changed since last run,
			// and persist the trigger values
			Hashtable htChanged;
			Hashtable htDoc = loadTriggers(out htChanged);

			// Dump debug info if configured
			dumpInfo(htDoc, htChanged, ref wrkRng);

			// Do the work
			doReplace(htDoc, htChanged, ref wrkRng);

			// All done
			// Set outgoing range
			inoutRange.End = wrkRng.End;
			setOutgoingRng(inoutRange);

			wdDoc_.UndoClear();

			icpRefMgr_.refreshAllDisplayValues();
		}

		private Hashtable loadTriggers(out Hashtable changed) 
		{
			long docID = tspdDoc_.getDocumentDetails().getId();
			changed = new Hashtable();

			// Get the ht for this doc, create if needed
			Hashtable htDoc = triggerCollection[docID] as Hashtable;
			if (htDoc == null) 
			{
				htDoc = new Hashtable();
				triggerCollection[docID] = htDoc;
			}

			// Read in the persisted trigger paths and values
			readPersistedTriggerValues(htDoc);


			// For each trigger XPath resolve the nodes
			foreach (string query in triggerPaths) 
			{
				try 
				{
					IXMLDOMNodeList result = icpInstMgr_.getRoot().selectNodes(query);
					foreach (IXMLDOMNode node in result) 
					{
						string sPath = getPath(node);

						// get old value
						string oldValue = htDoc[sPath] as string;
						string newValue = node.nodeTypedValue as string;

						// save old value
						if (!MacroBaseUtilities.isEmpty(oldValue) && oldValue != newValue) 
						{
							changed[sPath] = oldValue;
						}

						// update with new value
						htDoc[sPath] = newValue;
					}
				}
				catch (Exception ex) 
				{
					Log.exception(ex, "Error processing: " + query);
				}
			}

			// Read in the persisted trigger paths and values
			savePersistedTriggerValues(htDoc);

			return htDoc;
		}

		private void loadConfiguration() 
		{
			// Read configuration from triggers.txt
			triggerPaths = new ArrayList();
			macroText = "..";
			DUMP_INFO = false;
			confirmReplace = false;
			peristValues = false;
			targetNodes = "//*[@dataType='String']";

			string templatePath = tspdDoc_.getTrialProject().getTemplateDirPath();
			string triggerPathFile = templatePath + "\\" + "triggers.txt";

			string error = "";

			try 
			{
				error = " Opening: " + triggerPathFile;

				StreamReader str = File.OpenText(triggerPathFile);

				error = " Reading trigger paths";
				string s = "";
				while ((s = str.ReadLine()) != null) 
				{
					s = s.Trim();

					if (MacroBaseUtilities.isEmpty(s)) 
					{
						continue;
					}

					if (s.StartsWith("#"))
					{
						continue;
					}

					string cmd = "!text";
					if (s.StartsWith(cmd) && s.Length > cmd.Length)
					{
						string sText = s.Substring(cmd.Length).Trim();
						if (!MacroBaseUtilities.isEmpty(sText)) 
						{
							macroText = sText;
						}

						continue;
					}

					cmd = "!debug";
					if (s.StartsWith(cmd))
					{
						DUMP_INFO = true;
						continue;
					}

					cmd = "!confirmReplace";
					if (s.StartsWith(cmd))
					{
						confirmReplace = true;
						continue;
					}

					cmd = "!persistValues";
					if (s.StartsWith(cmd))
					{
						peristValues = true;
						continue;
					}

					cmd = "!targetNodes";
					if (s.StartsWith(cmd) && s.Length > cmd.Length)
					{
						string sText = s.Substring(cmd.Length).Trim();
						if (!MacroBaseUtilities.isEmpty(sText)) 
						{
							targetNodes = sText;
						}

						continue;
					}
				
					// Add the triggerpath
					triggerPaths.Add(s);
				}
			}
			catch (Exception ex) 
			{
				Log.exception(ex, "Error loading trigger configuration" + error);
			}

		}

		private void readPersistedTriggerValues(Hashtable htDoc) 
		{
			if (!peristValues) 
			{
				return;
			}

			Hashtable htPersisted = new Hashtable();
			Stream str = null;

			// Read persisted values
			try 
			{
				string trialPath = tspdDoc_.getTrialProject().getTrialDirPath();
				string triggerValueFile = trialPath + "\\" + "trigger.values";

				try 
				{
					str = File.OpenRead(triggerValueFile);
				}
				catch (Exception ex) 
				{
					return;
				}

				BinaryFormatter bf = new BinaryFormatter();

				htPersisted = bf.Deserialize(str) as Hashtable;
			}
			catch (Exception ex) 
			{
				Log.exception(ex, "Error reading persisted trigger values.");
			}
			finally
			{
				if (str != null) str.Close();
			}

			// Merge in the persisted values
			foreach (string triggerPath in htPersisted.Keys) 
			{
				// Verify that the persisted values still exist
				IXMLDOMNodeList icpElem = icpInstMgr_.getRoot().selectNodes(triggerPath);
				if (icpElem.length == 0) 
				{
					// remove from htDoc, does not exist in icp
					htDoc.Remove(triggerPath);

					continue;
				}

				string triggerValue = htPersisted[triggerPath] as string;

				if (!MacroBaseUtilities.isEmpty(triggerValue)) 
				{
					htDoc[triggerPath] = triggerValue;
				}
			}
		}

		private void savePersistedTriggerValues(Hashtable htDoc) 
		{
			if (!peristValues) 
			{
				return;
			}

			Stream str = null;

			try 
			{
				string trialPath = tspdDoc_.getTrialProject().getTrialDirPath();
				string triggerValueFile = trialPath + "\\" + "trigger.values";

				try 
				{
					str = File.OpenWrite(triggerValueFile);
				}
				catch (Exception ex) 
				{
					return;
				}

				BinaryFormatter bf = new BinaryFormatter();

				bf.Serialize(str, htDoc);
			}
			catch (Exception ex) 
			{
				Log.exception(ex, "Error writing persisted trigger values.");
			}
			finally
			{
				if (str != null) str.Close();
			}
		}

		private void doReplace(Hashtable htDoc, Hashtable htChanged, ref Word.Range wrkRng) 
		{
			IXMLDOMNodeList elementCollection = icpInstMgr_.getRoot().selectNodes(targetNodes);
			foreach (IXMLDOMNode elementNode in elementCollection) 
			{
				string elementPath = getPath(elementNode);
				string nodeValue = elementNode.nodeTypedValue as string;

				foreach (string changePath in htChanged.Keys) 
				{
					// Skip self
					if (elementPath == changePath) 
					{
						continue;
					}

					if (MacroBaseUtilities.isEmpty(nodeValue)) 
					{
						continue;
					}

					string oldValue = htChanged[changePath] as string;
					string newValue = htDoc[changePath] as string;

					// Don't replace or replace with empty values
					if (MacroBaseUtilities.isEmpty(oldValue) || MacroBaseUtilities.isEmpty(newValue))
					{
						continue;
					}

					string newNodeValue = nodeValue.Replace(oldValue, newValue);
					if (nodeValue != newNodeValue) 
					{
						if (confirmReplace) 
						{
							string label = "";

							string systemName = getSystemName(elementNode);
							if (!MacroBaseUtilities.isEmpty(systemName)) 
							{
								label = systemName;
							}

							string msg = "Do you wish to replace the value:                           \r\n";
							msg += "    \"" + oldValue + "\"\r\n";
							msg += "with: \r\n";
							msg += "    \"" + newValue + "\"\r\n\r\n";
							msg += "In the following element:\r\n";
							
							ArrayList fqn = icpRefMgr_.getFullyQualifiedName(elementPath, label);

							string indent = "";
							foreach (string s in fqn) 
							{
								msg += "\r\n" + indent + s;
								indent += ">";
							}

							DialogResult result = MessageBox.Show(WinApi.getForeGroundWindow(),
								msg,
								"Confirm Replacement", 
								MessageBoxButtons.YesNo, MessageBoxIcon.Question);

							if (result == DialogResult.No) 
							{
								continue;
							}
						}

						// Do the replacement
						elementNode.nodeTypedValue = newNodeValue;

						if (DUMP_INFO) 
						{
							wrkRng.InsertAfter("Replaced: " + elementPath);
							wrkRng.InsertAfter(", old: " + nodeValue);
							wrkRng.InsertAfter(", new: " + newNodeValue);
							wrkRng.InsertParagraphAfter();
							wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
							wdDoc_.UndoClear();
						}
					}
				}
			}
		}
		private string getPath(IXMLDOMNode node) 
		{
			string path = node.nodeName + addObjID(node);

			IXMLDOMNode parent = node.parentNode;
			while (parent != null && parent.nodeType == DOMNodeType.NODE_ELEMENT)
			{
				path = parent.nodeName + addObjID(parent) + "/" + path;
				parent = parent.parentNode;
			}

			return "/" + path;
		}

		private string getSystemName(IXMLDOMNode node) 
		{
			IXMLDOMNode systemName = node.attributes.getNamedItem(IcpDefines.SystemName);

			IXMLDOMNode parent = node.parentNode;
			while (systemName == null && parent != null && parent.nodeType == DOMNodeType.NODE_ELEMENT) 
			{
				systemName = parent.attributes.getNamedItem(IcpDefines.SystemName);
				parent = parent.parentNode;
			}
			
			if (systemName != null) 
			{
				return systemName.nodeTypedValue as string;
			}

			return null;
		}

		private string addObjID(IXMLDOMNode node) 
		{
			string ret = "";
			
			// If this node has an objID, append it
			IXMLDOMNode objID = node.attributes.getNamedItem(IcpDefines.ObjID);
			if (objID != null) 
			{
				ret = "[@" + IcpDefines.ObjID + "=" + objID.nodeValue + "]";
			}

			return ret;
		}

		private void dumpInfo(Hashtable htDoc, Hashtable htChanged, ref Word.Range wrkRng) 
		{
			if (!DUMP_INFO) return;

			if (htDoc.Keys.Count == 0) 
			{
				wrkRng.InsertAfter("There are no trigger paths defined or there are no matching trigger paths.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				wdDoc_.UndoClear();

				return;
			}

			wrkRng.InsertAfter("Loaded trigger paths:");
			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			foreach (string triggerPath in htDoc.Keys) 
			{
				string triggerValue = htDoc[triggerPath] as string;
				wrkRng.InsertAfter(triggerPath);
				wrkRng.InsertAfter(": " + triggerValue);
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			}

			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			wrkRng.InsertAfter("Target Nodes: ");
			wrkRng.InsertAfter(targetNodes);
			wrkRng.InsertParagraphAfter();
			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);


			if (htChanged.Keys.Count == 0) 
			{
				wrkRng.InsertAfter("There are no trigger value changes.");
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				wdDoc_.UndoClear();

				return;
			}
			
			wrkRng.InsertAfter("Changed trigger paths:");
			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);

			foreach (string changePath in htChanged.Keys) 
			{
				string oldValue = htChanged[changePath] as string;
				string newValue = htDoc[changePath] as string;
				wrkRng.InsertAfter(changePath);
				wrkRng.InsertAfter(", old: " + oldValue);
				wrkRng.InsertAfter(", new: " + newValue);
				wrkRng.InsertParagraphAfter();
				wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
				wdDoc_.UndoClear();
			}

			wrkRng.InsertParagraphAfter();
			wrkRng.Collapse(ref WordHelper.COLLAPSE_END);
			wdDoc_.UndoClear();
		}

		public override void postProcess()
		{
			// Clean up memory
		}
	}
}
