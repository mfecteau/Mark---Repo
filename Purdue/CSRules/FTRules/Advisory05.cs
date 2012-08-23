using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections;
using MSXML2;

using Tspd.Businessobject;
using Tspd.Tspddoc;
using Tspd.Icp;
using Tspd.Utilities;
using Tspd.Rules;
using Tspd.Context;
using WorX;
using Word = Microsoft.Office.Interop.Word;


namespace TspdCfg.FastTrack.Rules
{
	public class Advisory05 : ITSPDRule
	{
#if false
<rule id="Advisory05" type="CSHARP" displayName="Advisory05" source="TspdCfg.FastTrack.Rules.Advisory05.,FTRules.dll" categories="StyleAdv" debug="false"/>
#endif
		public static Hashtable Sel_Section = new Hashtable();
		public static readonly string MY_ID = "Advisory05";
		public ArrayList result = new ArrayList();
		private static string _ruleId = "";
		private static bool _debug = false;

		public void Init(string ruleId, bool debug) 
		{
			_ruleId = ruleId;
			_debug = debug;
		}

		public string AdvisoryPrefix
		{
			get { return MY_ID; }
		}


        public bool canRunInStandaloneDocument
        {
            get { return true; }

        }
		public ICollection Run() 
		{
			ArrayList advisories = new ArrayList();
			ArrayList docSec =  new ArrayList();

			ContextManager ctx = ContextManager.getInstance();
			TspdDocument doc = ctx.getActiveDocument();
			BusinessObjectMgr bom = doc.getBom();

			result.Clear();

			docSec = CheckBookMark(doc);
			
			docSec= CheckSize(doc);

			string smsg="" ;
		
			foreach (string str in docSec) 
            {
				 smsg = str;
			RuleAdvisory adv = new RuleAdvisory(
				_ruleId, MY_ID + str , smsg);
			advisories.Add(adv);
			}
			
			return advisories;
		}

	
		public ArrayList CheckBookMark(TspdDocument currDoc)
		{
				bool flag_found = false;
				IEnumerator bookmark =  currDoc.getActiveWordDocument().Bookmarks.GetEnumerator();

				while( bookmark.MoveNext())
				{
					Word.Bookmark bk =  bookmark.Current as Word.Bookmark;

					if (bk.Name.StartsWith("_Toc"))
					{
						flag_found= true;
						break;
					}
				}
				
				if (flag_found == false)
				{
					result.Add("Protocol Document must have atleast one Bookmark pointing to TOC");
				}

				return result;
			}
			
		public ArrayList CheckSize(TspdDocument currDoc)
		{
		//	ArrayList size = new ArrayList();
			string filepath="";

			filepath = currDoc.getTrialProject().getTrialDirPath() + "\\" +  currDoc.getDocumentDetails().getRelativeFileName();
			
			System.IO.FileInfo obj =  new System.IO.FileInfo(filepath);
			
			if (System.IO.File.Exists(filepath))
			{
				if (obj.Length > 104857600)
				{
					result.Add("File Size is greater than 100 MB");
				}
			}
			return result;
		}

//		public ArrayList HasPassword(TspdDocument currDoc)
//		{
//			if (currDoc.getActiveWordDocument().HasPassword)
//			{
//               // Word.WdUnits.
//			}
//		}

		

//		public ArrayList BlankPage(TspdDocument currDoc)
//		{
//			
//			Word.Document wdoc = currDoc.getActiveWordDocument();
//
//			int numpages = wdoc.ActiveWindow.Panes.Item(1).Pages.Count;
//
//			Word.Page pg = wdoc.ActiveWindow.Panes(1).Pages(1);
//
//
//
////			Word.Document doc = currDoc.getActiveWordDocument();
////			Word.Window dsd;
////			
////			object obj = currDoc.getActiveWordDocument().ActiveWindow.Panes;
////			Word.Panes p1 = currDoc.getActiveWordDocument().ActiveWindow.Panes;
////
////
////			Word.Range locRng = doc.Range;
////			locRng.SetRange(0,0);
////
////			//locRng.Move(
////
////			doc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
////
////			//currDoc.getActiveWordDocument().
//		}

	}



	}

