using System;
using System.Collections;
using Tspd.Businessobject;
using Tspd.MacroBase;
using Tspd.MacroBase.Table;
using Tspd.Tspddoc;
using Tspd.Icp;
using Tspd.Utilities;
using Word = Microsoft.Office.Interop.Word;

namespace TspdCfg.Purdue.DynTmplts.Table
{
	/// <summary>
	/// Summary description for RocheFootnoter.
	/// </summary>
	public class PurdueFootnoter : IFootnoter
	{
		Hashtable _footNotes = new Hashtable();
		private BusinessObjectMgr bom_;
		private Word.Document wdDoc_;

		public PurdueFootnoter(BusinessObjectMgr bom, Word.Document wdDoc) 
		{
			bom_ = bom;
			wdDoc_ = wdDoc;
		}

		public bool hasFootnotes() 
		{
			return _footNotes.Count > 0;
		}

		public Hashtable getFootnotes() 
		{
			return _footNotes;
		}

		public bool putAtRng(SOAObject soaObj, Word.Range rng) 
		{
			IList fnList =  bom_.getFootnotes(soaObj);
			if (fnList.Count == 0) 
			{
				return false;
			}

			Word.Range insRng = rng.Duplicate;
			Word.Range mvRng = insRng.Duplicate;
			mvRng.Start = mvRng.End + 1;

			bool first = true;
			foreach (SOAObject.FootNote fn in fnList) 
			{
				FootNoteWrapper fnw = _footNotes[fn.getFootNoteID()] as FootNoteWrapper;
				if (fnw == null) 
				{
					fnw = new FootNoteWrapper();
					fnw.footNoteNumber = _footNotes.Count + 1;
					fnw.footNoteNumberString = translateFootnoteNumber(fnw.footNoteNumber);
					// fnw.footNoteNumberString = fnw.footNoteNumber.ToString();
					fnw.footNote = fn;

					_footNotes.Add(fn.getFootNoteID(), fnw);
				}

				if (!first) 
				{
					insRng.InsertAfter(IcpReferenceManager.NBSPACE);
					insRng.Collapse(ref WordHelper.COLLAPSE_END);
				}

				insRng.InsertAfter(fnw.footNoteNumberString);
				insRng.Font.Superscript = VBAHelper.iTRUE;

				insRng.Collapse(ref WordHelper.COLLAPSE_END);
				insRng.Font.Superscript = VBAHelper.iFALSE;

				first = false;

				insRng = mvRng.Duplicate;
				insRng.End = insRng.Start - 1;
			}

			return true;
		}

		private string translateFootnoteNumber(int footNoteNumber)
		{
			string s = "";

			int numberBase = 26;
			int startNumber = 'a' - 1;

			while ((footNoteNumber / numberBase) > 0)
			{
				int modN = footNoteNumber % numberBase;
				int n1 = startNumber + modN;
				if (modN == 0) 
				{
					n1 += numberBase;
				}

				s += ((char)n1).ToString();

				footNoteNumber -= ((footNoteNumber / numberBase) * numberBase);
			}

			if (footNoteNumber != 0) 
			{
				int n2 = startNumber + footNoteNumber;
				s += ((char)n2).ToString();
			}

			return s;
		}
    }
}
