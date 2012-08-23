using System;
using System.IO; //lap
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using Tspd.Utilities;

using Word = Microsoft.Office.Interop.Word;

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	// Implements a mechanism to mine a WordRange for selected direct formatting and to 
	// convert it into a tagged text format
	// Implements a mechanism to type a text with embedded formatting into a WordSelection
	// Implements a mechanism to strip tagged formatting from ICP text to yield just the raw text.
	// Currently the mechanism supports Bold, Italic, Underline, Subscript and Superscript
	/// </summary>
	public class WordFormatter
	{
		// tag headers for search purposes
        public static readonly string start = "<ft:";
        public static readonly string end = "</ft:";
		// the available tags
		public static readonly string  bold = "<ft:b>";
		public static readonly string  ebold = "</ft:b>";
		public static readonly string  italic = "<ft:i>";
		public static readonly string  eitalic = "</ft:i>";
		public static readonly string  sub = "<ft:->";
		public static readonly string  esub = "</ft:->";
		public static readonly string  sup = "<ft:+>";
		public static readonly string  esup = "</ft:+>";
		public static readonly string  underl = "<ft:u>";
		public static readonly string  eunderl = "</ft:u>";
        public const string para = "<ft:p";
        public const string epara = "</ft:p>";
        public const string blist = "<ft:l>";
        public const string eblist = "</ft:l>";
        public const string nlist = "<ft:n>";
        public const string enlist = "</ft:n>";
        public const string litem = "<ft:t>";
        public const string elitem = "</ft:t>";
        public const string hlink = "<ft:a";
        public const string ehlink = "</ft:a>";
        public const string indent = " indent=\"";
        public const string eindent = "\">";
        public const string href = " href=\"";
        public const string ehref = "\">";
        // all start and end tags have the same length 
        public static readonly int StartTagLen = bold.Length; // all starts are the same
        public static readonly int EndTagLen = ebold.Length; // all ends are the same
        public static char[] endTagProto = "</ft: >".ToCharArray();

        protected WordFormatter()
		{
		}

        private enum ParagraphType
        {
            None,
            Simple,
            BulletedList,
            NumberedList
        }

		/// <summary>
		/// Mines the Word for direct formatting and outputs a tagged string.
		/// </summary>
		/// <param name="range"></param>
		/// <returns></returns>
		public static string WordFormatToFT(Word.Range orange)
		{
			bool startB = false;
			bool startUL = false;
			bool startI = false;
			bool startSUP = false;
			bool startSUB = false;
            bool startHLink = false;
			int somethingToDo = 0;
			int boldBit = 1;
			int italicBit = 2;
			int underlineBit = 4;
			int superBit = 8;
			int subBit = 16;
            int paraBit = 32;
            int hLinkBit = 64;
            List<Word.Hyperlink> hyperlinks = null;

			// optimized
			// see if anythingto do in any category for the whole range
			Word.Range range = orange.Duplicate;
			try
			{
				range.TextRetrievalMode.IncludeHiddenText = true;
				Word.Font fullFont = range.Font;

                if (range.Bold != VBAHelper.iFALSE)
					somethingToDo |= boldBit;
				if (range.Italic != VBAHelper.iFALSE)
					somethingToDo |= italicBit; 
				if (range.Underline != Word.WdUnderline.wdUnderlineNone)
					somethingToDo |= underlineBit;
                if (fullFont.Superscript != VBAHelper.iFALSE)
					somethingToDo |= superBit;
                if (fullFont.Subscript != VBAHelper.iFALSE)
					somethingToDo |= subBit;
                if (range.Paragraphs.Count > 1)
                    somethingToDo |= paraBit;
                foreach (Word.Hyperlink hlink in range.Document.Hyperlinks)
                {
                    if (range.InRange(hlink.Range) || hlink.Range.InRange(range))
                    {
                        somethingToDo |= hLinkBit;
                        if (hyperlinks == null)
                            hyperlinks = new List<Word.Hyperlink>();
                        hyperlinks.Add(hlink);
                    }
                }

				if (somethingToDo == 0)
				{
                    // we haven't found any of our supported formats
					return range.Text.Trim();
				}
			
				System.Text.StringBuilder bf = new StringBuilder(512);
                ParagraphType paragraphType = ParagraphType.None;

				// now work on the range a character at a time, but only for 
				// the formatting we know we have.  This is because the process
				// is very timeconsuming

                foreach (Word.Paragraph para in range.Paragraphs)
                {
                    // add the start paragraph tag
                    if ((somethingToDo & paraBit) != 0)
                        paragraphType = startParagraphEmit(paragraphType, para.Range.ListFormat.ListType, para.LeftIndent, bf);

                    IEnumerator eChars = para.Range.Characters.GetEnumerator();
                    while (eChars.MoveNext())
                    {
                        Word.Range wr = eChars.Current as Word.Range;
                        //Word.Font f = wr.Font;
                        if ((somethingToDo & boldBit) != 0)
                        {
                            bool isBold = (wr.Bold == VBAHelper.iTRUE);
                            stringEmit(isBold, ref startB, WordFormatter.bold, WordFormatter.ebold, bf);
                        }

                        if ((somethingToDo & underlineBit) != 0)
                        {
                            bool isUnderline = wr.Underline != Word.WdUnderline.wdUnderlineNone;
                            stringEmit(isUnderline, ref startUL, WordFormatter.underl, WordFormatter.eunderl, bf);
                        }

                        if ((somethingToDo & italicBit) != 0)
                        {
                            bool isItalic = wr.Italic == VBAHelper.iTRUE;
                            stringEmit(isItalic, ref startI, WordFormatter.italic, WordFormatter.eitalic, bf);
                        }

                        if ((somethingToDo & subBit) != 0)
                        {
                            Word.Font f = wr.Font;
                            bool isSubscripted = f.Subscript == VBAHelper.iTRUE;
                            stringEmit(isSubscripted, ref startSUB, WordFormatter.sub, WordFormatter.esub, bf);
                        }

                        if ((somethingToDo & superBit) != 0)
                        {
                            Word.Font f = wr.Font;
                            bool isSuperscripted = f.Superscript == VBAHelper.iTRUE;
                            stringEmit(isSuperscripted, ref startSUP, WordFormatter.sup, WordFormatter.esup, bf);
                        }
                        if ((somethingToDo & hLinkBit) != 0)
                        {
                            string target = null;
                            foreach (Word.Hyperlink hyperlink in hyperlinks)
                            {
                                if ((wr.Start >= hyperlink.Range.Start) && (wr.Start <= hyperlink.Range.End))
                                {
                                    target = hyperlink.Target;
                                    break;
                                }

                                if (!String.IsNullOrEmpty(target))
                                    break;
                            }

                            stringEmit(!String.IsNullOrEmpty(target), ref startHLink, WordFormatter.hlink, WordFormatter.ehlink, bf, target);
                        }
                        bf.Append(wr.Text);
                    }

                }

                // add the final end paragraph tags
                if ((somethingToDo & paraBit) != 0)
                    endParagraphEmit(true, paragraphType, Word.WdListType.wdListNoNumbering, bf);

                return formatCleaner(bf.ToString().Trim());  // char enumerator always returns a space that's not there
			} 
			catch (Exception)
			{
				return orange.Text;
			}
		}

		private static void stringEmit(bool prop, ref bool startFlag, string startSym, string endSym, StringBuilder bf, string hlinkRef = null)
		{
			if (!prop && startFlag)
			{
				// condition ending
				startFlag = false;
				bf.Append(endSym);
			}
			if (prop && !startFlag)
			{
				// condition starting
				startFlag = true;
                bf.Append(startSym);
                if (!String.IsNullOrEmpty(hlinkRef))
                {
                    bf.Append(href);
                    bf.Append(hlinkRef);
                    bf.Append(ehref);
                }
			}
			return;
		}

        private static ParagraphType startParagraphEmit(ParagraphType currentType, Word.WdListType listType, float indent, StringBuilder bf)
        {
            ParagraphType ret = ParagraphType.None;
            endParagraphEmit(false, currentType, listType, bf);

            switch (listType)
            {
                case Word.WdListType.wdListSimpleNumbering:
                    if (currentType != ParagraphType.NumberedList)
                        bf.Append(nlist);
                    bf.Append(litem);
                    ret = ParagraphType.NumberedList;
                    break;

                case Word.WdListType.wdListBullet:
                    if (currentType != ParagraphType.BulletedList)
                        bf.Append(blist);
                    bf.Append(litem);
                    ret = ParagraphType.BulletedList;
                    break;

                default:
                    int numIndents = (int)Math.Round(indent / 18.0);
                    if (numIndents == 0)
                    {
                        bf.Append(para);
                        bf.Append(">");
                    }
                    else
                    {
                        bf.Append(para);
                        bf.Append(indent);
                        bf.Append(numIndents.ToString());
                        bf.Append(eindent);
                    }
                    ret = ParagraphType.Simple;
                    break;
            }

            return ret;
        }

        private static void endParagraphEmit(bool forceWrite, ParagraphType currentType, Word.WdListType listType, StringBuilder bf)
        {
            switch (currentType)
            {
                case ParagraphType.None:
                    break;

                case ParagraphType.Simple:
                    bf.Append(epara);
                    break;

                case ParagraphType.BulletedList:
                    bf.Append(elitem);
                    if (forceWrite || (listType != Word.WdListType.wdListBullet))
                        bf.Append(eblist);
                    break;

                case ParagraphType.NumberedList:
                    bf.Append(elitem);
                    if (forceWrite || (listType != Word.WdListType.wdListSimpleNumbering))
                        bf.Append(enlist);
                    break;
            }
        }

        private class FormatTag
        {
            public enum TagType
            {
                NONE,             // no formatting
                BOLD,
                ITALIC,
                SUPERSCRIPT,
                SUBSCRIPT,
                UNDERLINE,
                PARAGRAPH,
                HYPERLINK,
                NUMBERED_LIST,
                BULLETED_LIST,
                LIST_ITEM
            }

            private List<FormatTag> _subTags;
            protected static int S_paragraphCount = 0;

            public FormatTag()
            {
                _subTags = new List<FormatTag>();
            }

            protected Word.Range FormatRange;

            public TagType Type;
            public FormatTag Parent;

            private int _startIndex;
            public int StartIndex
            {
                get { return (FormatRange == null) ? _startIndex : FormatRange.Start; }
                set
                {
                    if (FormatRange == null)
                        _startIndex = value;
                    else
                        FormatRange.Start = value;
                }
            }

            private int _endIndex;
            public int EndIndex
            {
                get { return (FormatRange == null) ? _endIndex : FormatRange.End; }
                set
                {
                    if (FormatRange == null)
                        _endIndex = value;
                    else
                        FormatRange.End = value;
                }
            }

            private string _text;
            public string Text
            {
                get
                {
                    if (!String.IsNullOrEmpty(_text))
                        return _text;

                    if (Head == null)
                        return String.Empty;

                    if (FormatRange == null)
                    {
                        int headIndex = Head.StartIndex;
                        int startIndex = StartIndex - headIndex;
                        int endIndex = EndIndex - headIndex;
                        return Head.Text.Substring(startIndex, endIndex - startIndex);
                    }
                    else
                        return FormatRange.Text;
                }
                set { _text = value; }
            }
            public List<FormatTag> SubTags { get { return _subTags; } }

            private int _markersAdded;
            public int MarkersAdded
            {
                get
                {
                    int offset = _markersAdded;
                    FormatTag current = Parent;
                    while (current != null)
                    {
                        offset += current._markersAdded;
                        current = current.Parent;
                    }
                    return offset;
                }
            }

            public virtual void StartAction(Word.Range range,
                Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null) 
            {
                FormatRange = range.Duplicate;
                FormatRange.Start = _startIndex;
                FormatRange.End = _endIndex;
            }

            public virtual void EndAction(Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            { }

            public virtual void Dump(StreamWriter sw, string text, int textStartIndex, int numIndents = 0)
            {
                StringBuilder sb = new StringBuilder();
                for (int ii = 0; ii < (numIndents * 4); ii++)
                    sb.Append(" ");
                string indent = sb.ToString();

                sw.WriteLine(indent + "FormatTag");
                sw.WriteLine(indent + "    Type       = " + this.GetType().Name);
                sw.WriteLine(indent + "    StartIndex = " + FormatRange.Start);
                sw.WriteLine(indent + "    EndIndex   = " + FormatRange.End);
                if (!String.IsNullOrEmpty(Text))
                    sw.WriteLine(indent + "    Text       = " + Text);

                foreach (FormatTag subTag in _subTags)
                    subTag.Dump(sw, text, textStartIndex, numIndents + 1);
            }

            protected FormatTag Head
            {
                get
                {
                    FormatTag head = null;
                    FormatTag current = this;
                    while (current != null)
                    {
                        if (current.Parent == null)
                        {
                            head = current;
                            break;
                        }
                        current = current.Parent;
                    }
                    return head;
                }
            }

            protected void InsertParagraphBefore(Word.Range range)
            {
                bool succeeded = false;
                while (!succeeded)
                {
                    try
                    {
                        range.InsertParagraphBefore();
                        //range.Start++; // move the range forward
                        succeeded = true;
                    }
                    catch (Exception)
                    {
                    }
                }

                if (Head != null)
                    Head._markersAdded++;
            }

            protected void InsertHyperlink(Word.Range range, string address, string text)
            {
                range.Document.Hyperlinks.Add(range, address, System.Type.Missing, System.Type.Missing, range.Text);
                if (Head != null)
                    Head._markersAdded++;
            }

            protected Word.Range TrimRange()
            {
                Word.Range range = FormatRange.Duplicate;
                if (range.Text.StartsWith("\r"))
                    range.Start++;
                if (range.Text.EndsWith("\r"))
                    range.End--;

                return range;
            }

            protected void DoListIndent()
            {
                Word.Range range = TrimRange();

                if (range.ListFormat == null)
                    return;

                range.ListFormat.ListIndent();
            }
        }

        private class NoneFormatTag : FormatTag
        {
            public NoneFormatTag(Word.Range range)
                : base()
            {
                Type = TagType.NONE;
                S_paragraphCount = 0;
            }

            public override void StartAction(Word.Range range, Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.StartAction(range, normalStyle, bulletedStyle, numberedStyle);
            }

            public override void EndAction(Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.EndAction(normalStyle, bulletedStyle, numberedStyle);
            }
        }

        private class BoldFormatTag : FormatTag
        {
            public BoldFormatTag(Word.Range range) : base() { Type = TagType.BOLD; }

            public override void StartAction(Word.Range range, Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.StartAction(range, normalStyle, bulletedStyle, numberedStyle);
                FormatRange.Font.Bold = VBAHelper.iTRUE;
            }

            public override void EndAction(Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.EndAction(normalStyle, bulletedStyle, numberedStyle);
            }
        }

        private class ItalicFormatTag : FormatTag
        {
            public ItalicFormatTag(Word.Range range) : base() { Type = TagType.ITALIC; }

            public override void StartAction(Word.Range range, Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.StartAction(range, normalStyle, bulletedStyle, numberedStyle);
                FormatRange.Font.Italic = VBAHelper.iTRUE;
            }

            public override void EndAction(Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.EndAction(normalStyle, bulletedStyle, numberedStyle);
                FormatRange.Font.Italic = VBAHelper.iTRUE;
            }
        }

        private class SuperscriptFormatTag : FormatTag
        {
            public SuperscriptFormatTag(Word.Range range) : base() { Type = TagType.SUPERSCRIPT; }

            public override void StartAction(Word.Range range, Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.StartAction(range, normalStyle, bulletedStyle, numberedStyle);
                FormatRange.Font.Superscript = VBAHelper.iTRUE;
            }

            public override void EndAction(Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.EndAction(normalStyle, bulletedStyle, numberedStyle);
            }
        }

        private class SubscriptFormatTag : FormatTag
        {
            public SubscriptFormatTag(Word.Range range) : base() { Type = TagType.SUBSCRIPT; }

            public override void StartAction(Word.Range range, Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.StartAction(range, normalStyle, bulletedStyle, numberedStyle);
                FormatRange.Font.Subscript = VBAHelper.iTRUE;
            }

            public override void EndAction(Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.EndAction(normalStyle, bulletedStyle, numberedStyle);
            }
        }

        private class UnderlineFormatTag : FormatTag
        {
            public UnderlineFormatTag(Word.Range range) : base() { Type = TagType.UNDERLINE; }

            public override void StartAction(Word.Range range, Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.StartAction(range, normalStyle, bulletedStyle, numberedStyle);
                FormatRange.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
            }

            public override void EndAction(Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.EndAction(normalStyle, bulletedStyle, numberedStyle);
            }
        }

        private class ParagraphFormatTag : FormatTag
        {
            private static bool S_haveParagraph = false;

            public ParagraphFormatTag(Word.Range range)
                : base()
            { 
                Type = TagType.PARAGRAPH;
                S_haveParagraph = false;  // We need to generate ALL Paragraph tags before processing...
                Number = S_paragraphCount++;
            }

            public int Indent;
            private int Number;

            public override void StartAction(Word.Range range, Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.StartAction(range, normalStyle, bulletedStyle, numberedStyle);
            }

            // NOTE:  The contents of this method will add a character (marker) to the document, so all children
            //        MUST be done first!
            public override void EndAction(Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.EndAction(normalStyle, bulletedStyle, numberedStyle);

                // first, add a carriage-return
                if (S_haveParagraph || ((S_paragraphCount > 1) && (Number > 0)))
                    InsertParagraphBefore(FormatRange);
                else
                    S_haveParagraph = true;

                // stylize the trimmed range
                Word.Range range = this.TrimRange();

                if (Number == 0)
                {
                    // apply the style
                    if (normalStyle != null)
                        range.set_Style(normalStyle);
                }
                else
                {
                    // remove numbers (if any) before all but the first paragraph
                    range.ListFormat.RemoveNumbers(Word.WdNumberType.wdNumberParagraph);
                    range.ParagraphFormat.LeftIndent = (normalStyle == null) ? 0 : normalStyle.ParagraphFormat.LeftIndent;
                }

                for (int ii = 0; ii < Indent; ii++)
                {
                    range.Paragraphs.Indent();
                }
            }

            public override void Dump(StreamWriter sw, string text, int textStartIndex, int numIndents = 0)
            {
                base.Dump(sw, text, textStartIndex, numIndents);

                StringBuilder sb = new StringBuilder();
                for (int ii = 0; ii < (numIndents * 4); ii++)
                    sb.Append(" ");
                string indent = sb.ToString();

                sw.WriteLine(indent + "    Number     = " + Number);
                if (Indent > 0)
                    sw.WriteLine(indent + "    Indent     = " + Indent);
            }
        }

        private class HyperlinkFormatTag : FormatTag
        {
            public HyperlinkFormatTag(Word.Range range) : base() { Type = TagType.HYPERLINK; }

            public string Address;

            public override void StartAction(Word.Range range, Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.StartAction(range, normalStyle, bulletedStyle, numberedStyle);
                InsertHyperlink(FormatRange, Address, Text);
            }

            public override void EndAction(Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.EndAction(normalStyle, bulletedStyle, numberedStyle);
            }

            public override void Dump(StreamWriter sw, string text, int textStartIndex, int numIndents = 0)
            {
                base.Dump(sw, text, textStartIndex, numIndents);

                if (!String.IsNullOrEmpty(Address))
                {
                    StringBuilder sb = new StringBuilder();
                    for (int ii = 0; ii < (numIndents * 4); ii++)
                        sb.Append(" ");
                    string indent = sb.ToString();
                    sw.WriteLine(indent + "    Address    = " + Address);
                }
            }
        }

        private class BulletedListFormatTag : FormatTag
        {
            public BulletedListFormatTag(Word.Range range) : base() { Type = TagType.BULLETED_LIST; }

            public override void StartAction(Word.Range range, Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.StartAction(range, normalStyle, bulletedStyle, numberedStyle);
            
                try
                {
//                    range.ListFormat.RemoveNumbers(Word.WdNumberType.wdNumberAllNumbers);
                }
                catch (Exception) { }
            }

            public override void EndAction(Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.EndAction(normalStyle, bulletedStyle, numberedStyle);

                DoListIndent();

                Word.Range trimRange = TrimRange();
                if (bulletedStyle == null)
                    trimRange.ListFormat.ApplyBulletDefault();
                else
                    trimRange.set_Style(bulletedStyle);
            }
        }

        private class NumberedListFormatTag : FormatTag
        {
            public NumberedListFormatTag(Word.Range range) : base() { Type = TagType.NUMBERED_LIST; }

            public override void StartAction(Word.Range range, Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.StartAction(range, normalStyle, bulletedStyle, numberedStyle);

                try
                {
                //    range.ListFormat.RemoveNumbers(Word.WdNumberType.wdNumberAllNumbers);
                }
                catch (Exception) { }
            }

            public override void EndAction(Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.EndAction(normalStyle, bulletedStyle, numberedStyle);

                DoListIndent();

                Word.Range trimRange = TrimRange();
                if (numberedStyle == null)
                    trimRange.ListFormat.ApplyNumberDefault();
                else
                    trimRange.set_Style(numberedStyle);

                PurdueUtil.resartListNumber(trimRange);
            }
        }

        private class ListItemFormatTag : FormatTag
        {
            public ListItemFormatTag(Word.Range range)
                : base()
            { 
                Type = TagType.LIST_ITEM;
                S_paragraphCount++; // we DO add a paragraph...
            }

            public override void StartAction(Word.Range range, Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.StartAction(range, normalStyle, bulletedStyle, numberedStyle);
            }

            public override void EndAction(Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
            {
                base.EndAction(normalStyle, bulletedStyle, numberedStyle);

                // each list item is a separate paragraph
                InsertParagraphBefore(FormatRange);

                //range.Start++; // we added a paragraph before...
            }
        }

        // FormatTag newTag = new FormatTag { Parent = currentTag, StartIndex = outText.Length + rangeStart, Type = findTagType(inputString[4]) };

        private static FormatTag FormatTagFactory(Word.Range range, char tag, FormatTag parent, int startIndex)
        {
            FormatTag newTag = null;
            switch (tag)
            {
                case '+':
                    newTag = new SuperscriptFormatTag(range)  { Parent = parent };
                    break;
                case '-':
                    newTag = new SubscriptFormatTag(range) { Parent = parent };
                    break;
                case 'a':
                    newTag = new HyperlinkFormatTag(range) { Parent = parent };
                    break;
                case 'b':
                    newTag = new BoldFormatTag(range) { Parent = parent };
                    break;
                case 'i':
                    newTag = new ItalicFormatTag(range) { Parent = parent };
                    break;
                case 'l':
                    newTag = new BulletedListFormatTag(range) { Parent = parent };
                    break;
                case 'n':
                    newTag = new NumberedListFormatTag(range) { Parent = parent };
                    break;
                case 'p':
                    newTag = new ParagraphFormatTag(range) { Parent = parent };
                    break;
                case 't':
                    newTag = new ListItemFormatTag(range) { Parent = parent };
                    break;
                case 'u':
                    newTag = new UnderlineFormatTag(range) { Parent = parent };
                    break;
                default:
                    newTag = new NoneFormatTag(range) { Parent = parent };
                    break;
            }

            newTag.StartIndex = startIndex;
            if (parent != null)
                parent.SubTags.Add(newTag);

            return newTag;
        }

        /// <summary>
		/// Given an input Range and a potentially tagged string, initializes the Range with the
		/// text, implementing the embedded formatting suggested by the tags.  Returns a Range containing this
		/// selection
		/// </summary>
		/// <param name="workingRange"></param>
		/// <param name="input"></param>
		/// <returns></returns>
		public static void FTToWordFormat2(ref Word.Range workingRange, string input,
            Word.Style normalStyle = null, Word.Style bulletedStyle = null, Word.Style numberedStyle = null)
		{
            workingRange.Collapse(ref WordHelper.COLLAPSE_END);

			string outText = "";
            int numParagraphs = 0;
			int rangeStart = workingRange.Start;

            string inputString = input.Trim();
            NoneFormatTag headTag = new NoneFormatTag(workingRange) { StartIndex = workingRange.Start, EndIndex = workingRange.End };
            FormatTag currentTag = headTag;

            // NOTE:  It's ugly, but the debug code is in-line...
            bool writeDebugFile = false;  // do not set to "true" in production code!
            StreamWriter sw = (writeDebugFile) ? new StreamWriter("c:\\formatting2.txt", true) : null;
            if (sw != null)
            {
                sw.WriteLine("FMT----------- start ---------------");
                //Log.trace("FMT----------- start ---------------");
            }
            bool dumpedTags = false;
 
            // let's get the normal style
            if (normalStyle == null)
            {
                try
                {
                    normalStyle = workingRange.get_Style() as Word.Style;
                }
                catch (Exception) { } // definitely not the end of the world...
            }

            while (true)
			{
                int tagStartIndex = inputString.IndexOf(start);
                int tagEndIndex   = inputString.IndexOf(end);

                if (tagStartIndex == -1 && tagEndIndex == -1)
				{
                    outText += inputString; // fill in the text after the last tag
                    workingRange.Text = outText;

                    // trim the tags
                    headTag.Text = outText;
                    headTag.EndIndex = outText.Length + rangeStart;
                    trimTags(headTag);

                    if (sw != null)
                    {
                        if (!dumpedTags)
                        {
                            sw.WriteLine("Input Text = " + input);
                            headTag.Dump(sw, input, rangeStart);
                            dumpedTags = true;
                        }

                        sw.WriteLine("Last Text: " + inputString);
                        //Log.trace("Last Text: " + bf);
                    }

                    if (headTag.SubTags.Count > 0)
                    {
                        applyTags(headTag, workingRange, numParagraphs, normalStyle, bulletedStyle, numberedStyle, sw, outText, rangeStart);
                    }

                    return;
				}
				else
				{
					// starts aways in from where we are now
					// because selection cannot bridge label boundary, there should be no
					// unclosed formatting instructions
					// <ft:i>here we go  no
					// there we went</ft:i>  no
					// <ft:i>hello and good bye</ft:i> yes
					// <ft:i>hello <ft:b>and</ft:b> good bye</ft:i> yes
					// This is a very detailed <ft:b>Now</ft:b> I'm going to Write C<ft:+>max </ft:+> desc<ft:u>r</ft:u>iption
					int tagNextInstruction = -1;

					if (tagStartIndex == -1)
                        tagNextInstruction = tagEndIndex;
					else if (tagEndIndex == -1)
                        tagNextInstruction = tagStartIndex;
					else
                        tagNextInstruction = Math.Min(tagStartIndex, tagEndIndex);

                    if (tagStartIndex != 0 && tagEndIndex != 0)
                    {
                        // xxxxxxx<ft:
                        // output the bit with no formatting
                        outText += (inputString.Substring(0, tagNextInstruction));
                        // position ourselves at the next instruction
                        if (sw != null)
                        {
                            sw.WriteLine("insert Text: " + inputString.Substring(0, tagNextInstruction));
                            //Log.trace("insert Text: " + bf.Substring(0, nextInstruction));
                        }
                        inputString = inputString.Substring(tagNextInstruction);

                        if (sw != null)
                        {
                            sw.WriteLine("after Text: " + inputString);
                            //Log.trace("after Text: " + bf);
                        }
                    }

                    // look for a start
                    if (tagNextInstruction == tagStartIndex)
                    {
                        FormatTag newTag = FormatTagFactory(workingRange, inputString[4], currentTag, outText.Length + rangeStart);
                        switch (newTag.Type)
                        {
                            case FormatTag.TagType.HYPERLINK:
                                if (inputString[5] == '>')
                                    inputString = inputString.Substring(StartTagLen);
                                else
                                {
                                    // do we have the address?
                                    string tmlHref = inputString.Substring(StartTagLen);
                                    int hrefIndex = tmlHref.IndexOf("href=\"");
                                    if (hrefIndex >= 0)
                                    {
                                        hrefIndex += 6; // length of "href=\""
                                        char[] splits = { '"', '>' };
                                        string[] strings = tmlHref.Substring(hrefIndex).Split(splits);
                                        (newTag as HyperlinkFormatTag).Address = strings[0];
                                    }

                                    // reset the string
                                    inputString = inputString.Substring(StartTagLen + tmlHref.IndexOf('>') + 1);
                                }
                                break;
                            case FormatTag.TagType.PARAGRAPH:
                                ParagraphFormatTag pTag = newTag as ParagraphFormatTag;
                                if (inputString[5] == '>')
                                    inputString = inputString.Substring(StartTagLen);
                                else
                                {
                                    // do we have indent?
                                    string tmpIndent = inputString.Substring(StartTagLen);
                                    int indent_index = tmpIndent.IndexOf("indent=\"");
                                    if (indent_index >= 0)
                                    {
                                        indent_index += 8; // length of "indent=\""
                                        char[] splits = { '"', '>' };
                                        string[] strings = tmpIndent.Substring(indent_index).Split(splits);
                                        int tryIndent = 0;
                                        Int32.TryParse(strings[0], out tryIndent);
                                        pTag.Indent = tryIndent;
                                    }
                                    else
                                        pTag.Indent = 0;

                                    // reset the string
                                    inputString = inputString.Substring(StartTagLen + tmpIndent.IndexOf('>') + 1);
                                }
                                break;
                        }

                        currentTag = newTag;

                        // reset the string for "simple" tags
                        if ((newTag.Type != FormatTag.TagType.HYPERLINK) && (newTag.Type != FormatTag.TagType.PARAGRAPH))
                            inputString = inputString.Substring(StartTagLen);
                    }

                    else if (tagNextInstruction == tagEndIndex)
                    {
                        // find the correct tag
                        FormatTag.TagType tagType = findTagType(inputString[5]);
                        FormatTag cTag = currentTag;
                        while (cTag != null)
                        {
                            if (cTag.Type == tagType)
                            {
                                cTag.EndIndex = outText.Length + rangeStart;
                                currentTag = cTag.Parent;

                                checkSubTagsEnd(currentTag);
                                break;
                            }
                            cTag = cTag.Parent;
                        }

                        // we've got a tag-end without a tag beginning
                        if (cTag == null)
                        {
                            // does the parent have other tags?
                            if (currentTag.SubTags.Count > 0)
                            {
                                FormatTag newTag = FormatTagFactory(workingRange, inputString[5], currentTag,
                                    currentTag.SubTags[currentTag.SubTags.Count - 1].EndIndex + 1);
                                newTag.EndIndex = outText.Length + rangeStart;
                            }
                        }

                        // reset the string
                        inputString = inputString.Substring(EndTagLen);
                    }                   
				}
            }
		}

        private static FormatTag.TagType findTagType(char tag)
        {
            switch (tag)
            {
                case '+':
                    return FormatTag.TagType.SUPERSCRIPT;
                case '-':
                    return FormatTag.TagType.SUBSCRIPT;
                case 'a':
                    return FormatTag.TagType.HYPERLINK;
                case 'b':
                    return FormatTag.TagType.BOLD;
                case 'i':
                    return FormatTag.TagType.ITALIC;
                case 'l':
                    return FormatTag.TagType.BULLETED_LIST;
                case 'n':
                    return FormatTag.TagType.NUMBERED_LIST;
                case 'p':
                    return FormatTag.TagType.PARAGRAPH;
                case 't':
                    return FormatTag.TagType.LIST_ITEM;
                case 'u':
                    return FormatTag.TagType.UNDERLINE;
                default:
                    return FormatTag.TagType.NONE;
            }
        }

        private static void checkSubTagsEnd(FormatTag tag)
        {
            foreach (FormatTag subTag in tag.SubTags)
            {
                if (subTag.EndIndex == 0)
                    subTag.EndIndex = tag.EndIndex;
                checkSubTagsEnd(subTag);
            }
        }

        private static void applyTags(FormatTag tag, Word.Range range, int paragraphCount, Word.Style normalStyle,
            Word.Style bulletedStyle, Word.Style numberedStyle, StreamWriter debugStream, string wholeText, int wholeTextStartIndex)
        {
            if (debugStream != null)
            {
                debugStream.WriteLine(">>> BEFORE");
                tag.Dump(debugStream, wholeText, wholeTextStartIndex, 0);
            }

            // perform start action
            tag.StartAction(range, normalStyle, bulletedStyle, numberedStyle);

            foreach (FormatTag subTag in tag.SubTags.OrderByDescending(tg => tg.StartIndex))
            {
                applyTags(subTag, range, paragraphCount, normalStyle, bulletedStyle, numberedStyle, debugStream, wholeText, wholeTextStartIndex);
            }

            // perform end action
            tag.EndAction(normalStyle, bulletedStyle, numberedStyle);

            if (debugStream != null)
            {
                debugStream.WriteLine(">>> AFTER");
                tag.Dump(debugStream, wholeText, wholeTextStartIndex, 0);
            }
        }

        private static void trimTags(FormatTag tag)
        {
            bool foundRealParagraph = false;
            for (int ii = tag.SubTags.Count - 1; ii >= 0; ii--)
            {
                FormatTag subTag = tag.SubTags[ii];

                // remove empty trailing paragraphs
                if ((tag is NoneFormatTag) 
                    && !foundRealParagraph 
                    && (subTag is ParagraphFormatTag) 
                    && String.IsNullOrWhiteSpace(subTag.Text))
                {
                    tag.SubTags.Remove(subTag);
                    continue;
                }
                else
                    foundRealParagraph = true;

                // it's empty - remove it!
                if (subTag.StartIndex >= subTag.EndIndex)
                {
                    tag.SubTags.Remove(subTag);
                    continue;
                }

                // it's out of range - remove it!
                if (subTag.StartIndex > tag.EndIndex)
                {
                    tag.SubTags.Remove(subTag);
                    continue;
                }

                // adjust the end index
                if ((subTag.EndIndex > tag.EndIndex) || (subTag.EndIndex == 0))
                    subTag.EndIndex = tag.EndIndex;

                // check the children
                trimTags(subTag);
            }
        }

        /// <summary>
        /// Removes embedded tags in the form <ft:.> </ft:.> from  an icpValue
        /// need to deal with 
        /// no formatting
        /// balanced pairs
        /// overlapping pairs
        /// missing either end
        /// </summary>
        /// <param name="icpValue"></param>
        /// <returns></returns>
        public static string stripFormatInstruction(string icpValue)
		{
			if (icpValue == null || icpValue.Length == 0)
				return icpValue;

			string compoundString = "";
			string safety = icpValue.Trim();
			try
			{

				while(true)
				{
					int nextStart = icpValue.IndexOf(WordFormatter.start);
					int nextEnd = icpValue.IndexOf(WordFormatter.end);
				
					if (nextStart == -1 && nextEnd == -1)
					{
						return compoundString + icpValue;
					}

						// note that both cannot be zero
					else if (nextStart > 0 && nextEnd > 0)
					{
						int next = Math.Min(nextStart, nextEnd);
						if (next > 0) // may be two tags up against each other
						{
							compoundString += icpValue.Substring(0, next);
						}

						icpValue = icpValue.Substring(next + 
							((next == nextStart) ? StartTagLen: EndTagLen));

						// leave pointing just after the next tag
					}

					else if (nextStart >= 0)
					{
						if (nextStart > 0)
						{
							compoundString += icpValue.Substring(0, nextStart);
						}
						icpValue = icpValue.Substring(nextStart + StartTagLen);
						// leave pointing just after the next tag
					}

					else if (nextEnd >= 0)
					{
						if (nextEnd > 0)
						{
							compoundString += icpValue.Substring(0, nextEnd);
						}
						icpValue = icpValue.Substring(nextEnd + EndTagLen);
						// leave pointing just after the next tag
					}
				}
			}
			catch(Exception)
			{
				return safety;
			}
		}

		/// <summary>
		/// Removes and patches format tags
		/// removes duplicate start, end without start, duplicate end
		/// adds missing end
		/// It does not remove fragmentary tags
		/// </summary>
		/// <param name="icpValue"></param>
		/// <returns></returns>
		public static string formatCleaner(string icpValue)
		{
			if (icpValue == null)
				return null;

			if (icpValue.Length == 0)
				return icpValue;

			string safety = icpValue;
			string compoundString = "";
			Hashtable openTags = new Hashtable();
			
			try
			{
				while(true)
				{
					int nextStart = icpValue.IndexOf(WordFormatter.start);
					int nextEnd = icpValue.IndexOf(WordFormatter.end);
				
					if (nextStart == -1 && nextEnd == -1)
					{
						compoundString += icpValue;
						// attach ending tags to any that arent closed
						foreach (char cmd in openTags.Values)
						{
							// copy the format instruction char {i,u,b,+,-) into the prototype end tag
							endTagProto[EndTagLen-2] = cmd;
							// and use it
							compoundString += new string(endTagProto);
						}
						return compoundString;
					}

						// note that both cannot be zero
					else if (nextStart > 0 && nextEnd > 0)
					{
						int next = Math.Min(nextStart, nextEnd);
						if (next > 0) // may be two tags up against each other
						{
							compoundString += icpValue.Substring(0, next);
						}

						if (next == nextStart)
						{
							string theTag = icpValue.Substring(next, StartTagLen);
							char cmd = theTag[4];
							if(openTags[cmd] == null)
							{
								// only add the tag if there isn't one open of this type
								openTags[cmd] = cmd;
								compoundString += theTag;
							}
						}
						else if(next == nextEnd)
						{
							string theTag = icpValue.Substring(next, EndTagLen);
							char cmd = theTag[5];
							if(openTags[cmd] == null)
							{
								// end without a start - ignore it
							}
							else
							{
								// it was open - now close and reset tag indicator
								openTags.Remove(cmd);
								compoundString += theTag;
							}
						}

						icpValue = icpValue.Substring(next + 
							((next == nextStart) ? StartTagLen: EndTagLen));

						// leave pointing just after the next tag
					}

					else if (nextStart >= 0)
					{
						if (nextStart > 0)
						{
							compoundString += icpValue.Substring(0, nextStart);
						}

						string theTag = icpValue.Substring(nextStart, StartTagLen);
						char cmd = theTag[4];
						if(openTags[cmd] == null)
						{
							// only add the tag if there isn't one open of this type
							openTags[cmd] = cmd;
							compoundString += theTag;
						}
						else
						{
							// ignore that format type is already open
						}

						icpValue = icpValue.Substring(nextStart + StartTagLen);
						// leave pointing just after the next tag
					}

					else if (nextEnd >= 0)
					{
						if (nextEnd > 0)
						{
							compoundString += icpValue.Substring(0, nextEnd);
						}

						string theTag = icpValue.Substring(nextEnd, EndTagLen);
						char cmd = theTag[5];
						if(openTags[cmd] == null)
						{
							// end without a start - ignore it
						}
						else
						{
							// it was open - now close and reset tag indicator
							openTags.Remove(cmd);
							compoundString += theTag;
						}

						icpValue = icpValue.Substring(nextEnd + EndTagLen);
						// leave pointing just after the next tag
					}
				}
			}
			catch(Exception)
			{
				return safety;
			}
		}

		public static bool hasFormatting(string s)
		{
			if (s == null || s.Length == 0)
				return false;
			return (s.IndexOf(start) != -1);
		}

		public static string newlineConvertForDisplay(string source)
		{
			if (source != null && source.Length > 0)
			{
				// Step 1, preserve \r\n
				string str = source.Replace("\r\n", "\n");

				// Step 2, catch any \n \r\n
				str = str.Replace("\n", "\r\n");

				return str;
			}
			return source;
		}

		public static bool Equals(string s1, string s2) 
		{
			if (s1 == null && s2 == null) 
			{
				return true;
			}

			if (s1 == null || s2 == null) 
			{
				return false;
			}

			s1 = stripFormatInstruction(s1.Trim());
			s2 = stripFormatInstruction(s2.Trim());

			return s1.Equals(s2);
		}

		/// <summary>
		/// Test the stripper and cleaner mechanisms for the possible and impossible tags one 
		/// may find in a given element.
		/// </summary>
		public static void testStripper()
		{
			string  s1 = "dr. no's nose knows no snows";
			bool  s2 = s1.Equals(stripFormatInstruction("<ft:i>dr. no's nose knows no snows</ft:i>"));
			s2 = s1.Equals(stripFormatInstruction("<ft:i>dr. no's</ft:i> nose knows no snows"));
			s2 = s1.Equals(stripFormatInstruction("dr. no's nose knows </ft:i>no snows"));
			s2 = s1.Equals(stripFormatInstruction("dr. no's <ft:i>nose knows no snows"));
			s2 = s1.Equals(stripFormatInstruction("<ft:i>dr. no's nose knows no snows"));
			s2 = s1.Equals(stripFormatInstruction("dr. no's nose knows no snows<ft:i>"));
			s2 = s1.Equals(stripFormatInstruction("<ft:i>dr. no's nose knows no snows"));
			s2 = s1.Equals(stripFormatInstruction("dr. no's nose knows no snows</ft:i>"));
			s2 = s1.Equals(stripFormatInstruction("</ft:i>dr. no's nose knows no snows<ft:i>"));
			s2 = s1.Equals(stripFormatInstruction("dr. no's nose <ft:i></ft:i>knows no snows"));
			s2 = s1.Equals(stripFormatInstruction("(<ft:b>dr. <ft:i>no's nose(</ft:i> knows (</ft:b>no snows"));

			s1 = "<ft:u>dr. <ft:i>no's</ft:i> nose knows no snows</ft:u>";
			// well formed
			s2 = s1.Equals(formatCleaner( "<ft:u>dr. <ft:i>no's</ft:i> nose knows no snows</ft:u>"));
			// duplicate start
			s2 = s1.Equals(formatCleaner( "<ft:u>dr. <ft:u><ft:i>no's</ft:i> nose knows no snows</ft:u>"));
			// end without start
			s2 = s1.Equals(formatCleaner( "<ft:u>dr. <ft:i>no's</ft:i> nose </ft:b>knows no snows</ft:u>"));
			// start without end
			s2 = s1.Equals(formatCleaner( "<ft:u>dr. <ft:i>no's</ft:i> nose knows no snows"));
		}


	}	
}
