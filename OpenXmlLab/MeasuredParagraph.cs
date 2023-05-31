using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlLab
{
    public class MeasuredParagraph
    {
        private Paragraph _p;
        public MeasuredParagraph(Paragraph p)
        {
            this._p = p;
        }

        public string ParaId
        {
            get
            {
                if (this._p.ParagraphId is null) {
                    return "-";
                }

                return this._p.ParagraphId!;
            }
        }

        public string SampleText
        {
            get
            {
                var text = this._p.InnerText;
                var limit = Math.Min(text.Length, 31);
                var suffix = limit < 31 ? "" : "...";
                return $"{text.Substring(0, limit)}{suffix}";
            }
        }

        public int InnerTextLength
        {
            get
            {
                return this._p.InnerText.Length;
            }
        }
    }
}