using iTextSharp.text.pdf.parser;
using System.Text;

namespace ADXAutomation
{
    public class PdfReader
    {
        public string GetTextFromPdf(string fileName) {
            StringBuilder text = new StringBuilder();
            using (iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(fileName)) {
                for (int i = 1; i <= reader.NumberOfPages; i++) {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                }
            }

            return text.ToString();
        }
    }
}