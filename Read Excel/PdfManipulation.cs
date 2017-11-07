using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.IO;

namespace Read_Excel
{
    public class PdfManipulation
    {
        public PdfManipulation(string ppath)
        {
            this.pdfPath = ppath;
        }

        private string pdfPath;
        public String PdfPath { set => pdfPath = value; }

        public void ReportLength()
        {
            PdfReader pdfDoc = new PdfReader(pdfPath);
            String ntext = PdfTextExtractor.GetTextFromPage(pdfDoc, 128);
            Console.WriteLine(ntext);
        }

        public void AddToPdf(int[] pages, string destinationFile)
        {
            int firstPage = pages[0];
            int lastPage = pages[1];
            PdfReader pdfDoc = new PdfReader(pdfPath);
            pdfDoc.SelectPages($"{firstPage}-{lastPage}");
            PdfStamper stp = new PdfStamper(pdfDoc, new FileStream(destinationFile, FileMode.Append));
            stp.Close();
            pdfDoc.Close();
        }

        public int[] FindPages(String name, bool timesheet)
        {
            string Keyword;
            int startPage;
            int endPage;
            if (timesheet)
            {
                Keyword = "TIMESHEET";
            }
            else
            {
                Keyword = "Page";
            }

            PdfReader pdfDoc = new PdfReader(pdfPath);

            for (int i = 1; i < pdfDoc.NumberOfPages; i++)
            {
                String text = PdfTextExtractor.GetTextFromPage(pdfDoc, i);

                if (!(wordintext(text, name) && wordintext(text, Keyword)))
                    continue;
                //if the name  and keyword are found on a page...
                else
                {
                    startPage = i;
                    endPage = i;
                    for (int j = i; j < pdfDoc.NumberOfPages; j++)
                    {
                        text = PdfTextExtractor.GetTextFromPage(pdfDoc, j);
                        //for expense sheets if 2 or 3 is in the first 25 words, i.e. page 2 or page 3 continue. 
                        //convoluted, a little, but its the best we've got for right now. 
                        if (!timesheet & (wordintext(text, "2") || wordintext(text, "3")))
                            continue;
                        if (wordintext(text, name))
                            continue;
                        else
                        {
                            endPage = j - 1;
                            break;
                        }
                    }
                    int[] range = new int[2] { startPage, endPage };
                    return range;
                }
            }
            int[] brange = new int[2] { 0, 0 };
            return brange;
        }

        //return true if a word is in first 20 words of text
        private bool wordintext(string text, string word)
        {
            int LengthCheck = 26;
            //if length of text is less than 30 words use length of text rather than default 30
            if (text.Split(null).Length < LengthCheck)
                LengthCheck = text.Split(null).Length;
            //Check each of the words in text against the key word, if one matches, return true, then check if its close, if not, return false.
            for (int i = 0; i < LengthCheck; i++)
            {
                if (text.Split(null)[i] == word)
                {
                    Console.WriteLine($"Found {word} at word #{i}!");
                    return true;
                }

                if (CloseMatch(text.Split(null)[i], word))
                    return true;
            }
            return false;
        }

        private bool CloseMatch(String NewWord, String Name)
        {
            double totalLetters = Name.Length;
            double matchingLetters = 0;

            for (int i = 0; i < GetMin(NewWord, Name); i++)
            {
                if (Char.ToLower(NewWord[i]) == Char.ToLower(Name[i]))
                    matchingLetters++;
            }

            if ((matchingLetters / totalLetters) > 0.7)
                return true;
            return false;
        }

        private int GetMin(string a, string b)
        {
            return a.Length < b.Length ? a.Length : b.Length;
        }


    }
}
