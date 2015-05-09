using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Novacode;
using System.IO;


namespace worddocumentreader
{
    class Program
    {
        static void Main(string[] args)
        {

           // System.Diagnostics.Debugger.Launch();

            Dictionary<string, float> TextVsValue = new Dictionary<string, float>();

            using (DocX doc = DocX.Load("d:\\Sample.docx"))
            {              
                
                for (int i = 0; i < doc.Paragraphs.Count; i++)
                {
                                        
                    foreach (var item in doc.Paragraphs[i].Text.Split(new string[] { "\n" }
                              , StringSplitOptions.RemoveEmptyEntries))
                    {
                        foreach (string searchText in args)
                        {
                            if (item.ToLower().Contains(searchText.ToLower()))
                            {
                                // Highlight text using an indirect method of replace since docX currently
                                //does not have support to highlight just a single word.
                                if (doc.Paragraphs[i] is Paragraph)
                                {
                                    Paragraph sen = doc.Paragraphs[i] as Paragraph;
                                    Formatting form = new Formatting();
                                    form.Highlight = Highlight.yellow;
                                    form.Bold = true;
                                    sen.ReplaceText(searchText, searchText, false, System.Text.RegularExpressions.RegexOptions.IgnoreCase,
                                        form, null, MatchFormattingOptions.ExactMatch);
                                }
                                StringAndValue SearchString = new StringAndValue(item, searchText);
                                float val = SearchString.ExtractStringAndValue();
                                if (!TextVsValue.ContainsKey(searchText))
                                {
                                    TextVsValue.Add(searchText, val);
                                }
                                else if(TextVsValue[searchText] < val)
                                {
                                    TextVsValue[searchText] = val;
                                }
                                doc.Save();
                            }
                        }
                    }
                }
                Console.ReadKey();
            }
        }
    }
}
