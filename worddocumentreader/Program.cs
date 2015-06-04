using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Novacode;
using System.IO;
using System.Reflection;


namespace worddocumentreader
{
    class Program
    {
        static void Main(string[] args)
        {

           // System.Diagnostics.Debugger.Launch();

            if (args.Length < 1)
            {
                Print_Help();
                Console.ReadKey();

            }           
            Dictionary<string, float> TextVsValue = new Dictionary<string, float>();

            string input = args[0];

            if (System.IO.File.Exists(input))
            {
                try
                {
                    using (DocX doc = DocX.Load(input))
                    {
                        for (int i = 0; i < doc.Paragraphs.Count; i++)
                        {

                            foreach (var item in doc.Paragraphs[i].Text.Split(new string[] { "\n" }
                                      , StringSplitOptions.RemoveEmptyEntries))
                            {
                                foreach (string searchText in args)
                                {

                                    if ((item.ToLower().Contains(searchText.ToLower())) && (searchText != input))
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
                                        else if (TextVsValue[searchText] < val)
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
                catch(SystemException ex)
                {
                    Console.WriteLine(ex);
                }
            }
            else
            {
                Console.WriteLine("The file " + input + " does not exist");
            }
        }

        private static void Print_Help()
        {
            Assembly myAssembly = Assembly.GetExecutingAssembly();

            string description = "";
            object[] attributes = myAssembly.GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
            if (attributes.Length > 0)
            {
                description = (attributes[0] as AssemblyDescriptionAttribute).Description;
            }
            Console.WriteLine();
            Console.WriteLine("Mandatory command line arguments:");
            Console.WriteLine("");
            Console.WriteLine("The code accepts command line arguments");
            Console.WriteLine("\t 1st argument : Specifies the path to the input file");
            Console.WriteLine("\t 2nd argument : the words to be searched separated by spaces");
            Console.WriteLine("Examples");
            Console.WriteLine("  worddocumentreader.exe D://Sample.docx oranges apples");
        }

    }
}
