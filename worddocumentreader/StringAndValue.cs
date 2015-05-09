using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace worddocumentreader
{
    class StringAndValue
    {

        private string sentence_in_word_doc;

        private string string_to_search;

        private string [] _currencySymbols = new string [] { "$","zł", "€", "£" };

        public string SentenceInWordDoc
        {
            get
            {
                return sentence_in_word_doc;
            }
            set
            {
                sentence_in_word_doc = value;
            }
        }

        public string StringToSearch
        {
            get
            {
                return string_to_search;
            }
        }

        public StringAndValue(string sentence, string tosearch)
        {
            this.sentence_in_word_doc = sentence;
            this.string_to_search = tosearch;
        }

        public float ExtractStringAndValue()
        {
            float Value = 0;
            char [] delimiters = new [] { ',', ';', ' ','.' };
            if (_currencySymbols.Where(x => SentenceInWordDoc.Contains(x)).Count() > 0)
            {
                foreach (string curr in _currencySymbols)
                {
                    if (SentenceInWordDoc.Contains(curr))
                    {
                        string NewString = SentenceInWordDoc.Replace(curr, "");
                        SentenceInWordDoc = NewString;
                    }
                }
            }
      
            string [] SplittedArray = SentenceInWordDoc.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < SplittedArray.Length; i++)
            {
                if (SplittedArray[i].ToLower().Equals(string_to_search))
                {
                    Value = GetNumber(SplittedArray,SplittedArray.Length,i,1);
                    break;
                }
            }
            return Value;
        }


        //The value can be either anywhere before or after the string. Hence look for the first number
        // before or after the string and stop
        private float GetNumber(string [] Array, int length, int CurrentIndex, int i)
        {
            float n;
            float ReturnValue = 0;
            bool ValFound = false;

            if(((CurrentIndex - i) >= 0) && (float.TryParse(Array[CurrentIndex-i], out n)))          
            {
                  ReturnValue = float.Parse(Array[CurrentIndex-i]);
                  ValFound = true;
            }
            else if(((CurrentIndex + i) < length) && (float.TryParse(Array[CurrentIndex + i], out n)))
            {
                ReturnValue = float.Parse(Array[CurrentIndex + i]);
                ValFound = true;               
            }
            if (!ValFound && (i < length))
            {
                GetNumber(Array, length, CurrentIndex, i+1);
            }

            return ReturnValue;
        }

    }
}
