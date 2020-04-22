using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Vocabulary
{
   class ReadWord
    {
        public Word MakeWord(string En,string Cn)
        {
            Word Newword = new Word(En,Cn);
            return Newword;
        }
           
    }
}
