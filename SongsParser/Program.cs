using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace SongsParser
{
    class Program
    {
        static void Main(string[] args)
        {
            var text = GetTextFromWord(@"D:\Repositories\songsTest.doc");
            var songTexts = PreprocessingToParsing(text);
            var songs = songTexts.Select(s => new Song()
            {
                Id = int.Parse(string.Concat(s.TakeWhile(p => p != '\r'))),
                Text = string.Concat(s.SkipWhile(p => p != '\n'))
            }).ToList();

            RemoveUnnecessarySymbolsAndSetName(songs);



        }

        private static IEnumerable<string> PreprocessingToParsing(string text)
        {
            var songTexts = text.Split('№').Select(s => s.Remove(0, 2)).Skip(1);
            return songTexts;
        }

        private static void RemoveUnnecessarySymbolsAndSetName(IEnumerable<Song> songs)
        {
            foreach (var song in songs)
            {
                song.Text=string.Concat(song.Text.Skip(6));
                if (song.Name == null)
                    song.Name = string.Concat(song.Text.TakeWhile(p => p != '\n'));
            }
        }

        /// <summary>  
        /// Reading Text from Word document  
        /// </summary>  
        /// <returns></returns>  
        private static string GetTextFromWord(string path)
        {
            StringBuilder text = new StringBuilder();
            Application word = new Application();
            object miss = System.Reflection.Missing.Value;
            object pathObj = path;
            object readOnly = true;
            Document docs = word.Documents.Open(ref pathObj, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

            for (int i = 0; i < docs.Paragraphs.Count; i++)
            {
                var paragraph = " \r\n " + docs.Paragraphs[i + 1].Range.Text.ToString();
                Console.WriteLine(paragraph);
                text.Append(paragraph);
            }

            return text.ToString();
        }
    }
}
