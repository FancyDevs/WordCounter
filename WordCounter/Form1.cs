using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordCounter
{
    public partial class Form1 : Form
    {
        string sFileName = "";
        public static List<WordsClass> listOfWords = new List<WordsClass>();

        public Form1()
        {
            InitializeComponent();
        }

        //TODO do you need this guy????
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void userPath_DoubleClick(object sender, EventArgs e)
        {

            OpenFileDialog dialogBox = new OpenFileDialog();
            dialogBox.Filter = "Word Docs (*.doc*)|*.docx*";
            dialogBox.FilterIndex = 1;
            dialogBox.Multiselect = false;

            //TODO before using try/catch try and check if userPath.Text is blank or null
            //also try to only use try/catch for a specific exception other whys you might not see any other bugs that pop up
            try
            {
                dialogBox.InitialDirectory = System.IO.Path.GetDirectoryName(userPath.Text);
            }
            catch (Exception)
            {
                dialogBox.InitialDirectory = @"C:\Users\alanw\Documents";
            }

            if (dialogBox.ShowDialog() == DialogResult.OK)
            {
                sFileName = dialogBox.FileName;
                userPath.Text = System.IO.Path.GetFileNameWithoutExtension(sFileName);
            }
        }

        public void goButton_Click(object sender, EventArgs e)
        {
            CountWords(sFileName);
            WriteResults(listOfWords);
        }

        private void WriteResults(List<WordsClass> listOfWordsInput)
        {
            resultScreen.Text = "";
            string tempResult = "";

            tempResult = "WORD".PadRight(20, ' ') + "| COUNT";

            //listOfWordsInput = listOfWordsInput.OrderBy(x => x.totalCount);

            foreach (WordsClass item in listOfWordsInput.OrderByDescending(x => x.totalCount))
            {
                tempResult += "\r\n";
                tempResult += item.nameOfWord.PadRight(20, ' ') + "| " + item.totalCount.ToString();
            }

            resultScreen.Text = tempResult;

        }

        public static void CountWords(string pathIn)
        {
            listOfWords.Clear();

            //TODO can you change this to a using instead of naming it all out here
            //TODO also you can replace the left side with var word to simplify it
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object miss = System.Reflection.Missing.Value;
            object readOnly = true;
            object pathForWord = pathIn;
            
            //TODO same down here replace with var doc
            Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(ref pathForWord, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss,
                ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

            string line = "";
            for (int i = 0; i < doc.Paragraphs.Count; i++)
            {
                line = doc.Paragraphs[i + 1].Range.Text.ToString();

                line = RemoveSpecialCharacters(line);

                string[] splitLine = line.Split(' ');
                int result = 0;

                foreach (string text in splitLine)
                {
                    result = CheckNames(listOfWords, text);
                    if (result > 0)
                    {
                        listOfWords[result].totalCount++;
                    }
                    else
                    {
                        listOfWords.Add(new WordsClass { nameOfWord = text.ToLower(), totalCount = 1 });
                    }
                }
            }

            doc.Close();
            word.Quit();
        }

        private static string RemoveSpecialCharacters(string input)
        {
            StringBuilder newString = new StringBuilder();

            foreach (char c in input)
            {
                if (char.IsLetterOrDigit(c) || c == '\'' || c == ' ')
                {
                    newString.Append(c);
                }
            }

            return newString.ToString();
        }

        public static int CheckNames(List<WordsClass> listToCheck, string nameToCheck)
        {
            int foundValue = 0;

            for (int i = 0; i < listToCheck.Count; i++)
            {
                if (string.Equals(listToCheck[i].nameOfWord.ToLower(), nameToCheck.ToLower()))
                {
                    foundValue = i;
                    break;
                }
            }

            return foundValue;
        }

        private void userPath_TextChanged(object sender, EventArgs e)
        {
            //TODO Can simplify this to goButton.Enabled = !string.IsNullOrEmpty(userPath.Text)
            if (string.IsNullOrEmpty(userPath.Text))
            {
                goButton.Enabled = false;
            }
            else
            {
                goButton.Enabled = true;
            }
        }
    }
}
