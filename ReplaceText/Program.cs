using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace ReplaceText
{
    class Program
    {
        static void Main(string[] args)
        {
            string date = DateTime.Today.ToString("dd-MM-yyyy");
            string newfolderLocation = @"..\Desktop\wheretocreatenewfolder";

            string copyFileSource = @"..\Desktop\youroriginalwordfile.docx";
            string fileName = @"..\Desktop\nameofnewwordfile.docx";

            string copyFileSource1 = @"..\Desktop\yoursecondoriginalwordfile.docx";
            string fileName1 = @"..\Desktop\nameofsecondnewwordfile.docx";

            string[] replacementKeys = new string[] { "Key1", "Key2" };
            string docText = "";

            //Getting user input
            Console.WriteLine("Enter new word to replace key (This is going to be the name of the new folder as well)");
            var userInput = Console.ReadLine();

            //Saving new destination for word document in a variable
            string newDir = newfolderLocation + @"\" + userInput;

            //Creating new directory
            Directory.CreateDirectory(newDir);

            //Copying first file to new destination
            File.Copy(copyFileSource, newDir + fileName);
            File.Copy(copyFileSource1, newDir + fileName1);

            //What word file to open
            string wordDoc = newDir + fileName;

            //Opening and replacing words in word document
            using (WordprocessingDocument wordProDoc = WordprocessingDocument.Open(wordDoc, true))
            {
                using (StreamReader sr = new StreamReader(wordProDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                //Replace words in word document
                Regex regexText = new Regex(replacementKeys[0]);
                docText = regexText.Replace(docText, date);

                Regex regexText2 = new Regex(replacementKeys[1]);
                docText = regexText2.Replace(docText, userInput);

                using (StreamWriter sw = new StreamWriter(wordProDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                    Console.WriteLine("Replacement done!");
                }
            }
        }
    }
}
