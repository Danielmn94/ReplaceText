using DocumentFormat.OpenXml.Drawing;
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
            //Gets date and year from DateTime
            string date = DateTime.Today.ToString("dd-MM-yyyy");
            string year = DateTime.Today.Year.ToString();

            //Getting user input
            Console.WriteLine("Enter new word to replace key (This is going to be the name of the new folder as well)");
            var userInput = Console.ReadLine();

            string newfolderLocation = @"..\Desktop\folderlocation";

            //Where the new folder is going to be created
            string newDir = newfolderLocation + @"\" + userInput;

            //If the folder doesn't exist it's going to create it, otherwise to program will exit
            if (!Directory.Exists(newDir))
            {
                //Creating the new directory
                Directory.CreateDirectory(newDir);

                //In this case there are room for two documents
                string copyFileSource = @"..\Desktop\firstdocumentlocation";
                string fileName = @"..\Desktop\whatfiletoselect";

                string copyFileSource1 = @"..\Desktop\seconddocumentlocation";
                string fileName1 = @"..\Desktop\whatfiletoselect";

                //What keys to replace in the document
                string[] replacementKeys = new string[] { "key1", "key2" };
                string docText = "";

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

                    //Replace keys in the first word document
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
            else
            {
                Console.WriteLine("Folder exists, no replacement done");                    
            }
        }
    }
}
