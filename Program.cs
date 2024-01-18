using System;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        // Get the file name from the user
        Console.WriteLine("Start ");
        string fileName = "Data/input.txt";

        // Call the method to search for the phrase and store positions in an array
        Console.WriteLine("Start:SearchPhrase1");
        int[] positions = SearchPhraseInFile(fileName, "Tydzień");
        Console.WriteLine("End:SearchPhare1");


        // Print the positions to the console
        Console.WriteLine("Start:PLExtracting");
        int[,] PL = PlExtracting(fileName, positions);
        Console.WriteLine("End: PlExtracting");

        Console.WriteLine("Start: PLDiving");
        int[,] PL2 = PlDiving(PL);
        Console.WriteLine("END: PLDiving");

        Console.WriteLine("start:WeekCounter");
        string[] Weeks = WeekCounter(positions,fileName);
        Console.WriteLine("Start: WeekCounter");

        Console.WriteLine("Start:Gen.Res");
        GenRes(Weeks, PL2);
        Console.WriteLine("End: Gen.Res");


       // Console.ReadKey();
    }

    static int[] SearchPhraseInFile(string fileName, string targetPhrase)
    {
        List<int> positions = new List<int>();

        try
        {
            using (StreamReader reader = new StreamReader(fileName))
            {
                int lineNumber = 0;

                while (!reader.EndOfStream)
                {
                    lineNumber++;
                    string line = reader.ReadLine();

                    if (line.Contains(targetPhrase))
                    {
                        // If the line contains the target phrase, add the line number to the array
                        positions.Add(lineNumber);
                    }
                }
            }
        }
        catch (Exception e)
        {
            Console.WriteLine($"An error occurred: {e.Message}");
        }

        return positions.ToArray();
    }

    static int[,] PlExtracting(string fileName, int[] positions)
    {

        int Xpos = 0;
        int[,] Res = new int[54, 120];
        foreach (int position in positions)
        {
            Xpos++;
            try
            {
                using (StreamReader reader = new StreamReader(fileName))
                {
                    int lineNumber = 0;
                    int startLine = position + 3;
                    int endLine = positions[Xpos];
                    int x = 0;
                    Console.WriteLine(position);
                    while (!reader.EndOfStream)
                    {
                        lineNumber++;
                        string lineAsString = reader.ReadLine();

                        // Check if the current line is within the specified range
                        if (lineNumber >= startLine && lineNumber < endLine)
                        {
                            double doubleline = double.Parse(lineAsString);
                            int intline = (int)Math.Round(doubleline);

                            Res[Xpos - 1, x] = intline;
                            x++;


                        }

                        // Break the loop if we have reached the end line
                        if (lineNumber == endLine)
                        {
                            break;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(Xpos);
                Console.WriteLine($"An error occurred: {e.Message}");
                break;
            }


        }
        return Res;
    }

    static string[] WeekCounter(int[] pos,string fileName)
    {
        List<string> lines = new List<string>();

        try
        {
            using (StreamReader reader = new StreamReader(fileName))
            {
                int currentRow = 0;
                int i = 0;
                while (!reader.EndOfStream )
                {
                    string line = reader.ReadLine();
                    

                    if (currentRow == pos[i]-1)
                    {
                        lines.Add(line);
                        Console.WriteLine(line) ;
                        i++;
                    }
                    currentRow++;
                }
            }
        }
        catch (Exception e)
        {
            Console.WriteLine($"An error occurred: {e.Message}");
        }
        return lines.ToArray();
    }

    static int[,] PlDiving(int[,] pl)
    {
        int[,] res = new int[54, 40];
        for (int i = 0; i < 54; i++)
        {
            int i3 = 0;
            for (int i2 = 0; i2 <120 ; i2=i2+3)
            {
                res[i, i3] = pl[i, i2];
                i3++;
            }
        }
        return res;

    }

    static void GenRes(string[] w, int[,] pl)
    {
        // Set the LicenseContext to NonCommercial
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        string path = "Data/output.xlsx";
        if (File.Exists(path))
        {
            // Generate a new file name to avoid overwriting
            path = GenerateNewFileName(path);
        }
        FileInfo fileInfo = new FileInfo(path);

        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Result");

            for (int i = 1; i < w.GetLength(0); i++)
            {
                w[i-1] = w[i-1].Replace("Tydzień ", string.Empty);
                worksheet.Cells[1, i].Value = w[i-1];
            }
            int x = pl.GetLength(0);
            for (int i2 = 1; i2 < pl.GetLength(1); i2++)
            {
                for (int i3 = 1; i3 < pl.GetLength(0); i3++)
                {
                    worksheet.Cells[i2+1, i3].Value = pl[i3-1,i2-1];
                }
            }
            package.Save();
        }
    }
    static string GenerateNewFileName(string existingFileName)
    {
        // Extract the file extension
        string fileExtension = Path.GetExtension(existingFileName);

        // Generate a new file name with a timestamp
        string newFileName = Path.GetFileNameWithoutExtension(existingFileName) +
                             "_" + DateTime.Now.ToString("HH_mm_ss") +
                             fileExtension;

        // Combine with the original directory
        return Path.Combine(Path.GetDirectoryName(existingFileName), newFileName);
    }
}


