﻿using OfficeOpenXml;
using SpecAdvisor;
using System.Text;

namespace SpecAdvisor
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.InputEncoding = Encoding.UTF8;
            Console.OutputEncoding = Encoding.UTF8;


            List<Faculty> faculties = new List<Faculty>();

            string excelFilePath = @"C:\Users\Gulshan Azimli\Desktop\journals\2023-2024copy.xlsx";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo excelFile = new FileInfo(excelFilePath);

            using (ExcelPackage package = new ExcelPackage(excelFile))
            {
                int id = 1;
                string universityName = "";
                string groupName = "";
                for (int i = 0; i < package.Workbook.Worksheets.Count; i++)
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[i]; // Change the index if needed

                    int rowCount = worksheet.Dimension.Rows;
                    //int columnCount = worksheet.Dimension.Columns;


                    for (int row = 1; row <= rowCount; row++)
                    {
                        if (worksheet.Cells[row, 2].Value == null && worksheet.Cells[row, 1].Value != null)
                        {
                            string value = Replace(worksheet.Cells[row, 1].Value.ToString());
                            if (value.Contains("qrup"))
                            {
                                if (value.Length <= 8)
                                {
                                    int index = value.IndexOf("qrup");
                                    groupName = value.Substring(0, index - 1);

                                }
                                else
                                {
                                    universityName = "";
                                    string[] str = value.Split();
                                    groupName = str[str.Length - 2];
                                    for (int j = 0; j < str.Length - 2; j++)
                                    {
                                        universityName += str[j];
                                        if (j == str.Length - 3)
                                            continue;
                                        universityName += " ";
                                    }
                                }
                            }
                            else
                            {
                                universityName = value;
                            }
                        }
                        else
                        {
                            if (worksheet.Cells[row, 1].Value == null)
                            {
                                string text = Replace(worksheet.Cells[row, 2].GetValue<string>());
                                faculties[id - 2].Name += " " + text;
                                continue;
                            }
                            Faculty faculty = new Faculty();
                            faculty.Id = id++;
                            faculty.Name = Replace(worksheet.Cells[row, 2].GetValue<string>());
                            faculty.IsVisual = worksheet.Cells[row, 3].GetValue<string>() == "Q" ? false : true;
                            faculty.GroupName = groupName;
                            faculty.University = new University();
                            faculty.University.UniversityName = universityName;
                            var scoreValues = worksheet.Cells[row, 4].GetValue<string>();
                            string[] scores = scoreValues.Split('(', ')');
                            if (scores.Length > 1)
                            {
                                if (scores[0].Contains('/'))
                                {
                                    int index = scores[0].IndexOf('/');
                                    scores[0] = scores[0].Substring(index + 1);
                                }
                                //else if (scores[0].Contains('-'))
                                //{
                                //    scores[0] = "0";
                                //}
                                faculty.ScoreWithPay = Convert.ToDouble(scores[0]);
                                if (scores[1].Contains('/'))
                                {
                                    int index = scores[1].IndexOf('/');
                                    scores[1] = scores[1].Substring(index + 1);
                                }
                                else if (scores[1].Contains('-'))
                                {
                                    scores[1] = "0";
                                }
                                faculty.Score = Convert.ToDouble(scores[1]);
                            }
                            else
                            {
                                string score = worksheet.Cells[row, 4].Value.ToString();
                                if (score.Contains('/'))
                                {
                                    int index = score.IndexOf('/');
                                    score = score.Substring(index + 1);
                                }
                                faculty.Score = Convert.ToDouble(score);
                            }

                            faculties.Add(faculty);

                        }
                    }
                }
            }

            //foreach (var f in faculties)
            //{
            //    Console.WriteLine($"{f.Id}.{f.University.UniversityName} {f.Name}  {f.GroupName} qrup  {f.Score}/{f.ScoreWithPay}");
            //}

            Console.WriteLine("Please enter your score...");
            double studentScore = Convert.ToDouble(Console.ReadLine());

            //Console.WriteLine("Please, enter university name or names which separated with comma...");
            string[] universityNames =  { "Bakı Dövlət Universiteti", "Azərbaycan Texniki Universiteti",
                "Azərbaycan Dövlət Neft və Sənaye Universiteti", "Azərbaycan Memarlıq və İnşaat Universiteti" ,
                "Azərbaycan Dövlət Pedaqoji Universiteti" , "Azərbaycan Dövlət İqtisad Universiteti" ,
                "Bakı Mühəndislik Universiteti" };

            Console.WriteLine("The faculty is visible?");
            string isFormal = Console.ReadLine();
            bool isVisual = isFormal == "Y" ? true : false;

            Console.WriteLine("The faculty is non-visible?");
            string isNotFormal = Console.ReadLine();
            bool isNotVisual = isNotFormal == "Y" ? true : false;


            Console.WriteLine("Should the faculty be free?");
            bool isFree = Console.ReadLine() == "Y" ? true : false;

            Console.WriteLine("Should the faculty be paid?");
            bool isPaid = Console.ReadLine() == "Y" ? true : false;

            Console.WriteLine("What group is  the faculty?");
            string group = Console.ReadLine();

            Console.WriteLine("How many choice do you want to see?");
            int totalChoices = Convert.ToInt32(Console.ReadLine());
            int countOfChoices = totalChoices / 5;

            List<Faculty> chosenFaculties = faculties.Where(f => (f.IsVisual == isVisual) || (!(f.IsVisual) == isNotVisual)).ToList();
            chosenFaculties = chosenFaculties.Where(f => universityNames.Contains(f.University.UniversityName) && f.GroupName == group).ToList();
            double plusFifty = studentScore + 50;
            double plusTwentyFive = studentScore + 25;
            double minusTwentyFive = studentScore - 25;
            double minusFifty = studentScore - 50;
            double minusEighty = studentScore - 80;
            List<Faculty> isBetweenPlus25Plus50 = new List<Faculty>();
            List<Faculty> isBetweenScorePlus25 = new List<Faculty>();
            List<Faculty> isBetweenMinus25Score = new List<Faculty>();
            List<Faculty> isBetweenMinus25Minus50 = new List<Faculty>();
            List<Faculty> isBetweenMinus50Minus80 = new List<Faculty>();

            if(isPaid)
            {
                isBetweenPlus25Plus50 = chosenFaculties.Where(f => (f.ScoreWithPay >= plusTwentyFive & f.ScoreWithPay <= plusFifty)).ToList();
                isBetweenScorePlus25 = chosenFaculties.Where(f => (f.ScoreWithPay >= studentScore & f.ScoreWithPay <= plusTwentyFive)).ToList();
                isBetweenMinus25Score = chosenFaculties.Where(f => (f.ScoreWithPay >= minusTwentyFive & f.ScoreWithPay <= studentScore)).ToList();
                isBetweenMinus25Minus50 = chosenFaculties.Where(f => (f.ScoreWithPay >= minusFifty & f.ScoreWithPay <= minusTwentyFive)).ToList();
                isBetweenMinus50Minus80 = chosenFaculties.Where(f => (f.ScoreWithPay >= minusEighty & f.ScoreWithPay <= minusFifty)).ToList();

            }
            if(isFree)
            {
                var isBetweenPlusFivePlusTwentyIsFree = chosenFaculties.Where(f => (f.Score >= plusTwentyFive & f.Score <= plusFifty)).ToList();
                isBetweenPlus25Plus50.AddRange(isBetweenPlusFivePlusTwentyIsFree);

                var isBetweenPlusFiveMinusFiveIsFree = chosenFaculties.Where(f => (f.Score >= studentScore & f.Score <= plusTwentyFive)).ToList();
                isBetweenScorePlus25.AddRange(isBetweenPlusFiveMinusFiveIsFree);

                var isBetweenMinusFiveMinusTwentyIsFree = chosenFaculties.Where(f => (f.Score >= minusTwentyFive & f.Score <= studentScore)).ToList();
                isBetweenMinus25Score.AddRange(isBetweenMinusFiveMinusTwentyIsFree);

                var isBetweenMinusTwentyMinusFiftyIsFree = chosenFaculties.Where(f => (f.Score >= minusFifty & f.Score <= minusTwentyFive)).ToList();
                isBetweenMinus25Minus50.AddRange(isBetweenMinusTwentyMinusFiftyIsFree);

                var isBetweenMinusFiftyMinusEightyIsFree = chosenFaculties.Where(f => (f.Score >= minusEighty & f.Score <= minusFifty)).ToList();
                isBetweenMinus50Minus80.AddRange(isBetweenMinusFiftyMinusEightyIsFree);
            }

            List<Faculty> facultiesForStudent = new List<Faculty>();

            int isBetweenPlus25Plus50Count = isBetweenPlus25Plus50.Count;  //24
            int isBetweenScorePlus25Count = isBetweenScorePlus25.Count;    //30
            int isBetweenMinus25ScoreCount = isBetweenMinus25Score.Count;   //26
            int isBetweenMinus25Minus50Count = isBetweenMinus25Minus50.Count;
            int isBetweenMinus50Minus80Count = isBetweenMinus50Minus80.Count;

            int sumOfFirstThree = totalChoices * 3 / 5;
            if (isBetweenMinus25Minus50Count + isBetweenMinus50Minus80Count < (totalChoices * 2 / 5))
                sumOfFirstThree = totalChoices - (isBetweenMinus25Minus50Count + isBetweenMinus50Minus80Count);

            if (isBetweenScorePlus25Count > countOfChoices)
            {
                if (isBetweenPlus25Plus50Count < countOfChoices && isBetweenMinus25ScoreCount < countOfChoices)
                {
                    int sum = isBetweenPlus25Plus50Count + isBetweenScorePlus25Count + isBetweenMinus25ScoreCount;
                    if (sum > sumOfFirstThree)
                        isBetweenScorePlus25Count = sumOfFirstThree - (isBetweenPlus25Plus50Count + isBetweenMinus25ScoreCount);
                }
                else if (isBetweenPlus25Plus50Count < countOfChoices && isBetweenMinus25ScoreCount > countOfChoices)
                {
                    int remain = sumOfFirstThree - isBetweenPlus25Plus50Count;
                    if (isBetweenMinus25ScoreCount < remain / 2)
                        isBetweenScorePlus25Count = remain - isBetweenMinus25ScoreCount;
                    else if (isBetweenScorePlus25Count < remain / 2)
                        isBetweenMinus25ScoreCount = remain - isBetweenScorePlus25Count;
                    else
                    {
                        isBetweenMinus25ScoreCount = remain / 2;
                        isBetweenScorePlus25Count = remain - isBetweenMinus25ScoreCount;
                    }
                }
                else if (isBetweenPlus25Plus50Count > countOfChoices && isBetweenMinus25ScoreCount < countOfChoices)
                {
                    int remain = sumOfFirstThree - isBetweenMinus25ScoreCount;

                    if (isBetweenPlus25Plus50Count < remain / 2)
                        isBetweenScorePlus25Count = remain - isBetweenPlus25Plus50Count;
                    else if (isBetweenScorePlus25Count < remain / 2)
                        isBetweenPlus25Plus50Count = remain - isBetweenScorePlus25Count;
                    else
                    {
                        isBetweenPlus25Plus50Count = remain / 2;
                        isBetweenScorePlus25Count = remain - isBetweenPlus25Plus50Count;
                    }
                }
                else
                {
                    int sumOfOne = sumOfFirstThree / 3;
                    isBetweenPlus25Plus50Count = isBetweenMinus25ScoreCount = sumOfOne;
                    isBetweenScorePlus25Count = sumOfFirstThree - (isBetweenPlus25Plus50Count + isBetweenMinus25ScoreCount);
                }
            }
            if (isBetweenScorePlus25Count < countOfChoices)
            {
                if (isBetweenPlus25Plus50Count > countOfChoices && isBetweenMinus25ScoreCount > countOfChoices)
                {
                    int remain = sumOfFirstThree - isBetweenScorePlus25Count;

                    if ((isBetweenPlus25Plus50Count > remain / 2) && (isBetweenMinus25ScoreCount > remain / 2))
                    {
                        isBetweenPlus25Plus50Count = remain / 2;
                        isBetweenMinus25ScoreCount = remain - isBetweenPlus25Plus50Count;
                    }
                    else if ((isBetweenPlus25Plus50Count < remain / 2) && (isBetweenMinus25ScoreCount > remain / 2))
                    {
                        if (isBetweenMinus25ScoreCount > (remain - isBetweenPlus25Plus50Count))
                            isBetweenMinus25ScoreCount = remain - isBetweenPlus25Plus50Count;
                    }
                    else if ((isBetweenMinus25ScoreCount < remain / 2) && (isBetweenPlus25Plus50Count > remain / 2))
                    {
                        if (isBetweenPlus25Plus50Count > (remain - isBetweenMinus25ScoreCount))
                            isBetweenPlus25Plus50Count = remain - isBetweenMinus25ScoreCount;
                    }
                }
                else if (isBetweenPlus25Plus50Count < countOfChoices && isBetweenMinus25ScoreCount > countOfChoices)
                {
                    int remain = sumOfFirstThree - (isBetweenPlus25Plus50Count + isBetweenScorePlus25Count);
                    if (isBetweenMinus25ScoreCount > remain)
                        isBetweenMinus25ScoreCount = remain;
                }
                else if (isBetweenPlus25Plus50Count > countOfChoices && isBetweenMinus25ScoreCount < countOfChoices)
                {
                    int remain = sumOfFirstThree - (isBetweenMinus25ScoreCount + isBetweenScorePlus25Count);
                    if (isBetweenPlus25Plus50Count > remain)
                        isBetweenPlus25Plus50Count = remain;
                }
            }

            int sumOfLastTwo = totalChoices - (isBetweenPlus25Plus50Count + isBetweenScorePlus25Count + isBetweenMinus25ScoreCount);
            if (isBetweenMinus25Minus50Count < countOfChoices & isBetweenMinus50Minus80Count > countOfChoices)
            {
                if (isBetweenMinus50Minus80Count > sumOfLastTwo - isBetweenMinus25Minus50Count)
                    isBetweenMinus50Minus80Count = sumOfLastTwo - isBetweenMinus25Minus50Count;
            }
            else if (isBetweenMinus25Minus50Count > countOfChoices & isBetweenMinus50Minus80Count < countOfChoices)
            {
                if (isBetweenMinus25Minus50Count > sumOfLastTwo - isBetweenMinus50Minus80Count)
                    isBetweenMinus25Minus50Count = sumOfLastTwo - isBetweenMinus50Minus80Count;
            }
            else if (isBetweenMinus25Minus50Count > countOfChoices & isBetweenMinus50Minus80Count > countOfChoices)
            {
                isBetweenMinus25Minus50Count = sumOfLastTwo / 2;
                isBetweenMinus50Minus80Count = sumOfLastTwo - isBetweenMinus25Minus50Count;
            }

            for (int i = 0; i < isBetweenPlus25Plus50.Count; i++)
            {
                facultiesForStudent.Add(isBetweenPlus25Plus50[i]);
            }
            for (int i = 0; i < isBetweenScorePlus25.Count; i++)
            {
                facultiesForStudent.Add(isBetweenScorePlus25[i]);
            }
            for (int i = 0; i < isBetweenMinus25Score.Count; i++)
            {
                facultiesForStudent.Add(isBetweenMinus25Score[i]);
            }
            for (int i = 0; i < isBetweenMinus25Minus50.Count; i++)
            {
                facultiesForStudent.Add(isBetweenMinus25Minus50[i]);
            }
            for (int i = 0; i < isBetweenMinus50Minus80.Count; i++)
            {
                facultiesForStudent.Add(isBetweenMinus50Minus80[i]);
            }

            foreach (var f in facultiesForStudent)
            {
                Console.WriteLine($"{f.Id}.{f.University.UniversityName} {f.Name} {f.IsVisual} {f.GroupName} qrup  {f.Score}/{f.ScoreWithPay}");
            }
            //foreach (var f in isBetweenPlusFivePlusTwenty)
            //{
            //    Console.WriteLine($"{f.Id}.{f.University.UniversityName} {f.Name} {f.IsVisual} {f.GroupName} qrup  {f.Score}/{f.ScoreWithPay}");
            //}
            //Console.WriteLine("************************************************************************************************");
            //foreach (var f in isBetweenPlusFiveMinusFive)
            //{
            //    Console.WriteLine($"{f.Id}.{f.University.UniversityName} {f.Name} {f.IsVisual} {f.GroupName} qrup  {f.Score}/{f.ScoreWithPay}");
            //}
            //Console.WriteLine("************************************************************************************************");
            //foreach (var f in isBetweenMinusFiveMinusTwenty)
            //{
            //    Console.WriteLine($"{f.Id}.{f.University.UniversityName} {f.Name} {f.IsVisual} {f.GroupName} qrup  {f.Score}/{f.ScoreWithPay}");
            //}
            //Console.WriteLine("************************************************************************************************");
            //foreach (var f in isBetweenMinusTwentyMinusFifty)
            //{
            //    Console.WriteLine($"{f.Id}.{f.University.UniversityName} {f.Name} {f.IsVisual} {f.GroupName} qrup  {f.Score}/{f.ScoreWithPay}");
            //}
            //Console.WriteLine("************************************************************************************************");
            //foreach (var f in isBetweenMinusFiftyMinusEighty)
            //{
            //    Console.WriteLine($"{f.Id}.{f.University.UniversityName} {f.Name} {f.IsVisual} {f.GroupName} qrup  {f.Score}/{f.ScoreWithPay}");
            //}



        }

        static string? Replace(string str)
        {
            return str == null ? null : str.Replace('ӂ', 'ə').Replace('ˬ', 'ə')
                       .Replace('ø', 'İ').Replace('$', 'ş')
                       .Replace('Õ', 'ı').Replace('÷', 'ğ')
                       .Replace('ú', 'ş').Replace('Ӂ', 'Ə')
                       .Replace('6', 'ə');
        }
    }
}
