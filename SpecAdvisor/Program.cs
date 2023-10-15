using OfficeOpenXml;
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
                "Azərbaycan Dillər Universiteti", "Bakı Mühəndislik Universiteti" };

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

            Console.WriteLine("Ühat group is  the faculty?");
            string group = Console.ReadLine();

            List<Faculty> chosenFaculties = faculties.Where(f => (f.IsVisual == isVisual) || (!(f.IsVisual) == isNotVisual)).ToList();
            chosenFaculties = chosenFaculties.Where(f => universityNames.Contains(f.University.UniversityName) && f.GroupName == group).ToList();
            double plusFive = studentScore + 5;
            double plusTwenty = studentScore + 20;
            double minusFive = studentScore - 5;
            double minusTwenty = studentScore - 20;
            double minusFifty = studentScore - 50;
            double minusEighty = studentScore - 80;
            List<Faculty> isBetweenPlusFivePlusTwenty = new List<Faculty>();
            List<Faculty> isBetweenPlusFiveMinusFive = new List<Faculty>();
            List<Faculty> isBetweenMinusFiveMinusTwenty = new List<Faculty>();
            List<Faculty> isBetweenMinusTwentyMinusFifty = new List<Faculty>();
            List<Faculty> isBetweenMinusFiftyMinusEighty = new List<Faculty>();

            if (isPaid)
            {
                isBetweenPlusFivePlusTwenty = chosenFaculties.Where(f => (f.ScoreWithPay >= plusFive & f.ScoreWithPay <= plusTwenty)).ToList();
                isBetweenPlusFiveMinusFive = chosenFaculties.Where(f => (f.ScoreWithPay >= minusFive & f.ScoreWithPay <= plusFive)).ToList();
                isBetweenMinusFiveMinusTwenty = chosenFaculties.Where(f => (f.ScoreWithPay >= minusTwenty & f.ScoreWithPay <= minusFive)).ToList();
                isBetweenMinusTwentyMinusFifty = chosenFaculties.Where(f => (f.ScoreWithPay >= minusFifty & f.ScoreWithPay <= minusTwenty)).ToList();
                isBetweenMinusFiftyMinusEighty = chosenFaculties.Where(f => (f.ScoreWithPay >= minusEighty & f.ScoreWithPay <= minusFifty)).ToList();

            }
            else if (isFree)
            {
                var isBetweenPlusFivePlusTwentyIsFree = chosenFaculties.Where(f => (f.Score >= plusFive & f.Score <= plusTwenty)).ToList();
                isBetweenPlusFivePlusTwenty.AddRange(isBetweenPlusFivePlusTwentyIsFree);

                var isBetweenPlusFiveMinusFiveIsFree = chosenFaculties.Where(f => (f.Score >= minusFive & f.Score <= plusFive)).ToList();
                isBetweenPlusFiveMinusFive.AddRange(isBetweenPlusFiveMinusFiveIsFree);

                var isBetweenMinusFiveMinusTwentyIsFree = chosenFaculties.Where(f => (f.Score >= minusTwenty & f.Score <= minusFive)).ToList();
                isBetweenMinusFiveMinusTwenty.AddRange(isBetweenMinusFiveMinusTwentyIsFree);

                var isBetweenMinusTwentyMinusFiftyIsFree = chosenFaculties.Where(f => (f.Score >= minusFifty & f.Score <= minusTwenty)).ToList();
                isBetweenMinusTwentyMinusFifty.AddRange(isBetweenMinusTwentyMinusFiftyIsFree);

                var isBetweenMinusFiftyMinusEightyIsFree = chosenFaculties.Where(f => (f.Score >= minusEighty & f.Score <= minusFifty)).ToList();
                isBetweenMinusFiftyMinusEighty.AddRange(isBetweenMinusFiftyMinusEightyIsFree);
            }

            List<Faculty> facultiesForStudent = new List<Faculty>();

            int isBetweenPlusFivePlusTwentyCount = isBetweenPlusFivePlusTwenty.Count;  //24
            int isBetweenPlusFiveMinusFiveCount = isBetweenPlusFiveMinusFive.Count;    //30
            int isBetweenMinusFiveMinusTwentyCount = isBetweenMinusFiveMinusTwenty.Count;   //26
            int isBetweenMinusTwentyMinusFiftyCount = isBetweenMinusTwentyMinusFifty.Count;
            int isBetweenMinusFiftyMinusEightyCount = isBetweenMinusFiftyMinusEighty.Count;

            int sumOfFirstThree = 60;
            if (isBetweenMinusTwentyMinusFiftyCount + isBetweenMinusFiftyMinusEightyCount < 40)
                sumOfFirstThree = 100 - (isBetweenMinusTwentyMinusFiftyCount + isBetweenMinusFiftyMinusEightyCount);

            if (isBetweenPlusFiveMinusFiveCount > 20)
            {
                if (isBetweenPlusFivePlusTwentyCount < 20 && isBetweenMinusFiveMinusTwentyCount < 20)
                {
                    int sum = isBetweenPlusFivePlusTwentyCount + isBetweenPlusFiveMinusFiveCount + isBetweenMinusFiveMinusTwentyCount;
                    if (sum > sumOfFirstThree)
                        isBetweenPlusFiveMinusFiveCount = sumOfFirstThree - (isBetweenPlusFivePlusTwentyCount + isBetweenMinusFiveMinusTwentyCount);
                }
                else if (isBetweenPlusFivePlusTwentyCount < 20 && isBetweenMinusFiveMinusTwentyCount > 20)
                {
                    int remain = sumOfFirstThree - isBetweenPlusFivePlusTwentyCount;
                    if (isBetweenMinusFiveMinusTwentyCount < remain / 2)
                        isBetweenPlusFiveMinusFiveCount = remain - isBetweenMinusFiveMinusTwentyCount;
                    else if (isBetweenPlusFiveMinusFiveCount < remain / 2)
                        isBetweenMinusFiveMinusTwentyCount = remain - isBetweenPlusFiveMinusFiveCount;
                    else
                    {
                        isBetweenMinusFiveMinusTwentyCount = remain / 2;
                        isBetweenPlusFiveMinusFiveCount = remain - isBetweenMinusFiveMinusTwentyCount;
                    }
                }
                else if (isBetweenPlusFivePlusTwentyCount > 20 && isBetweenMinusFiveMinusTwentyCount < 20)
                {
                    int remain = sumOfFirstThree - isBetweenMinusFiveMinusTwentyCount;

                    if (isBetweenPlusFivePlusTwentyCount < remain / 2)
                        isBetweenPlusFiveMinusFiveCount = remain - isBetweenPlusFivePlusTwentyCount;
                    else if (isBetweenPlusFiveMinusFiveCount < remain / 2)
                        isBetweenPlusFivePlusTwentyCount = remain - isBetweenPlusFiveMinusFiveCount;
                    else
                    {
                        isBetweenPlusFivePlusTwentyCount = remain / 2;
                        isBetweenPlusFiveMinusFiveCount = remain - isBetweenPlusFivePlusTwentyCount;
                    }
                }
                else
                {
                    int sumOfOne = sumOfFirstThree / 3;
                    isBetweenPlusFivePlusTwentyCount = isBetweenMinusFiveMinusTwentyCount = sumOfOne;
                    isBetweenPlusFiveMinusFiveCount = sumOfFirstThree - (isBetweenPlusFivePlusTwentyCount + isBetweenMinusFiveMinusTwentyCount);
                }
            }
            if (isBetweenPlusFiveMinusFiveCount < 20)
            {
                if (isBetweenPlusFivePlusTwentyCount > 20 && isBetweenMinusFiveMinusTwentyCount > 20)
                {
                    int remain = sumOfFirstThree - isBetweenPlusFiveMinusFiveCount;

                    if ((isBetweenPlusFivePlusTwentyCount > remain / 2) && (isBetweenMinusFiveMinusTwentyCount > remain / 2))
                    {
                        isBetweenPlusFivePlusTwentyCount = remain / 2;
                        isBetweenMinusFiveMinusTwentyCount = remain - isBetweenPlusFivePlusTwentyCount;
                    }
                    else if ((isBetweenPlusFivePlusTwentyCount < remain / 2) && (isBetweenMinusFiveMinusTwentyCount > remain / 2))
                    {
                        if (isBetweenMinusFiveMinusTwentyCount > (remain - isBetweenPlusFivePlusTwentyCount))
                            isBetweenMinusFiveMinusTwentyCount = remain - isBetweenPlusFivePlusTwentyCount;
                    }
                    else if ((isBetweenMinusFiveMinusTwentyCount < remain / 2) && (isBetweenPlusFivePlusTwentyCount > remain / 2))
                    {
                        if (isBetweenPlusFivePlusTwentyCount > (remain - isBetweenMinusFiveMinusTwentyCount))
                            isBetweenPlusFivePlusTwentyCount = remain - isBetweenMinusFiveMinusTwentyCount;
                    }
                }
                else if (isBetweenPlusFivePlusTwentyCount < 20 && isBetweenMinusFiveMinusTwentyCount > 20)
                {
                    int remain = sumOfFirstThree - (isBetweenPlusFivePlusTwentyCount + isBetweenPlusFiveMinusFiveCount);
                    if (isBetweenMinusFiveMinusTwentyCount > remain)
                        isBetweenMinusFiveMinusTwentyCount = remain;
                }
                else if (isBetweenPlusFivePlusTwentyCount > 20 && isBetweenMinusFiveMinusTwentyCount < 20)
                {
                    int remain = sumOfFirstThree - (isBetweenMinusFiveMinusTwentyCount + isBetweenPlusFiveMinusFiveCount);
                    if (isBetweenPlusFivePlusTwentyCount > remain)
                        isBetweenPlusFivePlusTwentyCount = remain;
                }
            }

            int sumOfLastTwo = 100 - (isBetweenPlusFivePlusTwentyCount + isBetweenPlusFiveMinusFiveCount + isBetweenMinusFiveMinusTwentyCount);
            if (isBetweenMinusTwentyMinusFiftyCount < 20 & isBetweenMinusFiftyMinusEightyCount > 20)
            {
                if (isBetweenMinusFiftyMinusEightyCount > sumOfLastTwo - isBetweenMinusTwentyMinusFiftyCount)
                    isBetweenMinusFiftyMinusEightyCount = sumOfLastTwo - isBetweenMinusTwentyMinusFiftyCount;
            }
            else if (isBetweenMinusTwentyMinusFiftyCount > 20 & isBetweenMinusFiftyMinusEightyCount < 20)
            {
                if (isBetweenMinusTwentyMinusFiftyCount > sumOfLastTwo - isBetweenMinusFiftyMinusEightyCount)
                    isBetweenMinusTwentyMinusFiftyCount = sumOfLastTwo - isBetweenMinusFiftyMinusEightyCount;
            }
            else if (isBetweenMinusTwentyMinusFiftyCount > 20 & isBetweenMinusFiftyMinusEightyCount > 20)
            {
                isBetweenMinusTwentyMinusFiftyCount = sumOfLastTwo / 2;
                isBetweenMinusFiftyMinusEightyCount = sumOfLastTwo - isBetweenMinusTwentyMinusFiftyCount;
            }

            for (int i = 0; i < isBetweenPlusFivePlusTwenty.Count; i++)
            {
                facultiesForStudent.Add(isBetweenPlusFivePlusTwenty[i]);
            }
            for (int i = 0; i < isBetweenPlusFiveMinusFive.Count; i++)
            {
                facultiesForStudent.Add(isBetweenPlusFiveMinusFive[i]);
            }
            for (int i = 0; i < isBetweenMinusFiveMinusTwenty.Count; i++)
            {
                facultiesForStudent.Add(isBetweenMinusFiveMinusTwenty[i]);
            }
            for (int i = 0; i < isBetweenMinusTwentyMinusFifty.Count; i++)
            {
                facultiesForStudent.Add(isBetweenMinusTwentyMinusFifty[i]);
            }
            for (int i = 0; i < isBetweenMinusFiftyMinusEighty.Count; i++)
            {
                facultiesForStudent.Add(isBetweenMinusFiftyMinusEighty[i]);
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
