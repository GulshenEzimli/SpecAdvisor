using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpecAdvisor
{
    public class Faculty
    {
        public Faculty()
        {
            University = new University();
        }
        public int Id { get; set; }
        public string? Name { get; set; }
        public bool IsVisual { get; set; }
        public double Score { get; set; }
        public double ScoreWithPay { get; set; }
        public string? GroupName { get; set; }
        public University University { get; set; }
    }
}
