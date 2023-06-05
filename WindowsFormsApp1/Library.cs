using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static WindowsFormsApp1.FrontPage;

namespace TimerLibrary
{
    public class Competition
    {
        public Competition()
        {
            this.Team = new List<TEAM>();
        }
        public List<TEAM> Team { get; set; }
    }
    public class TEAM : IComparable<TEAM>
    {
        public TEAM(string _Name)   //建構子
        {
            this.Name = _Name;
            this.Time = new List<TimerData>();
            this.MazeTime = new List<TimerData>();
        }
        public List<TimerData> Time { get; set; }//屬性
        public List<TimerData> MazeTime { get; set; }//屬性
        public string Oeder { get; set; }//屬性
        public string Name { get; set; }//屬性
        public string ID { get; set; }//屬性
        public string Organize { get; set; }//屬性
        private int Minimum_time_ms = 2147483647;
        public int Minimum
        {
            get
            {
                if(Minimum_time_ms != 2147483647)
                    return Minimum_time_ms;
                else
                    return 0;
            }
            set
            {
                if (value < Minimum_time_ms) Minimum_time_ms = value;
            }
        }
        public override string ToString()
        {
            string T_ms = Convert.ToString(Minimum % 1000);
            T_ms = T_ms.PadLeft(3, '0');

            string T_Second = Convert.ToString((Minimum % 60000) / 1000);
            T_Second = T_Second.PadLeft(2, '0');

            string T_minute = Convert.ToString((int)Minimum / 60000);
            T_minute = T_minute.PadLeft(2, '0');

            string Time = T_minute + ':' + T_Second + '.' + T_ms;
            return Time;
        }
        public int CompareTo(TEAM other)
        {
            if (other == null)
                return 1;
            else
                return this.Minimum.CompareTo(other.Minimum);
        }
    }
    public class TimerData : IComparable<TimerData>
    {
        private readonly Regex ReScore = new Regex(@"\s*(\d{1,2})\s*:\s*(\d{1,2})\s*.\s*(\d{1,3})", RegexOptions.IgnoreCase);
        public TimerData(int time_ms)
        {
            this.mSec = time_ms;
        }
        public TimerData(string str)
        {
            int Millisecond = 0;
            MatchCollection matches = ReScore.Matches(str);
            foreach (Match match in matches)// 一一取出 MatchCollection 內容
            {// 將 Match 內所有值的集合傳給 GroupCollection groups
                GroupCollection groups = match.Groups;
                Millisecond = Convert.ToInt32(groups[1].Value.Trim()) * 60 * 1000 + Convert.ToInt32(groups[2].Value.Trim()) * 1000 + Convert.ToInt32(groups[3].Value.Trim());
                break;
            }
            this.mSec = Millisecond;
        }
        public int mSec { get; set; }
        public string Score()
        {

            return "";
        }
        public int CompareTo(TimerData other)
        {
            if (other == null)
                return 1;
            else
                return this.mSec.CompareTo(other.mSec);
        }
        public override string ToString()
        {
            string T_ms = Convert.ToString(mSec % 1000);
            T_ms = T_ms.PadLeft(3, '0');

            string T_Second = Convert.ToString((mSec % 60000) / 1000);
            T_Second = T_Second.PadLeft(2, '0');

            string T_minute = Convert.ToString((int)mSec / 60000);
            T_minute = T_minute.PadLeft(2, '0');

            string Time = T_minute + ':' + T_Second + '.' + T_ms;
            return Time;
        }
    }
}
