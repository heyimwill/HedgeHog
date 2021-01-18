using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//using System.Collections.Generic;

namespace AccelerometerProcessing_FiveSecond
{
    class TeacherFolderDateTime
    {
        public string TeacherFolder { get; set; }
        public string StartDate { get; set; }
        public string StartTime { get; set; }
        public string EndDate { get; set; }
        public string EndTime { get; set; }

        public TeacherFolderDateTime(string tf, string sd, string st, string ed, string et)
        {
            TeacherFolder = tf;
            StartDate = sd;
            StartTime = st;
            EndDate = ed;
            EndTime = et;
        }
    }
}
