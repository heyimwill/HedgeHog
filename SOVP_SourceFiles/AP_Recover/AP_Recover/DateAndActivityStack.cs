// Decompiled with JetBrains decompiler
// Type: AccelerometerProcessing_FiveSecond.DateAndActivityStack
// Assembly: AccelerometerProcessing_FiveSecond, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: D74A3EC3-23D5-487A-8102-93D5769E4944
// Assembly location: C:\Users\l\AppData\Local\Apps\2.0\5KCOR2G4.TRZ\N5LHNGBE.QAQ\acce..tion_c04be0470dfe34fc_0001.0000_388283ee07e1429b\AccelerometerProcessing_FiveSecond.exe

namespace AccelerometerProcessing_FiveSecond
{
  internal class DateAndActivityStack
  {
    public string Date { get; set; }

    public string Time { get; set; }

    public string StudentNumber { get; set; }

    public string TeacherFolderName { get; set; }

    public string ActivityNumber { get; set; }

    public int AxisRegSumE { get; set; }

    public int AxisRegSumF { get; set; }

    public int AxisRegSumG { get; set; }

    public float AxisRegSumH { get; set; }

    public DateAndActivityStack(string one,string oneHalf, string two, string three, string four, int five, int six, int seven, float eight)
    {
      this.Date = one;
      this.Time = oneHalf;
      this.StudentNumber = two;
      this.TeacherFolderName = three;
      this.ActivityNumber = four;
      this.AxisRegSumE = five;
      this.AxisRegSumF = six;
      this.AxisRegSumG = seven;
      this.AxisRegSumH = eight;
    }
  }
}
