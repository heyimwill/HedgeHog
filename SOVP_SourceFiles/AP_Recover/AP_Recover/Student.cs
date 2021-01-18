// Decompiled with JetBrains decompiler
// Type: AccelerometerProcessing_FiveSecond.Student
// Assembly: AccelerometerProcessing_FiveSecond, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: D74A3EC3-23D5-487A-8102-93D5769E4944
// Assembly location: C:\Users\lauren.vonklinggraef\AppData\Local\Apps\2.0\5KCOR2G4.TRZ\N5LHNGBE.QAQ\acce..tion_c04be0470dfe34fc_0001.0000_388283ee07e1429b\AccelerometerProcessing_FiveSecond.exe

using System.Collections.Generic;

namespace AccelerometerProcessing_FiveSecond
{
  internal class Student
  {
    public string StudentID { get; set; }

    public List<Row> StudentRows { get; set; }

    public Student()
    {
      this.StudentRows = new List<Row>();
    }
  }
}
