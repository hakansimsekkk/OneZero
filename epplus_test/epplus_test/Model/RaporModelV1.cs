using System;
using System.Collections.Generic;
using System.Text;

namespace epplus_test.Model
{
    public class RaporModelV1
    {//Raporun işlem sonucu kısmını tanımladım
        public int SIRA_NO { get; set; }
        public int ADET { get; set; }
        public int KG_DESI { get; set; }
        public string MESAFE { get; set; }
        public double Ucret { get; set; }
    }
}
