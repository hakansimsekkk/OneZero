using OfficeOpenXml;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using epplus_test.Model;

namespace epplus_test
{
    class Program
    {
        public static string file = "C:\\Users\\Hakan\\Downloads\\TEST-01\\EKSTRE-GIRDI.xlsx";  //data dosyasının konumu
        public static string file2 = "C:\\Users\\Hakan\\Downloads\\TEST-01\\ekstre_rapor.xlsx";  //oluşturulan dosyanın konumu ve adı
        public static List<EkstreGirdiModel> ekstreGirdi = new List<EkstreGirdiModel>();
        public static List<FormulModel> formulModel = new List<FormulModel>();

        static void Main(string[] args)
        {
            GetEkstreGirdi();
            new Program().ExportRapor(file2);
        }

        public static void GetEkstreGirdi()
        {            
            FileInfo existingFile = new FileInfo(file);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                ekstreGirdi = new Program().GetList<EkstreGirdiModel>(worksheet).Take(10).ToList();
                //".Take(10).ToList()" yazarak ilk 10 satır için işlem yapıldı, kaldırılırsa 107000 veri kullanılır
            }
        }

        public void ExportRapor(string filePath)
        {
            var list = new List<RaporModelV1>();

            foreach (var item in ekstreGirdi)
            {
                RaporModelV1 model = new RaporModelV1();
                FormulModel formul = null;
                double ucret = 0;


                if (item.MESAFE.Contains("KISA") || item.MESAFE.Contains("ŞEHİRİÇİ") || item.MESAFE.Contains("YAKIN")) //isimleri karşılaştırdım
                {

                    if (item.KG_DESI < 6)//6 dan küçükse yapılacak işlem
                        ucret = 7;
                    else if (item.KG_DESI < 11)//11 dan küçükse yapılacak işlem
                        ucret = 9;
                    else if (item.KG_DESI < 16)//16 dan küçükse yapılacak işlem
                        ucret = 13;
                    else if (item.KG_DESI < 21)//21 dan küçükse yapılacak işlem
                        ucret = 15;
                    else if (item.KG_DESI < 31)//31 dan küçükse yapılacak işlem
                        ucret = 21;
                    else
                        ucret = 21 + ((item.KG_DESI - 30) * 0.7); //30 dan fazlaysa kullanılacak formül 

                    list.Add(
                        new RaporModelV1() // yeni liste oluşturuldu
                        {
                            KG_DESI = item.KG_DESI,
                            ADET = item.ADET,
                            MESAFE = item.MESAFE,
                            SIRA_NO = item.SIRA_NO,
                            Ucret = ucret
                        });
                }
                else // ilk isimler uymadıysa kullanılacak işlemler
                {

                    if (item.KG_DESI < 6)
                        ucret = 7.75;
                    else if (item.KG_DESI < 11)
                        ucret = 10;
                    else if (item.KG_DESI < 16)
                        ucret = 14.5;
                    else if (item.KG_DESI < 21)
                        ucret = 16.5;
                    else if (item.KG_DESI < 31)
                        ucret = 23.5;
                    else
                        ucret = 23.5 + ((item.KG_DESI - 30) * 0.78);

                    list.Add(// ve yeni liste oluşturuldu
                        new RaporModelV1()
                        {
                            KG_DESI = item.KG_DESI,
                            ADET = item.ADET,
                            MESAFE = item.MESAFE,
                            SIRA_NO = item.SIRA_NO,
                            Ucret = ucret
                        });
                }
            }

            List<RaporModel> pivot = new List<RaporModel>() //oluşturulan girdilerin sayılması
            {
                new RaporModel{MESAFE ="KISA", KargoAdedi = ekstreGirdi.Where(x => x.MESAFE == "KISA").Count() },
                new RaporModel{MESAFE ="ORTA", KargoAdedi = ekstreGirdi.Where(x => x.MESAFE == "ORTA").Count() },
                new RaporModel{MESAFE ="ŞEHİRİÇİ", KargoAdedi = ekstreGirdi.Where(x => x.MESAFE == "ŞEHİRİÇİ").Count() },
                new RaporModel{MESAFE ="UZAK", KargoAdedi = ekstreGirdi.Where(x => x.MESAFE == "UZAK").Count() },
                new RaporModel{MESAFE ="YAKIN", KargoAdedi = ekstreGirdi.Where(x => x.MESAFE == "YAKIN").Count() },
                new RaporModel{MESAFE ="Grand Total", KargoAdedi = ekstreGirdi.Count() },
            };

            using (ExcelPackage package = new ExcelPackage()) // yeni listelerin kaydedilmesi ve hangi satır sütundan başlayacağı
            {
                package.Workbook.Worksheets.Add("İşlem Sonucu").Cells[1, 1].LoadFromCollection(list, true);
                package.Workbook.Worksheets.Add("Pivot Rapor").Cells[2, 2].LoadFromCollection(pivot, true);
                package.SaveAs(new FileInfo(filePath));
            }
        }

        private List<T> GetList<T>(ExcelWorksheet sheet) //excel'lerin kullanıldığı kısım
        {
            List<T> list = new List<T>();
            var columnInfo = Enumerable.Range(1, sheet.Dimension.Columns).ToList().Select(x =>
                    new { Index = x, ColumnName =sheet.Cells[1,x].Value.ToString() }
            );

            for (int row = 2;  row < sheet.Dimension.Rows; row++) //okumaların ve verilerin çekileceği döngü 
            {
                T obj = (T)Activator.CreateInstance(typeof(T));
                foreach (var prop in typeof(T).GetProperties() )
                {
                    int col = columnInfo.SingleOrDefault(c => c.ColumnName == prop.Name).Index;
                    var val = sheet.Cells[row, col].Value;
                    var propType = prop.PropertyType;
                    prop.SetValue(obj, Convert.ChangeType(val,propType));
                }
                list.Add(obj);
            }

            return list;
        }
    }
}
