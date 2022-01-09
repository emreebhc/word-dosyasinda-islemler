using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmreBahce
{
    class Program
    {
        static string[] dosya;                

        static void Main(string[] args)
        {
            // Ana ekrandan yapılacak işlemi tutması için değişken tanımlaması
            string islem = ""; 
            
            do
            {
                //Uygulamanın işlem seçtirmek için ana ekranı 
                Console.WriteLine("");
                Console.WriteLine("****İŞLEMLER****");
                Console.WriteLine("1-Dosya Oku");                
                Console.WriteLine("2-Kelime Sayısı Hesapla");
                Console.WriteLine("3-Karakter Sayısı Hesapla");
                Console.WriteLine("4-Boşluk Karakter Sayısı Hesapla");
                Console.WriteLine("5-, den Sonra Dene Yaz");
                Console.WriteLine("6-Paragraf Sayısı Hesapla");
                Console.WriteLine("7-Sesli Harf Sayısı Hesapla");
                Console.WriteLine("8-Satır Değiştirme");
                Console.WriteLine("9-Sembolleri Getir");
                Console.WriteLine("10-Kıta Harf Sayısı Hesapla");
                Console.WriteLine("11-Çıkış");
                Console.Write("Yapmak İstediğiniz İşlemi Seçiniz:");

                islem = Console.ReadLine();

                // İşlemi okuduktan sonra seçilen işleme göre koşul sağlandığında yapılması gereken işlemler

                if (islem == "1")
                {                    
                    Console.Write("Dosya adresini giriniz:");
                    string dosya_adres = Console.ReadLine();

                    dosya = DosyaOku(dosya_adres);
                }
                else if (islem == "2")
                {
                    int kelime_sayisi = KelimeSayisiGetir();
                    Console.WriteLine("Dosyadaki toplam kelime sayısı:" + kelime_sayisi.ToString());
                }
                else if (islem == "3")
                {
                    Console.Write("Hesaplanmasını istediğini karakteri giriniz:");
                    string karakter = Console.ReadLine();

                    int karakter_sayisi = KarakterSayisiGetir(karakter);
                    Console.WriteLine("Dosyadaki toplam "+karakter+" sayısı:" + karakter_sayisi.ToString());
                }
                else if (islem == "4")
                {                    
                    int bosluk_karakter_sayisi = BoslukKarakterSayisiGetir();
                    Console.WriteLine("Dosyadaki toplam boşluk karakter sayısı:" + bosluk_karakter_sayisi.ToString());
                }
                else if (islem == "5")
                {
                    string metin = DeneYazGetir();
                    Console.WriteLine("Değiştirilmiş Metin:" + metin);
                }
                else if (islem == "6")
                {
                    int paragraf_sayisi = ParagrafSayisiGetir();
                    Console.WriteLine("Dosyadaki paragraf sayısı:" + paragraf_sayisi.ToString());
                }
                else if (islem == "7")
                {
                    int[] satir_sesli_harf_sayisi = SesliHarfSayisiGetir();
                    for (int i = 0; i < satir_sesli_harf_sayisi.Length; i++)
                    {
                        Console.WriteLine((i+1).ToString()+". satır sesli harf sayısı:" + satir_sesli_harf_sayisi[i].ToString());
                    }                    
                }
                else if (islem == "8")
                {
                    Console.Write("Değiştirilecek ilk satırı giriniz:");
                    int satir1 = Convert.ToInt32(Console.ReadLine());
                    Console.Write("Değiştirilecek ikinci satırı giriniz:");
                    int satir2 = Convert.ToInt32(Console.ReadLine());

                    string metin = SatirDegistirGetir(satir1, satir2);
                    Console.WriteLine("Değiştirilmiş Metin:" + metin);
                }
                else if (islem == "9")
                {                    
                    string semboller = SembolGetir();
                    Console.WriteLine("Semboller:" + semboller);
                }
                else if (islem == "10")
                {                   
                }
                else if (islem != "11")
                    Console.WriteLine("Yanlış seçim yaptınız. Lütfen tekrar deneyiniz!!!");

                Console.Write("Devam etmek için bir tuşa basınız.");
                Console.ReadLine();
            } while (islem != "11");
        }               

        
        public static string[] DosyaOku(string pdosya_adres)
        {                                                            
            Application application = new Application();
            Document document = application.Documents.Open(pdosya_adres);

                                                                     //Dosyadaki bütün kelimeleri words dizisine atama işlemi
            string [] words = new string[document.Words.Count];
                                             
            for (int i = 1; i <= document.Words.Count; i++)
            {
                                                                    //dizinin 0. indeksi boş kalmasın diye -1 kullandık 
                words[i-1] = document.Words[i].Text;                
            }
            
            application.Quit();
            
            Console.WriteLine("Dosya başarıyla okundu.");

            return words;
        }
        public static int KelimeSayisiGetir()
        {
            int kelime_sayisi = 0;
                                                         //Dosya dizisinin uzunluğu kadar döngü oluştur
            for (int i = 0; i < dosya.Length; i++)
            {
                if (!(dosya[i] == "." || dosya[i] == "," || dosya[i] == ";" || dosya[i] == "\r" || dosya[i] == "\v" || dosya[i] == "!" || dosya[i] == "..." || dosya[i] == "?" || dosya[i] == "'" || dosya[i] == ":" || dosya[i] == "."))
                {
                    kelime_sayisi++;
                }
            }

            return kelime_sayisi;
        }
        public static int KarakterSayisiGetir(string pkarakter_sayisi)
        {
            int karakter_sayisi = 0;
                                                                     //Dosya dizisi uzunluğu kadar döngü oluştur
            for (int i = 0; i < dosya.Length; i++)
            {
                                                                     // Eğer dosya dizisinde içeriyorsa karakter sayısını arttır
                if (dosya[i].Contains(pkarakter_sayisi))
                {
                    karakter_sayisi++;
                }
            }

            return karakter_sayisi;
        }
        public static int BoslukKarakterSayisiGetir()
        {
            int karakter_sayisi = 0;
                                                            // Dosya dizisinin uzunluğu kadar döngü oluştur
            for (int i = 0; i < dosya.Length; i++)
            {
                                                             //Eğer Dosya dizisinde boşluk karakteri içeriyorsa sayacı arttır
                if (dosya[i].Contains(" "))
                {
                    karakter_sayisi++;
                }
            }

            return karakter_sayisi;
        }
        public static string DeneYazGetir()
        {            
            string metin = "";

            for (int i = 0; i < dosya.Length; i++)
            {  
                                                                        //Eğer dosya dizisinde , karakteri içeriyorsa 
                if (dosya[i].Contains(","))  
                                                                        // replace değiştirme anlamında olan , ile , dene yi değiştir 
                    metin = metin + dosya[i].Replace(",", ", dene");                
                else                
                    metin = metin + dosya[i];                
            }

            return metin;
        }
        public static int ParagrafSayisiGetir()
        {
            int karakter_sayisi = 0;
                                                            // Dosya dizisinin uzunluğu kadar döngü oluştur
            for (int i = 0; i < dosya.Length; i++)
            {                                               // dosya dizisinde \r yani paragraf içeriyorsa sayacı bir arttır         
                if (dosya[i].Contains("\r"))
                {
                    karakter_sayisi++;
                }
            }

            return karakter_sayisi;
        }
        private static int[] SesliHarfSayisiGetir()
        {
            int satir_sayisi = 0;
                        
            for (int i = 0; i < dosya.Length; i++)
            {
                if (dosya[i] =="\v" || dosya[i] == "\r")                
                    satir_sayisi++;                
            }

            int[] sesli_harf_sayilari = new int[satir_sayisi + 1];
            int satir_sayac = 0;
            int satir_harf_sayisi = 0;

            for (int i = 0; i < dosya.Length; i++)
            {
                if (dosya[i] =="\v" || dosya[i] == "\r")
                {
                    satir_sayac++;
                    satir_harf_sayisi = 0;
                }
                else 
                {
                    for (int j = 0; j < dosya[i].Length; j++)
                    {
                        if (dosya[i][j] == 'a' || dosya[i][j] == 'A' || dosya[i][j] == 'e' || dosya[i][j] == 'E' || dosya[i][j] == 'ı' || dosya[i][j] == 'I' || dosya[i][j] == 'i' || dosya[i][j] == 'İ' || dosya[i][j] == 'o' || dosya[i][j] == 'O' || dosya[i][j] == 'ö' || dosya[i][j] == 'Ö' || dosya[i][j] == 'u' || dosya[i][j] == 'U' || dosya[i][j] == 'ü' || dosya[i][j] == 'Ü')
                        {
                            satir_harf_sayisi++;
                            sesli_harf_sayilari[satir_sayac] = satir_harf_sayisi;
                        }
                    }                                            
                }
            }

            return sesli_harf_sayilari;
        }
        private static string SatirDegistirGetir(int satir1, int satir2)
        {
            string metin = "";
            int satir_sayisi = 0;
            
            //word dizimizdeki satırların sayısını yakalamak için
            for (int i = 0; i < dosya.Length; i++)
            {
                if (dosya[i] == "\v" || dosya[i] == "\r")
                    satir_sayisi++;                
            }

            // yakaladığımı satır sayısı limitinde dizi oluşturmak için.
            string[] satirlar = new string[satir_sayisi + 1];
            satir_sayisi = 0;
            
            //yakaladığımız satırları yeni satırlar dizimize aktarıyoruz. 
                                                                            //dosya[0] korkma dosya[1] , dosya[2] sönmez .....
            

            for (int i = 0; i < dosya.Length; i++)
            {
                if (dosya[i] == "\v" || dosya[i] == "\r")
                {
                    satir_sayisi++;
                }
                else
                {
                    satirlar[satir_sayisi] = satirlar[satir_sayisi] + dosya[i];
                }                        
            }

            //swop yer değiştirme işlemi satırları 
            string satir = satirlar[satir1-1];
            satirlar[satir1-1] = satirlar[satir2-1];
            satirlar[satir2-1] = satir;

            for (int i = 0; i < satirlar.Length; i++)
            {
                metin = metin + satirlar[i];
            }

            return metin;
        }
        private static string SembolGetir()
        {
            string semboller = "";                                                

            for (int i = 0; i < dosya.Length; i++)
            {
                string kelime = dosya[i];

                for (int j = 0; j < kelime.Length; j++)
                {
                    //Eğer harf değilse yani sembol ise ve kelimeleri dahil etmiyoruz 
                    if (!Char.IsLetter(kelime[j]) && !semboller.Contains(kelime[j]))
                    {
                        semboller = semboller + kelime[j];
                    }                        
                }
            }

            return semboller;
        }
    }
}
