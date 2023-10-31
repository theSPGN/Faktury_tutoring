using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq.Expressions;
using System.Reflection.Metadata;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Forms;


// Autorstwo: Mateusz Zajda

// Licencja: CC

namespace Faktura_zadanie_tutoring_
{

    public partial class Form1 : Form

    {

        public class Produkt
        {
            public static int liczba_porzadkowa = 0;
            public int nr;
            public string nazwa;
            public string liczba_sztuk;
            public string cena_netto;
            public string wartosc_netto;
            public string stawka_vat;
            public string kwota_vat;
            public string wartosc_brutto;
            public string termin_platnosci;
            public string forma_platnosci;
            public List<string> lista_cech = new List<string>();
            public Produkt(string nazwa, string liczba_sztuk, string cena_netto, string wartosc_netto, string stawka_vat, string kwota_vat, string wartosc_brutto, string termin_platnosci, string forma_platnosci)
            {
                liczba_porzadkowa++;
                nr = liczba_porzadkowa;
                this.nazwa = nazwa;               
                this.liczba_sztuk = liczba_sztuk;               
                this.cena_netto = cena_netto;     
                this.wartosc_netto = wartosc_netto;       
                this.stawka_vat = stawka_vat;        
                this.kwota_vat = kwota_vat;         
                this.wartosc_brutto = wartosc_brutto; 
                this.termin_platnosci = termin_platnosci;
                this.forma_platnosci = forma_platnosci;
                lista_cech.Add(this.nazwa);
                lista_cech.Add(this.liczba_sztuk);
                lista_cech.Add(this.cena_netto);
                lista_cech.Add(this.wartosc_netto);
                lista_cech.Add(this.stawka_vat);
                lista_cech.Add(this.kwota_vat);
                lista_cech.Add(this.wartosc_brutto);
                lista_cech.Add(this.termin_platnosci);
                lista_cech.Add(this.forma_platnosci);
            }
            
            
        }
        public class Faktura
        {
            public string SpNazwa;
            public string SpNip;
            public string SpAdres;
            public string SpKod;
            public string SpMiasto;
            public string NaNazwa;
            public string NaNip;
            public string NaAdres;
            public string NaKod;
            public string NaMiasto;
            public List<Produkt> prod1;
            public Faktura(string spNazwa, string spNip, string spAdres, string spKod, string spMiasto, string naNazwa, string naNip, string naAdres, string naKod, string naMiasto, List<Produkt> prod1)
            {
                SpNazwa = spNazwa;
                SpNip = spNip;
                SpAdres = spAdres;
                SpKod = spKod;
                SpMiasto = spMiasto;
                NaNazwa = naNazwa;
                NaNip = naNip;
                NaAdres = naAdres;
                NaKod = naKod;
                NaMiasto = naMiasto;
                this.prod1 = prod1;
            }
        }

        public int numer_faktury = 0; 
        public List<Produkt> lista_produktow = new List<Produkt>();
        public int pozycja_ostatniego = 0;
        static int ostatni_element_index = 0;
        static string path = "my_file.txt"; //bin\debug'
        public Form1()
        {
            InitializeComponent();
            
            try { Int32.TryParse(File.ReadLines(path).Last(), out numer_faktury); }
            catch(Exception e) { Console.WriteLine(e); }
            finally
            {
                numer_faktury++;// ostatnia faktura miala numer np. 10, wiêc ta ma 11
                                // jeœli nie by³o faktury to tworzy nowy plik z numerem 1
            }
            
            
        }

        private void Drukuj_Click(object sender, EventArgs e)
        {

            P1dialog.Document = P1doc;
            P1dialog.AllowSomePages = false;
            P1dialog.AllowSelection = false;
            if (P1dialog.ShowDialog() == DialogResult.OK) {
                P1doc.PrintPage += new PrintPageEventHandler
                  (this.pd_PrintPage);
                using (StreamWriter sw = new StreamWriter(path, false))
                    sw.Write(numer_faktury.ToString());
                P1doc.Print();
            }
               
            
        }

        private void zapisz_dane_do_pliku(object Faktura)
        {
            // Wpisywanie do pliku danych (pod bazy danych?)
        }

        private void nazwa_faktury()
        {
            DateTime localDate = DateTime.Today;
            string Faktura_data = localDate.ToString();
            string Faktura_name = Faktura_data.Substring(0, 10);
            Faktura_name = Faktura_name.Replace('.', '\0');
            string numer_faktury_word = "00000000";
            numer_faktury_word = numer_faktury_word.Substring(0, numer_faktury_word.Length - numer_faktury.ToString().Length);
            numer_faktury_word += numer_faktury.ToString();
            P1doc.DocumentName = "Faktura: " + numer_faktury_word + Faktura_name;
        }
        private void naglowki(object sender, PrintPageEventArgs ev, Faktura faktura_do_druku)
        {
            int wielkosc_tekstu_danych = 12;
            int wielkosc_naglowka = 15;
            Font printFontTitles = new Font("Arial", wielkosc_naglowka, FontStyle.Bold);
            Font printFont = new Font("Arial", wielkosc_tekstu_danych);

            int pozycja_sprzedawca_lewo = 50;
            int pozycja_sprzedawca_gora = 50;
            int pozycja_sprzedawca_tekst = 0;
            ev.Graphics.DrawString("Dane Sprzedawcy:", printFontTitles, Brushes.Black, new PointF(pozycja_sprzedawca_lewo, pozycja_sprzedawca_gora)); //800 x 1150 max
            pozycja_sprzedawca_tekst += pozycja_sprzedawca_gora + wielkosc_naglowka + wielkosc_tekstu_danych;
            ev.Graphics.DrawString("Nazwa: " + faktura_do_druku.SpNazwa, printFont, Brushes.Black, new PointF(pozycja_sprzedawca_lewo, pozycja_sprzedawca_tekst));
            pozycja_sprzedawca_tekst += wielkosc_tekstu_danych * 2;
            ev.Graphics.DrawString("NIP: " + faktura_do_druku.SpNip, printFont, Brushes.Black, new PointF(pozycja_sprzedawca_lewo, pozycja_sprzedawca_tekst));
            pozycja_sprzedawca_tekst += wielkosc_tekstu_danych * 2;
            ev.Graphics.DrawString("Adres: " + faktura_do_druku.SpAdres, printFont, Brushes.Black, new PointF(pozycja_sprzedawca_lewo, pozycja_sprzedawca_tekst));
            pozycja_sprzedawca_tekst += wielkosc_tekstu_danych * 2;
            ev.Graphics.DrawString("Kod pocztowy: " + faktura_do_druku.SpKod + " " + faktura_do_druku.SpMiasto, printFont, Brushes.Black, new PointF(pozycja_sprzedawca_lewo, pozycja_sprzedawca_tekst));
            pozycja_sprzedawca_tekst += wielkosc_tekstu_danych * 2 + 2 * wielkosc_naglowka;
            ev.Graphics.DrawString("Data: " + DateTime.Today.ToString().Substring(0,10), printFont, Brushes.Black, new PointF(pozycja_sprzedawca_lewo, pozycja_sprzedawca_tekst));
            pozycja_sprzedawca_tekst += wielkosc_tekstu_danych * 2;
            ev.Graphics.DrawString(P1doc.DocumentName, printFont, Brushes.Black, new PointF(pozycja_sprzedawca_lewo, pozycja_sprzedawca_tekst));

            pozycja_ostatniego = pozycja_sprzedawca_tekst;

            int pozycja_nabywca_lewo = 450;
            int pozycja_nabywca_gora = 50;
            int pozycja_nabywca_tekst = 0;
            ev.Graphics.DrawString("Dane Nabywcy:", printFontTitles, Brushes.Black, new PointF(pozycja_nabywca_lewo, pozycja_nabywca_gora));
            pozycja_nabywca_tekst += pozycja_nabywca_gora + wielkosc_naglowka + wielkosc_tekstu_danych;
            ev.Graphics.DrawString("Nazwa: " + faktura_do_druku.NaNazwa, printFont, Brushes.Black, new PointF(pozycja_nabywca_lewo, pozycja_nabywca_tekst));
            pozycja_nabywca_tekst += wielkosc_tekstu_danych * 2;
            ev.Graphics.DrawString("NIP: " + faktura_do_druku.NaNip, printFont, Brushes.Black, new PointF(pozycja_nabywca_lewo, pozycja_nabywca_tekst));
            pozycja_nabywca_tekst += wielkosc_tekstu_danych * 2;
            ev.Graphics.DrawString("Adres :" + faktura_do_druku.NaAdres, printFont, Brushes.Black, new PointF(pozycja_nabywca_lewo, pozycja_nabywca_tekst));
            pozycja_nabywca_tekst += wielkosc_tekstu_danych * 2;
            ev.Graphics.DrawString("Kod pocztowy: " + faktura_do_druku.NaKod + " " + faktura_do_druku.NaMiasto, printFont, Brushes.Black, new PointF(pozycja_nabywca_lewo, pozycja_nabywca_tekst));
            
        }
        
        private void wypisz_produkty(object sender, PrintPageEventArgs ev, List<Produkt> lista_produktow, int pozycja_ostatniego)
        {
            int wielkosc_tekstu_danych = 12;
            int wielkosc_naglowka = 15;
            int wysokosc = pozycja_ostatniego + wielkosc_naglowka * 2 + 25;
            //wysokosc = 840; //do testów skoñczenia strony

            Font printFontTitles = new Font("Arial", wielkosc_naglowka, FontStyle.Bold);
            Font printFont = new Font("Arial", wielkosc_tekstu_danych);
            Font opisFont = new Font("Arial", wielkosc_tekstu_danych - 2);

            Func<string, Font, float> dlugosc = (mytext,myfont) => ev.Graphics.MeasureString(mytext, myfont, 0, StringFormat.GenericTypographic).Width;
            Func<string, Font, float> pozycja_srodka = (mytext,myfont) => (ev.PageSettings.PaperSize.Width / 2) - (dlugosc(mytext,myfont) / 2);

            ev.Graphics.DrawString("Produkty:", printFontTitles, Brushes.Black,new PointF(pozycja_srodka("Produkty:", printFontTitles), wysokosc));
            
            wysokosc += 2 * wielkosc_tekstu_danych;
            
            string opis = "Lp./ Nazwa/ liczba sztuk/ cena netto/ wartosc netto/ stawka vat/ kwota vat/ wartosc brutto/ termin platnosci/ forma platnosci";
            ev.Graphics.DrawString(opis, opisFont, Brushes.Black, new PointF(pozycja_srodka(opis, opisFont), wysokosc));

            List<Produkt> sublist = lista_produktow.GetRange(ostatni_element_index, lista_produktow.Count - ostatni_element_index);
            foreach (Produkt elementlisty in sublist)
            {
                wysokosc += 2 * wielkosc_tekstu_danych;
                int margines = 30;
                int odstep = 20;
                int pozycja_liczby_porzadkowej = margines;
                int pozycja_lewo_nazwa = margines + odstep*2;
                int pozycja_lewo_liczba_sztuk = pozycja_lewo_nazwa + (int) dlugosc(elementlisty.nazwa, printFont) + odstep;
                int pozycja_lewo_cena_netto = pozycja_lewo_liczba_sztuk + (int)dlugosc(elementlisty.liczba_sztuk, printFont) + odstep; 
                int pozycja_lewo_wartosc_netto = pozycja_lewo_cena_netto + (int)dlugosc(elementlisty.cena_netto, printFont) + odstep;
                int pozycja_lewo_stawka_vat = pozycja_lewo_wartosc_netto + (int)dlugosc(elementlisty.wartosc_netto, printFont) + odstep; 
                int pozycja_lewo_kwota_vat = pozycja_lewo_stawka_vat + (int)dlugosc(elementlisty.stawka_vat, printFont) + odstep;    
                int pozycja_lewo_wartosc_brutto = pozycja_lewo_kwota_vat + (int)dlugosc(elementlisty.kwota_vat, printFont) + odstep; 
                int pozycja_lewo_termin_platnosci = pozycja_lewo_wartosc_brutto + (int)dlugosc(elementlisty.wartosc_brutto, printFont) + odstep; 
                int pozycja_lewo_forma_platnosci = pozycja_lewo_termin_platnosci + (int)dlugosc(elementlisty.termin_platnosci, printFont) + odstep;

                int [] pozycja = {pozycja_lewo_nazwa,pozycja_lewo_liczba_sztuk,pozycja_lewo_cena_netto,
                    pozycja_lewo_wartosc_netto, pozycja_lewo_stawka_vat, pozycja_lewo_kwota_vat, 
                    pozycja_lewo_wartosc_brutto, pozycja_lewo_termin_platnosci,pozycja_lewo_forma_platnosci
                };
                
                
                
                if (wysokosc >= ev.PageSettings.PaperSize.Height - wielkosc_tekstu_danych * 20)
                {
                    ev.HasMorePages = true;
                    ostatni_element_index += sublist.IndexOf(elementlisty);
                    return;
                }
                else { ev.HasMorePages = false;
                    ostatni_element_index = 0;
                }



                ev.Graphics.DrawString((elementlisty.nr).ToString() + ".", new Font(printFont,FontStyle.Bold), Brushes.Black, new PointF(pozycja_liczby_porzadkowej, wysokosc));
                int index = 0;
                foreach (string parametr in elementlisty.lista_cech)
                {
                    int ilezmiennych = elementlisty.lista_cech.Count;
                    string parametr_to_draw = parametr;
                    if (index < ilezmiennych - 1)
                        parametr_to_draw += " / ";

                    int polozenie_max = ev.PageSettings.PaperSize.Width - margines - (int)dlugosc(parametr, printFont);

                    if (pozycja[index] >= polozenie_max)
                    {
                        int new_line = pozycja[index];
                        for (int a = 0; a < pozycja.Length; a++)
                        {
                            pozycja[a] = pozycja[a] - new_line + margines;
                        }

                        wysokosc += 2 * wielkosc_tekstu_danych;
                    }

                    ev.Graphics.DrawString(parametr_to_draw, printFont, Brushes.Black, new PointF(pozycja[index], wysokosc));
                    index++;
                }

            }
           
        }

        private void podsumowanie(object sender, PrintPageEventArgs ev, List<Produkt> lista_produktow)
        {
            int wielkosc_tekstu_danych = 12;
            Font printFont = new Font("Arial", wielkosc_tekstu_danych);
            int wysokosc = ev.PageSettings.PaperSize.Height;
            string Tekst_do_wypisania = "Podsumowanie: \n";
            float suma_cena_netto = 0;
            float suma_wartosc_netto = 0;
            float suma_kwota_vat = 0;
            float suma_wartosc_brutto = 0;
            foreach(var produkt in lista_produktow)
            {
                
                float.TryParse(produkt.cena_netto, out float cena_n);
                
                float.TryParse(produkt.wartosc_netto, out float wartosc_n);
               
                float.TryParse(produkt.kwota_vat, out float kwota_v);
             
                float.TryParse(produkt.wartosc_brutto, out float wartosc_b);


                suma_cena_netto += cena_n;
                suma_wartosc_netto += wartosc_n;
                suma_kwota_vat += kwota_v;
                suma_wartosc_brutto += wartosc_b;
            }
            Tekst_do_wypisania += "Suma cen netto: " + suma_cena_netto.ToString() + "\n";
            Tekst_do_wypisania += "Suma wartosci netto: " + suma_wartosc_netto.ToString() + "\n";
            Tekst_do_wypisania += "Suma kwot netto: " + suma_kwota_vat.ToString() + "\n";
            Tekst_do_wypisania += "Suma wartosci brutto: " + suma_wartosc_brutto.ToString();

            ev.Graphics.DrawString(Tekst_do_wypisania, printFont, Brushes.Black, new PointF(50, wysokosc - wielkosc_tekstu_danych * 18));
            


        }


        private void miejsce_na_podpis(object sender, PrintPageEventArgs ev)
        {
            int wielkosc_tekstu_danych = 15;
            Font printFont = new Font("Arial", wielkosc_tekstu_danych);
            int wysokosc = ev.PageSettings.PaperSize.Height;
            ev.Graphics.DrawString("Podpis sprzedawcy:", printFont, Brushes.Black, new PointF(50, wysokosc - wielkosc_tekstu_danych*7));
            ev.Graphics.DrawString("Podpis Nabywcy:", printFont, Brushes.Black, new PointF(600, wysokosc - wielkosc_tekstu_danych*7));
            ev.Graphics.DrawString("............................", printFont, Brushes.Black, new PointF(50, wysokosc - wielkosc_tekstu_danych*4));
            ev.Graphics.DrawString("............................", printFont, Brushes.Black, new PointF(600, wysokosc - wielkosc_tekstu_danych*4));
            
        }

        private void pd_PrintPage(object sender, PrintPageEventArgs ev)
        {
            Faktura faktura_do_druku = new(SprzedawcaNazwaFirmy.Text,
                SprzedawcaNIP.Text, SprzedawcaAdres.Text, SprzedawcaKodPocztowy.Text,
                SprzedawcaMiasto.Text, NabywcaNazwaFirmy.Text, NabywcaNIP.Text,
                NabywcaAdres.Text, NabywcaKodPocztowy.Text, NabywcaMiasto.Text, lista_produktow);

            zapisz_dane_do_pliku(faktura_do_druku);
            nazwa_faktury();
            naglowki(sender, ev, faktura_do_druku);
            wypisz_produkty(sender, ev, lista_produktow, pozycja_ostatniego);
            miejsce_na_podpis(sender, ev);
            podsumowanie(sender, ev, lista_produktow);

        }

        private void DodajPozycje_Click(object sender, EventArgs e)
        {
            if (NazwaProduktu.Text != "" && LiczbaSztuk.Text != "" && CenaNetto.Text != "" && WartoscNetto.Text != "" && StawkaVAT.Text != "" && KwotaVAT.Text != "" && WartoscBrutto.Text != "" && TerminPlatnosci.Text != "" && FormaPlatnosci.Text != "")
            {
                Produkt p1 = new Produkt(NazwaProduktu.Text, LiczbaSztuk.Text,
                    CenaNetto.Text, WartoscNetto.Text, StawkaVAT.Text,
                    KwotaVAT.Text, WartoscBrutto.Text, TerminPlatnosci.Text, FormaPlatnosci.Text);
                lista_produktow.Add(p1);
                NazwaProduktu.Text = "";
                LiczbaSztuk.Text = "";
                CenaNetto.Text = "";
                WartoscNetto.Text = "";
                StawkaVAT.Text = "";
                KwotaVAT.Text = "";
                WartoscBrutto.Text = "";
                TerminPlatnosci.Text = "";
                FormaPlatnosci.Text = "";
            }
            else
            {
                MessageBox.Show("Podaj wszystkie dane");
            }
        }

        private void Preview_Click(object sender, EventArgs e)
        {
            P1doc.PrintPage += new PrintPageEventHandler
                   (this.pd_PrintPage);
            if (P1preview.ShowDialog() == DialogResult.OK)
            {
                P1doc.Print();
                using (StreamWriter sw = new StreamWriter(path, false))
                    sw.Write(numer_faktury.ToString());
            }
                

        }
    }
}