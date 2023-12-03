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

using System.Data;
using System.Collections.ObjectModel;
using System.Security.Policy;
using System.Diagnostics.Tracing;
using Microsoft.Data.SqlClient;
using static Faktura_zadanie_tutoring_.Form1;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;
using System.CodeDom.Compiler;
using System.Text;
// Autorstwo: Mateusz Zajda

// Licencja: CC

// data gridview
// data table
namespace Faktura_zadanie_tutoring_
{

    public partial class Form1 : Form

    {

        public class Produkt
        {
            public static int liczba_porzadkowa = 0;
            public int nr { get; set; }
            public string nazwa { get; set; }
            public string liczba_sztuk { get; set; }
            public string cena_netto { get; set; }
            public string wartosc_netto { get; set; }
            public string stawka_vat { get; set; }
            public string kwota_vat { get; set; }
            public string wartosc_brutto { get; set; }
            public string termin_platnosci { get; set; }
            public string forma_platnosci { get; set; }
            public List<string> lista_cech { get; set; }

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
                lista_cech = new List<string>();
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
            catch (Exception e) { Console.WriteLine(e); }
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

            if (P1dialog.ShowDialog() == DialogResult.OK)
            {
                P1doc.PrintPage += new PrintPageEventHandler
                  (this.pd_PrintPage);
                using (StreamWriter sw = new StreamWriter(path, false))
                    sw.Write(numer_faktury.ToString());
                P1doc.Print();

                Faktura faktura_do_druku = new(SprzedawcaNazwaFirmy.Text,
                SprzedawcaNIP.Text, SprzedawcaAdres.Text, SprzedawcaKodPocztowy.Text,
                SprzedawcaMiasto.Text, NabywcaNazwaFirmy.Text, NabywcaNIP.Text,
                NabywcaAdres.Text, NabywcaKodPocztowy.Text, NabywcaMiasto.Text, lista_produktow);

                WriteToDatabase(lista_produktow);
            }


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
            P1doc.DocumentName = "Faktura: " + numer_faktury_word + "_" + Faktura_name;
            P1doc.PrinterSettings.PrintFileName = P1doc.DocumentName;
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
            ev.Graphics.DrawString("Data: " + DateTime.Today.ToString().Substring(0, 10), printFont, Brushes.Black, new PointF(pozycja_sprzedawca_lewo, pozycja_sprzedawca_tekst));
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

            Font printFontTitles = new Font("Arial", wielkosc_naglowka, FontStyle.Bold);
            Font printFont = new Font("Arial", wielkosc_tekstu_danych);
            Font opisFont = new Font("Arial", wielkosc_tekstu_danych - 2);

            Func<string, Font, float> dlugosc = (mytext, myfont) => ev.Graphics.MeasureString(mytext, myfont, 0, StringFormat.GenericTypographic).Width;
            Func<string, Font, float> pozycja_srodka = (mytext, myfont) => (ev.PageSettings.PaperSize.Width / 2) - (dlugosc(mytext, myfont) / 2);

            ev.Graphics.DrawString("Produkty:", printFontTitles, Brushes.Black, new PointF(pozycja_srodka("Produkty:", printFontTitles), wysokosc));

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
                int pozycja_lewo_nazwa = margines + odstep * 2;
                int pozycja_lewo_liczba_sztuk = pozycja_lewo_nazwa + (int)dlugosc(elementlisty.nazwa, printFont) + odstep;
                int pozycja_lewo_cena_netto = pozycja_lewo_liczba_sztuk + (int)dlugosc(elementlisty.liczba_sztuk, printFont) + odstep;
                int pozycja_lewo_wartosc_netto = pozycja_lewo_cena_netto + (int)dlugosc(elementlisty.cena_netto, printFont) + odstep;
                int pozycja_lewo_stawka_vat = pozycja_lewo_wartosc_netto + (int)dlugosc(elementlisty.wartosc_netto, printFont) + odstep;
                int pozycja_lewo_kwota_vat = pozycja_lewo_stawka_vat + (int)dlugosc(elementlisty.stawka_vat, printFont) + odstep;
                int pozycja_lewo_wartosc_brutto = pozycja_lewo_kwota_vat + (int)dlugosc(elementlisty.kwota_vat, printFont) + odstep;
                int pozycja_lewo_termin_platnosci = pozycja_lewo_wartosc_brutto + (int)dlugosc(elementlisty.wartosc_brutto, printFont) + odstep;
                int pozycja_lewo_forma_platnosci = pozycja_lewo_termin_platnosci + (int)dlugosc(elementlisty.termin_platnosci, printFont) + odstep;

                int[] pozycja = {pozycja_lewo_nazwa,pozycja_lewo_liczba_sztuk,pozycja_lewo_cena_netto,
                    pozycja_lewo_wartosc_netto, pozycja_lewo_stawka_vat, pozycja_lewo_kwota_vat,
                    pozycja_lewo_wartosc_brutto, pozycja_lewo_termin_platnosci,pozycja_lewo_forma_platnosci
                };



                if (wysokosc >= ev.PageSettings.PaperSize.Height - wielkosc_tekstu_danych * 20)
                {
                    ev.HasMorePages = true;
                    ostatni_element_index += sublist.IndexOf(elementlisty);
                    return;
                }
                else
                {
                    ev.HasMorePages = false;
                    ostatni_element_index = 0;
                }



                ev.Graphics.DrawString((elementlisty.nr).ToString() + ".", new Font(printFont, FontStyle.Bold), Brushes.Black, new PointF(pozycja_liczby_porzadkowej, wysokosc));
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
            foreach (var produkt in lista_produktow)
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

        private void Wypisz_produkty_w_tabeli(object sender, PrintPageEventArgs ev, List<Produkt> lista_produktow, int pozycja_ostatniego)
        {
            int wielkosc_tekstu_danych = 12;
            int wielkosc_naglowka = 15;
            int wysokosc = pozycja_ostatniego + wielkosc_naglowka * 2 + 25;
            Font printFontTitles = new Font("Arial", wielkosc_naglowka, FontStyle.Bold);
            Func<string, Font, float> dlugosc = (mytext, myfont) => ev.Graphics.MeasureString(mytext, myfont, 0, StringFormat.GenericTypographic).Width;
            Func<string, Font, float> pozycja_srodka = (mytext, myfont) => (ev.PageSettings.PaperSize.Width / 2) - (dlugosc(mytext, myfont) / 2);

            ev.Graphics.DrawString("Produkty:", printFontTitles, Brushes.Black, new PointF(pozycja_srodka("Produkty:", printFontTitles), wysokosc));

            var DataGridView1 = new DataGridView();
            DataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            DataGridView1.BackgroundColor = Color.White;

            List<Produkt> list_with_spaces = new List<Produkt>();
            int numer = 1;
            foreach (var Produkty in lista_produktow)
            {
                Produkt produkt_with_spaces = new Produkt(
                dodaj_spacje(Produkty.nazwa, 12),
                dodaj_spacje(Produkty.liczba_sztuk, 12),
                dodaj_spacje(Produkty.cena_netto, 12),
                dodaj_spacje(Produkty.wartosc_netto, 12),
                dodaj_spacje(Produkty.stawka_vat, 12),
                dodaj_spacje(Produkty.kwota_vat, 12),
                dodaj_spacje(Produkty.wartosc_brutto, 12),
                dodaj_spacje(Produkty.termin_platnosci, 12),
                dodaj_spacje(Produkty.forma_platnosci, 8)
                    );
                list_with_spaces.Add(produkt_with_spaces);
                produkt_with_spaces.nr = numer;
                numer++;
            }



            while (ostatni_element_index < list_with_spaces.Count)
            {
                List<Produkt> sublist = list_with_spaces.GetRange(ostatni_element_index, list_with_spaces.Count - ostatni_element_index);

                if (sublist.Count == 0)
                    break;

                DataGridView1.DataSource = sublist;
                DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                DataGridView1.RowHeadersVisible = false;
                this.Controls.Add(DataGridView1);
                this.PerformLayout();
                DataGridView1.Visible = false;
                DataGridView1.ScrollBars = ScrollBars.None;
                int summary_width_of_columns = 0;
                int summary_height_of_rows = 0;

                DataGridView1.Columns[0].Width = 25;

                for (int i = 1; i < list_with_spaces[0].lista_cech.Count + 1; i++)
                {
                    var column = DataGridView1.Columns[i];
                    column.Width = (800 - 100) / (list_with_spaces[0].lista_cech.Count - 1);
                    summary_width_of_columns += column.Width;
                }

                while (true)
                {
                    summary_height_of_rows = 0;
                    for (int i = 0; i < sublist.Count; i++)
                    {
                        summary_height_of_rows += DataGridView1.Rows[i].Height;
                    }
                    if (summary_height_of_rows > 550)
                    {
                        sublist = sublist.GetRange(0, sublist.Count - 1);
                        DataGridView1.DataSource = sublist;
                    }
                    else break;
                }

                DataGridView1.DataSource = sublist;
                DataGridView1.Width = summary_width_of_columns;
                DataGridView1.Height = summary_height_of_rows + 1 * DataGridView1.RowTemplate.Height;

                Bitmap bitmap = new Bitmap(DataGridView1.Width, DataGridView1.Height);
                DataGridView1.ClearSelection();
                DataGridView1.DrawToBitmap(bitmap, new Rectangle(0, 0, DataGridView1.Width, DataGridView1.Height));

                ev.Graphics.DrawImage(bitmap, (ev.PageSettings.PaperSize.Width - DataGridView1.Width - DataGridView1.RowHeadersWidth + 10) / 2, wysokosc + wielkosc_tekstu_danych * 2);

                ostatni_element_index += sublist.Count;

                if (ostatni_element_index < list_with_spaces.Count)
                    ev.HasMorePages = true;
                else
                {
                    ev.HasMorePages = false;
                    ostatni_element_index = 0;
                }
                return;
            }

        }

        private void miejsce_na_podpis(object sender, PrintPageEventArgs ev)
        {
            int wielkosc_tekstu_danych = 15;
            Font printFont = new Font("Arial", wielkosc_tekstu_danych);
            int wysokosc = ev.PageSettings.PaperSize.Height;
            ev.Graphics.DrawString("Podpis sprzedawcy:", printFont, Brushes.Black, new PointF(50, wysokosc - wielkosc_tekstu_danych * 7));
            ev.Graphics.DrawString("Podpis Nabywcy:", printFont, Brushes.Black, new PointF(600, wysokosc - wielkosc_tekstu_danych * 7));
            ev.Graphics.DrawString("............................", printFont, Brushes.Black, new PointF(50, wysokosc - wielkosc_tekstu_danych * 4));
            ev.Graphics.DrawString("............................", printFont, Brushes.Black, new PointF(600, wysokosc - wielkosc_tekstu_danych * 4));

        }

        private void pd_PrintPage(object sender, PrintPageEventArgs ev)
        {
            ev.Graphics.Clear(Color.White);
            Faktura faktura_do_druku = new(SprzedawcaNazwaFirmy.Text,
                SprzedawcaNIP.Text, SprzedawcaAdres.Text, SprzedawcaKodPocztowy.Text,
                SprzedawcaMiasto.Text, NabywcaNazwaFirmy.Text, NabywcaNIP.Text,
                NabywcaAdres.Text, NabywcaKodPocztowy.Text, NabywcaMiasto.Text, lista_produktow);

            ev.HasMorePages = false;
            ev.PageSettings.PrinterSettings.PrintFileName = P1doc.DocumentName;
            ev.Graphics.DrawString("", new Font("Arial", 1), Brushes.Black, 0, 0);
            nazwa_faktury();
            naglowki(sender, ev, faktura_do_druku);
            //wypisz_produkty(sender, ev, lista_produktow, pozycja_ostatniego);
            Wypisz_produkty_w_tabeli(sender, ev, lista_produktow, pozycja_ostatniego);
            miejsce_na_podpis(sender, ev);
            podsumowanie(sender, ev, lista_produktow);

        }

        private string dodaj_spacje(string str, int coile)
        {
            string[] array = str.Split(" ");
            StringBuilder new_str = new StringBuilder();
            for(int k = 0; k < array.Length; k++)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < array[k].Length; i++)
                {
                    if (i % coile == 0 && i != 0 && i != array[k].Length-1)
                    {
                        sb.Append("- ");
                    }
                    sb.Append(array[k][i]) ;
                }

                string wynik = sb.ToString();
                new_str.Append(wynik);
                new_str.Append(" ");
            }
            return new_str.ToString();

        }

        private void DodajPozycje_Click(object sender, EventArgs e)
        {
            if (NazwaProduktu.Text != "" && LiczbaSztuk.Text != "" && CenaNetto.Text != "" && WartoscNetto.Text != "" && StawkaVAT.Text != "" && KwotaVAT.Text != "" && WartoscBrutto.Text != "" && TerminPlatnosci.Text != "" && FormaPlatnosci.Text != "")
            {
                Produkt p1 = new Produkt(
                    NazwaProduktu.Text,
                    LiczbaSztuk.Text,
                    CenaNetto.Text,
                    WartoscNetto.Text,
                    StawkaVAT.Text,
                    KwotaVAT.Text,
                    WartoscBrutto.Text,
                    TerminPlatnosci.Text,
                    FormaPlatnosci.Text
                    );

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
                using (StreamWriter sw = new StreamWriter(path, false))
                    sw.Write(numer_faktury.ToString());
                WriteToDatabase(lista_produktow);
            }


        }

        public void WriteToDatabase(List<Produkt> lista_produktow)
        {
            string connectionString = "Server=DESKTOP-OOPT79L\\SQLEXPRESS;Database=FAKTURA;" +
                "Integrated Security=True;Encrypt=False;TrustServerCertificate=true;";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                foreach (var produkt in lista_produktow)
                {
                    string query = "INSERT INTO Produkty (numer_faktury, nr, nazwa, liczba_sztuk, cena_netto, wartosc_netto, stawka_vat, kwota_vat, wartosc_brutto, termin_platnosci, forma_platnosci) " +
                                   "VALUES (@numer_faktury, @nr, @nazwa, @liczba_sztuk, @cena_netto, @wartosc_netto, @stawka_vat, @kwota_vat, @wartosc_brutto, @termin_platnosci, @forma_platnosci)";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@numer_faktury", numer_faktury);
                        command.Parameters.AddWithValue("@nr", produkt.nr);
                        command.Parameters.AddWithValue("@nazwa", produkt.nazwa);
                        command.Parameters.AddWithValue("@liczba_sztuk", produkt.liczba_sztuk);
                        command.Parameters.AddWithValue("@cena_netto", produkt.cena_netto);
                        command.Parameters.AddWithValue("@wartosc_netto", produkt.wartosc_netto);
                        command.Parameters.AddWithValue("@stawka_vat", produkt.stawka_vat);
                        command.Parameters.AddWithValue("@kwota_vat", produkt.kwota_vat);
                        command.Parameters.AddWithValue("@wartosc_brutto", produkt.wartosc_brutto);
                        command.Parameters.AddWithValue("@termin_platnosci", produkt.termin_platnosci);
                        command.Parameters.AddWithValue("@forma_platnosci", produkt.forma_platnosci);

                        command.ExecuteNonQuery();
                    }
                }

                Faktura faktura_do_druku = new(SprzedawcaNazwaFirmy.Text,
                   SprzedawcaNIP.Text, SprzedawcaAdres.Text, SprzedawcaKodPocztowy.Text,
                   SprzedawcaMiasto.Text, NabywcaNazwaFirmy.Text, NabywcaNIP.Text,
                   NabywcaAdres.Text, NabywcaKodPocztowy.Text, NabywcaMiasto.Text, lista_produktow);
                string query2 = "INSERT INTO Klient (SprzedawcaNazwaFirmy, SprzedawcaNIP, " +
                    "SprzedawcaAdres, SprzedawcaKodPocztowy, SprzedawcaMiasto, NabywcaNazwaFirmy, NabywcaNip, NabywcaAdres" +
                    "NabywcaKodPocztowy, NabywcaMiasto) " +
                               "VALUES (@SprzedawcaNazwaFirmy, @SprzedawcaNIP, " +
                    "@SprzedawcaAdres, @SprzedawcaKodPocztowy, @SprzedawcaMiasto, @NabywcaNazwaFirmy, @NabywcaNip, " +
                    "@NabywcaAdres, @NabywcaKodPocztowy, @NabywcaMiasto)";
                using (SqlCommand command = new SqlCommand(query2, connection))
                {
                    command.Parameters.AddWithValue("@SprzedawcaNazwaFirmy", SprzedawcaNazwaFirmy.Text);
                    command.Parameters.AddWithValue("@SprzedawcaNIP", SprzedawcaNIP.Text);
                    command.Parameters.AddWithValue("@SprzedawcaAdres", SprzedawcaAdres.Text);
                    command.Parameters.AddWithValue("@SprzedawcaKodPocztowy", SprzedawcaKodPocztowy.Text);
                    command.Parameters.AddWithValue("@SprzedawcaMiasto", SprzedawcaMiasto.Text);
                    command.Parameters.AddWithValue("@NabywcaNazwaFirmy", NabywcaNazwaFirmy.Text);
                    command.Parameters.AddWithValue("@NabywcaNip", NabywcaNIP.Text);
                    command.Parameters.AddWithValue("@NabywcaAdres", NabywcaAdres.Text);
                    command.Parameters.AddWithValue("@NabywcaKodPocztowy", NabywcaKodPocztowy.Text);
                    command.Parameters.AddWithValue("@NabywcaMiasto", NabywcaMiasto.Text);

                    command.ExecuteNonQuery();
                }



            }
        }


    }
}