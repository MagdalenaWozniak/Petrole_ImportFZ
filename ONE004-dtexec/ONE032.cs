using System;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using CDNBase;
using ExcelDataReader;

using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ONE032
{
    public partial class ONE032WIN : Form
    {
        public static string sKnt_Kod = "SCA";
        public static string Operator = "BARTEK";
        public static string Haslo = "";
        public static string Firma = "Demo srodki";
        protected static CDNBase.Application Application = null;
        protected static ILogin Login = null;



        public ONE032WIN()
        {
            InitializeComponent();
            string[] akronimy = new[] { "GASPOL", "ANWIM", "UNIMOT", "BARTER_LPG", "ARAMCO", "BARTER_BUTLE" };
            knt.Items.AddRange(akronimy);
            knt.SelectedItem = "GASPOL";
            string towary = "SELECT Twr_Kod FROM CDN.TwrKarty WHERE Twr_Typ=1";
            JakaBaza("klient test");
            LogowanieZObslogaBladu();
            CDNBase.AdoSession Sesja2 = Login.CreateSession();
            object dummy = Type.Missing;
            var wynikQuery = Sesja2.Connection.Execute(towary, out dummy, 0);

            string[] kody = new string[wynikQuery.RecordCount];

            for (int j = 0; j < wynikQuery.RecordCount; j++)
            {
                kody[j] = (wynikQuery.Fields[0].Value).ToString().Trim();

                //MessageBox.Show((wynikQuery.Fields[0].Value).ToString().Trim());
                wynikQuery.MoveNext();

            }
            twr.Items.AddRange(kody);
            twr.SelectedItem = "00055P";
            Wylogowanie();

        }

        #region funkcje optimy

        void JakaBaza(string baza)
        {
            if (baza == "one")
            {
                Operator = "BARTEK";
                Haslo = "";
                Firma = "Demo srodki";
                sKnt_Kod = "SCA";
            }
            else if (baza == "klient")
            {
                Operator = "ADMIN";
                Haslo = "NE!w0rk12";
                Firma = "SCA Petrole";
                sKnt_Kod = "ORLEN PALIWA";
                baza = "CDN_SCA_Petrole";
                // Firma = "SCA_Petrole Test";
                // sKnt_Kod = "ORLEN";
            }
            else if (baza == "klient test")
            {
                Operator = "ADMIN";
                Haslo = "NE!w0rk12";
                Firma = "SCA_Petrole 24092024";
                baza = "CDN_SCA_Petrole_Testy_Orlen";
                sKnt_Kod = "ORLEN PALIWA";
                // Firma = "SCA_Petrole Test";
                // sKnt_Kod = "ORLEN";
            }

        }

        static protected string LogowanieAutomatyczne()
        {
            object[] hPar = new object[] {
                         0,  0,   0,  0,  0,   0,  0,    0,   0,   0,   0,   0,   0,   0,  0,  1, 1 };    // do jakich modułów się logujemy
            /* Kolejno: KP, KH, KHP, ST, FA, MAG, PK, PKXL, CRM, ANL, DET, BIU, SRW, ODB, KB, KBP, HAP*/
            // katalog, gdzie jest zainstalowana Optima (bez ustawienia tej zmiennej nie zadziała, chyba że program odpalimy z katalogu O!)
            System.Environment.CurrentDirectory = @"C:\Program Files (x86)\Comarch ERP Optima";
            // tworzymy nowy obiekt apliakcji
            Application = new CDNBase.Application();
            Application.HASPServerCommandLine = "172.18.34.50\\OPTIMA::5000242577";
            // blokujemy
            Application.LockApp(1);//(513, 5000, null, null, null, null);
            // logujemy się do podanej Firmy, na danego operatora, do podanych modułów
            Login = Application.Login(vUser: Operator, vChk: Haslo, vFirm: Firma, _ModulKP: hPar[0], _ModulKH: hPar[1], _ModulKHP: hPar[2], _ModulST: hPar[3], _ModulFA: hPar[4], _ModulMAG: hPar[5], _ModulPK: hPar[6],
                _ModulPKXL: hPar[7], _ModulCRM: hPar[8], _ModulANL: hPar[9], _ModulDET: hPar[10], _ModulBIU: hPar[11], _ModulSRW: hPar[12], _ModulOBD: hPar[13], _ModulKB: hPar[14], _ModulKBP: hPar[15], _ModulHAP: hPar[16]);
            return Application.HASPKeyServer;
        }

        protected static void Wylogowanie()
        {
            // niszczymy Login
            Login = null;
            // odblokowanie (wylogowanie) O!
            Application.UnlockApp();
            // niszczymy obiekt Aplikacji
            Application = null;
        }
        #endregion

        void WyjscieZApki(string a)
        {
            Cursor.Current = Cursors.Default;
            Wylogowanie();
            if (a == "done")
            {
                MessageBox.Show("Zakończono generacje FZ.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Generowanie FZ zakończyło się niepowodzeniem.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            this.Close();
            return;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                AppDomain.CurrentDomain.UnhandledException +=
                     new UnhandledExceptionEventHandler(HandleException);
                #region excel5

                #region Pobranie danych z pliku
                Cursor.Current = Cursors.WaitCursor;
                if (tB1.Text == "nie wybrano pliku")
                {
                    MessageBox.Show("Nie wybrano pliku arkusza.", "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string filePath = tB1.Text;
                DataTable dataTable0;
                //DataTable dataTable;
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    IExcelDataReader reader;
                    reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                    var conf = new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true
                        }
                    };

                    var dataSet = reader.AsDataSet(conf);
                    dataTable0 = dataSet.Tables[0];
                    //var dataTable = dataTable0.Select("[Data dostawy] > '01/01/1950'");

                }
                //#endregion

                string typArkusza = JakiTypZaciagnietegoArkusza(dataTable0);

                //Przypisywanie danych z datasetu do listy przyszłych pozycji.
                List<ONE032Data> listaPozycjiPz = new List<ONE032Data>();

                SprawdzamCzyArkuszPusty(dataTable0);

                SprawdzamCzyBartex(typArkusza);

                string pierwszaDataTemp = SprawdzanieZgodnosciDat(dataTable0, typArkusza);


                int ileWierszy = 0;
                if (typArkusza == "ORLEN")
                {
                    ileWierszy = dataTable0.Rows.Count - 1;

                }
                else if (typArkusza == "GASPOL")
                {
                    ileWierszy = dataTable0.Rows.Count;

                }
                else if (typArkusza == "BARTER_BUTLE")
                {
                    ileWierszy = dataTable0.Rows.Count - 1;

                }
                else
                {
                    ObslugaInnegoFormatuTODO("Wybrano arkusz o nie obsługiwanej formie, lub nie rozpoznano w arkuszu stukrury odpowiedniej dla arkuszy ORLEN i GASPOL.");

                }
                for (int i = 0; i < ileWierszy; i++)
                {
                    try
                    {
                        if (typArkusza == "ORLEN")
                        {
                            listaPozycjiPz.Add(new ONE032Data()
                            {
                                NazwaBazy = Convert.ToString(dataTable0.Rows[i][1]),
                                NumerZlecenia = Convert.ToString(dataTable0.Rows[i][12]),
                                Material = Convert.ToString(dataTable0.Rows[i][11]),
                                Odbiorca = Convert.ToString(dataTable0.Rows[i][2]),
                                MiejsceDostawy = Convert.ToString(dataTable0.Rows[i][3]),
                                NrDowoduWydania = Convert.ToString(dataTable0.Rows[i][4]),
                                IloscWM15 = Convert.ToDouble(dataTable0.Rows[i][5]) * 1000,
                                Cena = Convert.ToDecimal(dataTable0.Rows[i][6]) / 1000,
                                WartoscNetto = Convert.ToDecimal(dataTable0.Rows[i][7]),
                                KwotaPodatku = Convert.ToDecimal(dataTable0.Rows[i][8]),
                                WartoscBrutto = Convert.ToDecimal(dataTable0.Rows[i][9])
                            }); ;
                        }
                        else if (typArkusza == "GASPOL")
                        {
                            listaPozycjiPz.Add(new ONE032Data()
                            {
                                NazwaBazy = "",
                                NumerZlecenia = Convert.ToString(dataTable0.Rows[i][2]),
                                Material = twr.SelectedItem.ToString(),
                                Odbiorca = Convert.ToString(dataTable0.Rows[i][3]),
                                MiejsceDostawy = Convert.ToString(dataTable0.Rows[i][4]),
                                NrDowoduWydania = Convert.ToString(dataTable0.Rows[i][0]),
                                IloscWM15 = Convert.ToDouble(dataTable0.Rows[i][7]),
                                Cena = Convert.ToDecimal(dataTable0.Rows[i][8]),
                                WartoscNetto = Convert.ToDecimal(dataTable0.Rows[i][9]),
                                KwotaPodatku = Convert.ToDecimal(dataTable0.Rows[i][10]),
                                WartoscBrutto = Convert.ToDecimal(dataTable0.Rows[i][11])
                            }); ;
                        }
                        else if (typArkusza == "BARTER_BUTLE")
                        {
                            listaPozycjiPz.Add(new ONE032Data()
                            {
                                NazwaBazy = "",
                                NumerZlecenia = Convert.ToString(dataTable0.Rows[i][0]),
                                Material = Convert.ToString(dataTable0.Rows[i][4]), //twr.SelectedItem.ToString(),
                                Odbiorca = Convert.ToString(dataTable0.Rows[i][3]),
                                MiejsceDostawy = Convert.ToString(dataTable0.Rows[i][12]),
                                NrDowoduWydania = Convert.ToString(dataTable0.Rows[i][0]),
                                IloscWM15 = Convert.ToDouble(dataTable0.Rows[i][6]),
                                Cena = Convert.ToDecimal(dataTable0.Rows[i][8]),
                                WartoscNetto = Convert.ToDecimal(dataTable0.Rows[i][9]),
                                KwotaPodatku = Convert.ToDecimal(dataTable0.Rows[i][10]),
                                WartoscBrutto = Convert.ToDecimal(dataTable0.Rows[i][11])
                            }); ;
                        }
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show("Niepoprawny format zmiennych podanych w pliku arkusza.\n+Kolumna:"+i+"\n" + exp.Message, "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                Cursor.Current = Cursors.Default;
                #endregion

                #region dzialanie na optimie
                #region logowanie
                //Logowanie do O! bez wyświetlania okienka logowania

                LogowanieZObslogaBladu();
                Cursor.Current = Cursors.WaitCursor;
                CDNBase.AdoSession Sesja = Login.CreateSession();

                #endregion

                #region Sprawdzanie zgodności pól kolumny "Odbiorca(stacja)"

                List<string> listOdbiorcowBezAtr = new List<string>();

                foreach (var itemm in listaPozycjiPz)
                {
                    try
                    {
                        //pobieranie akronimu kontrahenta na podstawie wartości atrybutuKnt z importu
                        object dummy = Type.Missing;
                        var sqlQuery1 = "SELECT TOP(1)[KnA_PodmiotId] FROM [CDN].[KntAtrybuty] where KnA_WartoscTxt =  '" + itemm.Odbiorca + "'";
                        var wynikQuery = Sesja.Connection.Execute(sqlQuery1, out dummy, 0);
                        if (wynikQuery.EOF)
                        {
                            listOdbiorcowBezAtr.Add(itemm.Odbiorca);
                        }
                    }
                    catch (Exception en)
                    {
                        MessageBox.Show("Błąd w pobieraniu akronimu kontrahenta na podstawie wartości atrybutuKnt z importu -> linia 251/n" + en.Message);
                    }

                }
                var disListOdbiorcowBezAtr = listOdbiorcowBezAtr.Distinct().ToList<string>();
                if (disListOdbiorcowBezAtr.Count == 1)
                {

                    string combindedString = string.Join(",", disListOdbiorcowBezAtr.Distinct());
                    MessageBox.Show("Nie udało się znaleść kontrahenta odpowiadającego wartości kolumny Odbiorca: \"" + combindedString + "\" Proszę sprawdzić, czy dany kontrahent posiada wypełniony atrybut ORLEN_ODBIORCA", "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    WyjscieZApki("");
                    return;
                }
                else if (listOdbiorcowBezAtr.Count > 1)
                {
                    string combindedString = string.Join(",\n", disListOdbiorcowBezAtr.Distinct());
                    MessageBox.Show("Nie udało się znaleść kontrahentów odpowiadających wartościom kolumny Odbiorca: \n\"" + combindedString + "\"\n Proszę sprawdzić, czy dany kontrahenci posiadają wypełnione atrybut ORLEN_ODBIORCA", "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    WyjscieZApki("");
                    return;
                }

                #endregion


                try
                {
                    try
                    {
                        CDNHlmn.DokumentyHaMag PeZety = (CDNHlmn.DokumentyHaMag)Sesja.CreateObject("CDN.DokumentyHaMag", null);
                        CDNHlmn.IDokumentHaMag PeZet = (CDNHlmn.IDokumentHaMag)PeZety.AddNew(null);
                        CDNTwrb1.Towary zTowary = (CDNTwrb1.Towary)(Sesja.CreateObject("CDN.Towary", null));
                        var kontrahenci = (CDNBase.ICollection)(Sesja.CreateObject("CDN.Kontrahenci", null));

                       /* if (typArkusza == "ORLEN")
                        {
                            sKnt_Kod = "ORLEN PALIWA";

                        }
                        else if (typArkusza == "GASPOL")
                        {*/
                            switch (knt.SelectedItem.ToString())
                            {
                                case "GASPOL":
                                    {
                                        sKnt_Kod = "GASPOL";
                                        break;
                                    }
                                case "ANWIM":
                                    {
                                        sKnt_Kod = "ANWIM";
                                        break;
                                    }
                                case "UNIMOT":
                                    {
                                        sKnt_Kod = "UNIMOT_PALIWA";
                                        break;
                                    }
                                case "BARTER_LPG":
                                    {
                                        sKnt_Kod = "BARTER";
                                        break;
                                    }
                            case "BARTER_BUTLE":
                                {
                                    sKnt_Kod = "BARTER";
                                    break;
                                }
                            case "ARAMCO":
                                    {
                                        sKnt_Kod = "ARAMCO_FUELS_POLAND";
                                        break;
                                    }
                                default:
                                    {
                                        sKnt_Kod = "GASPOL";
                                        break;
                                    }


                            }
                      /*  }
                        else if (typArkusza == "BARTER_BUTLE")
                        {
                            sKnt_Kod = "BARTER";


                        }
                        else
                        {
                            ObslugaInnegoFormatuTODO("Wybrano arkusz o nie obsługiwanej formie, lub nie rozpoznano w arkuszu stukrury odpowiedniej dla arkuszy ORLEN i GASPOL.");

                        }*/
                        var kontrahent = (CDNHeal.IKontrahent)kontrahenci["Knt_Kod='" + sKnt_Kod + "'"];   //Tutaj kod ostrawki

                        //dodac typ PZ
                        PeZet.Rodzaj = 301000;
                        PeZet.TypDokumentu = 301;
                        PeZet.Bufor = 1;
                        if (typArkusza == "ODBIORCA")
                        {
                            string NumerObcyPZ = Path.GetFileName(filePath);

                            try
                            {
                                NumerObcyPZ = NumerObcyPZ.Substring(24, 9);
                                PeZet.NumerObcy = NumerObcyPZ;
                            }
                            catch (Exception)
                            {
                                //MessageBox.Show("Nie udalo się pobrać informacji z nazwy inmportowanego pliku.");
                            }
                        }


                        //Ustawiamy date
                        PeZet.DataDok = Convert.ToDateTime(pierwszaDataTemp);
                        PeZet.DataWys = Convert.ToDateTime(pierwszaDataTemp);
                        PeZet.Podmiot = (CDNHeal.IPodmiot)kontrahent;

                        Sesja.Save();

                        var id_PZ = PeZet.ID;


                        //CDNBase.ICollection pozycje = PeZet.Elementy;
                        foreach (var item in listaPozycjiPz)
                        {
                            try
                            {

                                Sesja = Login.CreateSession();
                                PeZet = Sesja.CreateObject("CDN.DokumentyHaMag").Item("TrN_TrNId=" + id_PZ);
                                CDNBase.ICollection pozycje = PeZet.Elementy;
                                var pozycja = (CDNHlmn.IElementHaMag)pozycje.AddNew(null);
                                //var pozycja = (CDNHlmn.IElementHaMag)pozycje.AddNew(null);




                                //pobieranie akronimu kontrahenta na podstawie wartości atrybutuKnt z importu
                                object dummy = Type.Missing;
                                var sqlQuery1 = "SELECT TOP(1)[KnA_PodmiotId] FROM [CDN].[KntAtrybuty] where KnA_WartoscTxt =  '" + item.Odbiorca + "'";
                                var wynikQuery = Sesja.Connection.Execute(sqlQuery1, out dummy, 0);
                                if (wynikQuery.EOF)
                                {
                                    MessageBox.Show("Nie udało się znaleść kontrahenta odpowiadającego wartości pola kolumny Odbiorca: \"" + item.Odbiorca + "\" Proszę sprawdzić, czy dany kontrahent posiada wypełniony atrybut ORLEN_ODBIORCA", "Komunikat",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    WyjscieZApki("");
                                    return;
                                }
                                var KntNaAtrybut = (CDNHeal.IKontrahent)kontrahenci[wynikQuery.Fields[0].Value.ToString()];
                                var KntKodNaAtrybut = KntNaAtrybut.Akronim;

                                bool wynikPoz = PrzypisanieWartosciPozycji(typArkusza, zTowary, item, pozycja, PeZet, KntKodNaAtrybut, Sesja);
                                if (wynikPoz == false)
                                {
                                    WyjscieZApki("");

                                    return;
                                }

                                //obsługa atrybutów pozycji
                                var AtrPozycji = pozycja.Atrybuty;
                                var rAtrybuty = (CDNBase.ICollection)Sesja.CreateObject("CDN.DefAtrybuty", null);
                                string dea_Kod_PZ_MIEJSCE_DOS = "ORLEN_PZ_MIEJSCE_DOS";
                                string dea_Kod_PZ_ODBIORCA = "ORLEN_PZ_ODBIORCA";
                                string dea_Kod_NR_ZLECENIA = "ORLEN_PZ_NR_ZLECENIA";
                                string dea_Kod_NR_DOW_WYD = "ORLEN_PZ_NR_DOW_WYD";

                                var Atr1 = AtrPozycji.AddNew(null);
                                Atr1.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod_PZ_MIEJSCE_DOS + "'"];
                                Atr1.Wartosc = item.MiejsceDostawy;

                                var Atr2 = AtrPozycji.AddNew(null);
                                Atr2.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod_PZ_ODBIORCA + "'"];
                                Atr2.Wartosc = KntKodNaAtrybut;

                                var Atr3 = AtrPozycji.AddNew(null);
                                Atr3.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod_NR_ZLECENIA + "'"];
                                Atr3.Wartosc = item.NumerZlecenia;

                                var Atr4 = AtrPozycji.AddNew(null);
                                Atr4.DefAtrybut = (CDNTwrb1.IDefAtrybut)rAtrybuty["dea_Kod = '" + dea_Kod_NR_DOW_WYD + "'"];
                                Atr4.Wartosc = item.NrDowoduWydania;

                                Sesja.Save();
                                Sesja.Dispose();

                            }
                            catch (Exception enx)
                            {
                                MessageBox.Show("Błąd w dodawaniu pozycji: " + enx.Message);
                            }
                        }

                        //zapisujemy
                        Sesja.Save();
                    }
                    catch(System.AccessViolationException ex)
                    {
                        MessageBox.Show("Błąd w dodawaniu całego dokumentu NAccess Validation Exception : " + ex.Message);
                    }
                }
                catch ( Exception en)//System.NullReferenceException x)
                {
                    MessageBox.Show("Błąd w dodawaniu całego dokumentu: " + en.Message);
                }

                WyjscieZApki("done");
                #endregion
            }
            catch ( Exception  ek)
            {
                MessageBox.Show("Błąd w ButtonClick: " + ek.Message);
                WyjscieZApki("");

            }

        }

        static void HandleException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show($"Sorry, there was a problem occurs. The program is terminated. \n {e.ExceptionObject}");
        }
        private bool PrzypisanieWartosciPozycji(string typArkusza, CDNTwrb1.Towary zTowary, ONE032Data item, CDNHlmn.IElementHaMag pozycja, CDNHlmn.IDokumentHaMag pezet, string kntKod, AdoSession Sesja)
        {
            //przypisywanie wartości pozycji
            try
            {
                try
                {
                    if (typArkusza == "ORLEN")
                    {
                        CDNTwrb1.ITowar zTowar = (CDNTwrb1.ITowar)zTowary["Twr_KodDostawcy='" + item.Material + "'"];
                        pozycja.TowarKod = zTowar.Kod;
                    }
                    else if (typArkusza == "GASPOL")
                    {
                        pozycja.TowarKod = item.Material;
                    }
                    else if (typArkusza == "BARTER_BUTLE")
                    {
                        object dummy = Type.Missing;
                        var sqlQuery1 = "SELECT TOP(1) [TwA_TwrID] FROM [CDN].[TwrAtrybuty] where TwA_WartoscTxt =  '" + item.Material + "'";
                        var wynikQuery = Sesja.Connection.Execute(sqlQuery1, out dummy, 0);
                        if (!wynikQuery.EOF)
                        {
                            pozycja.TowarID = Convert.ToInt32(wynikQuery.Fields[0].Value.ToString());
                        }
                        else
                        {
                            MessageBox.Show("Brak towaru: " + item.Material, "Komunikat",
                                                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }
                    }
                    pozycja.Ilosc = item.IloscWM15;
                    pozycja.Cena0 = item.Cena;
                    pozycja.Cena0WD = item.Cena;
                    pozycja.WartoscNetto = item.WartoscNetto;
                    pozycja.WartoscBrutto = item.WartoscBrutto;
                    Sesja.Save();
                    var rDostawy = pozycja.Dostawy;
                    string cecha_kod = "NR_DOSTAWY";
                    try
                    {
                        var rDostawa = rDostawy.AddNew();
                        pozycja.UstawDostawy();
                        rDostawa.Cecha1_Kod = cecha_kod;
                        rDostawa.Cecha1_Wartosc = pezet.NumerPelny + "#" + kntKod + "#" + pezet.DataDok.ToShortDateString();
                        rDostawa.Ilosc = item.IloscWM15;
                        pozycja.UsunDostawy();

                    }
                    catch (Exception empty)
                    {
                        MessageBox.Show("Błąd w dodawaniu zasobu: " + Convert.ToString(empty), "Komunikat",
                                                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                    Sesja.Save();
                }
                catch(System.AccessViolationException ex)
                {
                    MessageBox.Show("Przypisywanie wartości pozycji -> linia 584 /n" + ex.Message);
                    return false;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Przypisywanie wartości pozycji -> linia 411 /n" + e.Message);
                return false;
            }
            return true;

        }

        private static void LogowanieZObslogaBladu()
        {
            try
            {
                //MessageBox.Show(LogowanieAutomatyczne());
                LogowanieAutomatyczne();
            }
            catch (COMException eo)
            {
                if (eo.InnerException != null)
                {
                    MessageBox.Show("Błąd logowania: " + eo.InnerException.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    System.Windows.Forms.Application.Exit();
                }
                else
                {
                    MessageBox.Show("Błąd logowania: " + eo.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    System.Windows.Forms.Application.Exit();
                }
            }
            catch (Exception ei)
            {
                if (ei.InnerException != null)
                {
                    MessageBox.Show("Błąd logowania: " + ei.InnerException.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    System.Windows.Forms.Application.Exit();
                }
                else
                {
                    MessageBox.Show("Błąd logowania: " + ei.Message, "Komunikat",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    System.Windows.Forms.Application.Exit();
                }


            }
        }

        private static void SprawdzamCzyArkuszPusty(DataTable dataTable)
        {
            //Sprawdzam czy arkusz nie jest pusty
            if (dataTable.Rows.Count == 0)
            {
                MessageBox.Show("W arkuszu brak rekordów.", "Komunikat",
                MessageBoxButtons.OK, MessageBoxIcon.Error);//błąd
                System.Windows.Forms.Application.Exit();

            }
        }

        private static void SprawdzamCzyBartex(string typArkusza)
        {
            if (typArkusza == "BARTEX")
            {
                ObslugaInnegoFormatuTODO("Wybrano arkusz o nie obsługiwanej formie, lub nie rozpoznano w arkuszu stukrury odpowiedniej dla arkuszy ORLEN i GASPOL.");
            }
        }

        private static string SprawdzanieZgodnosciDat(DataTable dataTable, string typArkusza)
        {
            //Sprawdzam, czy wszystkie daty się zgadzaja
            string pierwszaDataTemp = "";
            if (typArkusza == "ORLEN")
            {
                pierwszaDataTemp = Convert.ToString(dataTable.Rows[0][0]);

            }
            else if (typArkusza == "GASPOL")
            {
                pierwszaDataTemp = Convert.ToString(dataTable.Rows[0][1]);
                //MessageBox.Show(pierwszaDataTemp);

            }
            else if (typArkusza == "BARTER_BUTLE")
            {
                pierwszaDataTemp = Convert.ToString(dataTable.Rows[0][1]);
                //MessageBox.Show(pierwszaDataTemp);

            }
            else
            {
                ObslugaInnegoFormatuTODO("Wybrano arkusz o nie obsługiwanej formie, lub nie rozpoznano w arkuszu stukrury odpowiedniej dla arkuszy ORLEN i GASPOL.");

            }

            int ileWierszy = 0;
            if (typArkusza == "ORLEN")
            {
                ileWierszy = dataTable.Rows.Count - 1;

            }
            else if (typArkusza == "GASPOL")
            {
                ileWierszy = dataTable.Rows.Count;

            }
            else if (typArkusza == "BARTER_BUTLE")
            {
                ileWierszy = dataTable.Rows.Count - 1;

            }
            else
            {
                ObslugaInnegoFormatuTODO("Wybrano arkusz o nie obsługiwanej formie, lub nie rozpoznano w arkuszu stukrury odpowiedniej dla arkuszy ORLEN i GASPOL.");

            }

            for (int i = 1; i < ileWierszy; i++)
            {

                if (typArkusza == "ORLEN")
                {
                    string tempp = Convert.ToString(dataTable.Rows[i][0]);

                    if (tempp != pierwszaDataTemp)
                    {
                        MessageBox.Show("W wybranym arkuszu wprowadzone daty są ze sobą niezgodne. Wszystkie pozycje powinny mieć tę samą datę dostawy.", "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                        System.Windows.Forms.Application.Exit();
                    }
                }
                else
                {
                    string tempp = Convert.ToString(dataTable.Rows[i][1]);

                    if (tempp != pierwszaDataTemp)
                    {
                        MessageBox.Show("W wybranym arkuszu wprowadzone daty są ze sobą niezgodne. Wszystkie pozycje powinny mieć tę samą datę dostawy.\nBłąd dla wiersza" + i, "Komunikat",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                        System.Windows.Forms.Application.Exit();
                    }
                }
            }

            return pierwszaDataTemp;
        }

        private static void ObslugaInnegoFormatuTODO(string messageStr)
        {
            //"Wybrano arkusz o nie obsługiwanej formie, lub nie rozpoznano w arkuszu stukrury odpowiedniej dla arkuszy ORLEN i GASPOL."
            MessageBox.Show(messageStr, "Komunikat",
            MessageBoxButtons.OK, MessageBoxIcon.Error);
            System.Windows.Forms.Application.Exit();
        }

        private static string JakiTypZaciagnietegoArkusza(DataTable dataTable)
        {
            //sprawdzam czy dokument ma formę arkusza ORLEN, czy Gazpol czy Bartex
            string poleZeroZero = dataTable.Columns[0].ColumnName;
            string typZaciagnietegoDok;
            if (poleZeroZero == "Data dostawy")
            {
                typZaciagnietegoDok = "ORLEN";
            }
            else if (poleZeroZero == "Nr dowodu wydania")
            {
                typZaciagnietegoDok = "GASPOL";
            }
            else if (poleZeroZero == "Numer zew.")
            {
                typZaciagnietegoDok = "BARTER_BUTLE";
            }
            else
            {
                typZaciagnietegoDok = "BARTEX";
            }
            return typZaciagnietegoDok;
        }
        #region wybranie pliku
        void WybraniePliku()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog()
            {
                InitialDirectory = Directory.GetCurrentDirectory(),
                DefaultExt = "xlsx",
                Filter = "Pliki xlsx (*.xlsx)|*.xlsx|Pliki xls (*.xls)|*.xls|Wszystkie pliki (*.*)|*.*",
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tB1.Text = openFileDialog1.FileName;
                //MessageBox.Show(path);
            }
            return;
        }
        private void Button2_Click(object sender, EventArgs e)
        {
            WybraniePliku();
        }
        private void TextBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            WybraniePliku();
        }
        #endregion
    }
}
#endregion