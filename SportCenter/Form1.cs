using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace SportCenter
{

    public partial class Form1 : Form
    {
        // Kullanıcı oturum bilgilerini tutan global sınıf
        public static class UserSession
        {
            public static int CurrentUserId { get; set; } = 0;
        }
        int GuncellenenRezervasyonID = 0;

        public Form1()
        {
            InitializeComponent();

        }

        string dbPath = @"Data Source=C:\Users\zeroo\Desktop\vss-istatistik(düzeltme olucak)\vss-istatistik(düzeltme olucak)\vss\SportCenter\SporCenter.db;
        Version=3;
Journal Mode=WAL;
         BusyTimeout=5000;";

       // public static int GirisYapanKullaniciID = 0;
        int SecilenEgitmenID = 0;
        double sahaUcret = 0;
        double egitmenUcret = 0;
        int guncellenecekBookingID = 0;
        double eskiOdenenTutar = 0;
        bool GuncellemeModu = false;
        double farkTutari = 0;
        bool FarkOdemeModu = false;
        double OdenecekFark = 0;





        // 1. PROGRAM AÇILINCA (LOAD)

        private void Form1_Load(object sender, EventArgs e)
        {

            using (var con = new SQLiteConnection(dbPath))
            {
                con.Open();
                using (var cmd = new SQLiteCommand("PRAGMA journal_mode=WAL;", con))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            ListViewKolonlariHazirla();
            flpEgitmenler.Visible = false; // ← Panel başlangıçta gizli!
            dtpTarih.MinDate = DateTime.Today;

            using (var con = new SQLiteConnection(dbPath))
            {
                con.Open();
                using (var da = new SQLiteDataAdapter("SELECT DISTINCT SportType FROM Fields", con))
                {
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    cmbSporTuru.DataSource = dt;
                    cmbSporTuru.DisplayMember = "SportType";
                    cmbSporTuru.ValueMember = "SportType";
                }
            }

         

        }



        private void btnKayitOl_Click(object sender, EventArgs e)
        {
            if (txtAdSoyad.Text == "" || txtKullaniciKayit.Text == "" || txtSifreKayit.Text == "")
            {
                MessageBox.Show("Lütfen zorunlu alanları doldurunuz.");
                return;
            }

            DatabaseHelper db = new DatabaseHelper();

            string sql = "INSERT INTO Users (NameSurname, Phone, Email, Username, Password, Role) " +
                         "VALUES (@n, @ph, @em, @u, @p, @r)";

            SQLiteParameter[] p = {
        new SQLiteParameter("@n", txtAdSoyad.Text),
        new SQLiteParameter("@ph", mskTelefon.Text),
        new SQLiteParameter("@em", txtEmail.Text),
        new SQLiteParameter("@u", txtKullaniciKayit.Text),
        new SQLiteParameter("@p", txtSifreKayit.Text),
        new SQLiteParameter("@r", "User") // 🔥 yeni eklenen Role default User
    };

            if (db.ExecuteQuery(sql, p))
            {
                MessageBox.Show("Kayıt Başarılı!");
                tabControl1.SelectedTab = tabGiris;

                // Temizleme
                txtAdSoyad.Clear();
                mskTelefon.Clear();
                txtEmail.Clear();
                txtKullaniciKayit.Clear();
                txtSifreKayit.Clear();
            }
            else
            {
                MessageBox.Show("Kayıt sırasında bir hata oluştu!");
            }
        }

        // "Hesabınız yok mu? Kayıt Ol" Butonuna tıklayınca
        private void btnKayitGecis_Click(object sender, EventArgs e)
        {
            // Bizi Kayıt Ol sekmesine (tabPage2) atar
            tabControl1.SelectedTab = tabKayit;
        }

        private void btnGiris_Click(object sender, EventArgs e)
        {
            DatabaseHelper db = new DatabaseHelper();
            // Id ve Role bilgisini çekiyoruz
            string sql = "SELECT Id, Role FROM Users WHERE Username=@u AND Password=@p";

            SQLiteParameter[] p = {
        new SQLiteParameter("@u", txtKullanici.Text),
        new SQLiteParameter("@p", txtSifre.Text)
    };

            DataTable dt = db.GetData(sql, p);

            if (dt.Rows.Count > 0)
            {
                // --- HATALI KISIM BURASIYDI, ŞÖYLE DEĞİŞTİR: ---
                // GirisYapanKullaniciID yerine global sınıfımızı kullanıyoruz:
                UserSession.CurrentUserId = Convert.ToInt32(dt.Rows[0]["Id"]);
                // ----------------------------------------------

                string role = dt.Rows[0]["Role"].ToString();

                ((Control)tabAnaSayfa).Enabled = true;
                ((Control)tabOdeme).Enabled = true;

                if (role == "Admin")
                {
                    ((Control)tabIstatistik).Enabled = true;
                    tabControl1.SelectedTab = tabIstatistik;
                }
                else
                {
                    tabControl1.SelectedTab = tabAnaSayfa;
                }

                // Kullanıcıya giriş yaptığını hissettirmek için ufak bir mesaj iyi olabilir
                // MessageBox.Show("Hoşgeldiniz!"); 
            }
            else
            {
                MessageBox.Show("Hatalı giriş!");
            }
        }

        private void chkEgitmen_CheckedChanged(object sender, EventArgs e)
        {
            if (chkEgitmen.Checked)
            {
                flpEgitmenler.Visible = true;
                EgitmenleriGetir();
            }
            else
            {
                flpEgitmenler.Visible = false;
                SecilenEgitmenID = 0;
                egitmenUcret = 0;
                Hesapla(); // 🔥 ÜCRET DÜŞSÜN
            }
        }
        private void EgitmenleriGetir()
        {
            flpEgitmenler.Controls.Clear();

            using (var con = new SQLiteConnection(dbPath))
            {
                con.Open();
                using (var da = new SQLiteDataAdapter(
                    "SELECT * FROM Trainers WHERE SportType=@s", con))
                {
                    da.SelectCommand.Parameters.AddWithValue("@s", cmbSporTuru.SelectedValue);

                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    foreach (DataRow r in dt.Rows)
                    {
                        Button btn = new Button();
                        btn.Width = 200;
                        btn.Height = 55;
                        btn.BackColor = Color.MintCream;
                        btn.ForeColor = Color.Black;
                        btn.Text = r["NameSurname"] + " - " + r["PricePerHour"] + " ₺";
                        btn.Tag = r;

                        btn.Click += (s, e) =>
                        {
                            var row = (DataRow)((Button)s).Tag;
                            SecilenEgitmenID = Convert.ToInt32(row["ID"]);
                            egitmenUcret = Convert.ToDouble(row["PricePerHour"]);
                            Hesapla();
                        };

                        flpEgitmenler.Controls.Add(btn);
                    }
                }
            }
        }


        DataRow secilenEgitmen = null;



        private void EgitmenSecildi(object sender, EventArgs e)
        {
            RadioButton rb = sender as RadioButton;
            secilenEgitmen = rb.Tag as DataRow;
            egitmenUcret = Convert.ToInt32(secilenEgitmen["Ucret"]);
            Hesapla();
        }
        private void Hesapla()
        {
            double toplam = sahaUcret + egitmenUcret;
            lblToplamUcret.Text = toplam + " ₺";
        }
        private void OdemeTabinaGec()
        {
            // 🔴 KART BİLGİLERİNİ TEMİZLE
            txtKartSahibi.Clear();
            mskKartNo.Clear();
            mskSKT.Clear();
            mskCVV.Clear();

            // (istersen)
            txtKartSahibi.Focus();

            // 🔵 TAB'A GEÇ
            tabControl1.SelectedTab = tabOdeme;
        }

        private void btnSahaSec_Click(object sender, EventArgs e)
        {
            // -----------------------------------------------------------
            // 1. GÜVENLİK KONTROLÜ: Kullanıcı Giriş Yapmış mı?
            // -----------------------------------------------------------
            if (UserSession.CurrentUserId == 0)
            {
                MessageBox.Show("Ödeme işlemine devam edebilmek için lütfen önce giriş yapınız veya kayıt olunuz.",
                                "Giriş Gerekli",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);

                // Kullanıcıyı Giriş sekmesine at
                tabControl1.SelectedTab = tabGiris;

                // Metodu burada kes, aşağıya inip ödeme sayfasına geçmesin.
                return;
            }
            double guncelToplam = sahaUcret + egitmenUcret;

            // 🔵 NORMAL REZERVASYON
            if (!GuncellemeModu)
            {
                OdemeSayfasiBilgileriDoldur();
                //  OdemeButonlariniAyarla(false); // ✅ normal ödeme
                OdemeTabinaGec();
                return;
            }

            // 🔴 GÜNCELLEME MODU
            if (guncelToplam <= eskiOdenenTutar)
            {
                RezervasyonuGuncelle(guncelToplam);

                MessageBox.Show(
                    guncelToplam < eskiOdenenTutar
                    ? "Fazla tutar kartınıza iade edilecektir."
                    : "Rezervasyon güncellendi."
                );

                GuncellemeModu = false;
                RezervasyonlariGetir();
                return;
            }

            // 🔥 FARK VAR → ÖDEME EKRANI
            farkTutari = guncelToplam - eskiOdenenTutar;
            lblToplamValue.Text = farkTutari + " ₺";

            //  OdemeButonlariniAyarla(true); // ✅ sadece fark öde
            tabControl1.SelectedTab = tabOdeme;
        }




        private void cmbSporTuru_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSporTuru.SelectedIndex == -1) return;

            using (var con = new SQLiteConnection(dbPath))
            {
                con.Open();
                using (var da = new SQLiteDataAdapter(
                    "SELECT Id, Name FROM Fields WHERE SportType=@s", con))
                {
                    da.SelectCommand.Parameters.AddWithValue("@s", cmbSporTuru.SelectedValue);

                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    cmbSaha.DataSource = dt;
                    cmbSaha.DisplayMember = "Name";
                    cmbSaha.ValueMember = "Id";
                }
            }

            flpEgitmenler.Controls.Clear();
            egitmenUcret = 0;
            Hesapla();

        }


        private void dtpTarih_ValueChanged(object sender, EventArgs e)
        {
            LoadSaat();

        }

        private void cmbSaha_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSaha.SelectedIndex == -1) return;

            using (var con = new SQLiteConnection(dbPath))
            {
                con.Open();
                using (var cmd = new SQLiteCommand("SELECT PricePerHour FROM Fields WHERE Id=@id", con))
                {
                    cmd.Parameters.AddWithValue("@id", cmbSaha.SelectedValue);
                    sahaUcret = Convert.ToDouble(cmd.ExecuteScalar());
                }
            }

            Hesapla();
            LoadSaat();

        }
        private void LoadSaat()
        {
            if (cmbSaha.SelectedIndex == -1) return;

            // 1. Standart saat listesi
            var saatler = new List<string>()
            {
                "09:00","10:00","11:00","12:00","13:00",
                "14:00","15:00","16:00","17:00","18:00",
                "19:00","20:00"
            };

            // 2. 🔥 BUGÜNSE GEÇMİŞ SAATLERİ SİLME KISMI BURADA 🔥
            if (dtpTarih.Value.Date == DateTime.Today)
            {
                int suankiSaat = DateTime.Now.Hour;
                // Örnek: Saat 14:00 ise 14 ve öncesini siler.
                saatler.RemoveAll(s => Convert.ToInt32(s.Substring(0, 2)) <= suankiSaat);
            }

            // 3. Veritabanından dolu saatleri silme
            using (var con = new SQLiteConnection(dbPath))
            {
                con.Open();
                using (var cmd = new SQLiteCommand(
                    "SELECT Time FROM Bookings WHERE FieldID=@f AND Date=@d", con))
                {
                    cmd.Parameters.AddWithValue("@f", cmbSaha.SelectedValue);
                    cmd.Parameters.AddWithValue("@d", dtpTarih.Value.ToShortDateString());

                    using (var dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                            saatler.Remove(dr["Time"].ToString());
                    }
                }
            }

            cmbSaat.DataSource = saatler;
        }
        private void cmbIstatistikSecimi_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbIstatistikSecimi.SelectedItem == null) return;

            string secim = cmbIstatistikSecimi.SelectedItem.ToString();

            // 1. ÖNCEKİ HER ŞEYİ TEMİZLE (Legend Dahil!)
            lblOzet.Text = "";
            GrafikTemizle();

            using (var con = new System.Data.SQLite.SQLiteConnection(dbPath))
            {
                try
                {
                    con.Open();

                    switch (secim)
                    {
                        // ---------------------------------------------------------
                        // 1. GELİR DAĞILIMI (Pasta Grafik)
                        // ---------------------------------------------------------
                        case "Gelir Dağılımı (Pasta Grafik)":
                            GorunumAyarla(grafikAcik: true);

                            chart1.Titles.Add("Saha Bazlı Ciro Dağılımı (Yüzdesel)");

                            // Legend Ekleme (Temizlendiği için sıfırdan ekliyoruz)
                            var legend = chart1.Legends.Add("Legend1");
                            legend.Docking = Docking.Right;
                            legend.Alignment = StringAlignment.Center;
                            legend.Font = new Font("Arial", 10, FontStyle.Regular);
                            legend.Title = "Saha ve Toplam Gelir Listesi"; // Sorun çıkaran yazı buydu, şimdi sadece burada çıkacak.
                            legend.TitleFont = new Font("Arial", 10, FontStyle.Bold);

                            var seriPie = chart1.Series.Add("Gelirler");
                            seriPie.ChartType = SeriesChartType.Pie;
                            seriPie["PieLabelStyle"] = "Inside";

                            string sqlPie = @"SELECT F.Name as SahaAdi, SUM(B.TotalPrice) as ToplamPara 
                                      FROM Bookings B 
                                      LEFT JOIN Fields F ON B.FieldID = F.Id 
                                      GROUP BY F.Name";

                            using (var cmd = new System.Data.SQLite.SQLiteCommand(sqlPie, con))
                            using (var dr = cmd.ExecuteReader())
                            {
                                while (dr.Read())
                                {
                                    string saha = dr["SahaAdi"] != DBNull.Value ? dr["SahaAdi"].ToString() : "Tanımsız";
                                    double para = dr["ToplamPara"] != DBNull.Value ? Convert.ToDouble(dr["ToplamPara"]) : 0;

                                    if (para > 0)
                                    {
                                        int pIndex = seriPie.Points.AddXY(saha, para);
                                        seriPie.Points[pIndex].Label = "#PERCENT{P1}";
                                        seriPie.Points[pIndex].LabelForeColor = Color.White;
                                        seriPie.Points[pIndex].LegendText = $"{saha} ({para} ₺)";
                                    }
                                }
                            }
                            lblOzet.Text = "Dilimler yüzdelik payı, yan liste net kazancı gösterir.";
                            break;

                        // ---------------------------------------------------------
                        // 2. EN ÇOK HARCAMA YAPAN ÜYELER (Tablo)
                        // ---------------------------------------------------------
                        case "En Çok Para Harcayan Üyeler (Tablo)":
                            GorunumAyarla(grafikAcik: false);
                            string sqlTablo = @"SELECT U.NameSurname, SUM(B.TotalPrice) as ToplamHarcama, COUNT(*) as IslemSayisi 
                                        FROM Bookings B 
                                        INNER JOIN Users U ON B.UserID = U.Id 
                                        GROUP BY U.NameSurname 
                                        ORDER BY ToplamHarcama DESC";
                            VeriyiTabloyaDok(con, sqlTablo);
                            lblOzet.Text = "En değerli müşterilerimiz.";
                            break;

                        // ---------------------------------------------------------
                        // 3. SPOR DOLULUK (Sütun Grafik)
                        // ---------------------------------------------------------
                        case "Sporlara Göre Doluluk (Sütun Grafik)":
                            GorunumAyarla(grafikAcik: true);

                            // Sütun grafik için basit bir Legend ekleyelim (Başlıksız)
                            chart1.Legends.Add(new Legend("Default"));

                            var seriCol = chart1.Series.Add("Doluluk");
                            seriCol.ChartType = SeriesChartType.Column;
                            chart1.Titles.Add("Saha Kullanım Sayıları");

                            string sqlCol = @"SELECT F.Name as SahaAdi, COUNT(*) as Adet 
                                      FROM Bookings B 
                                      LEFT JOIN Fields F ON B.FieldID = F.Id 
                                      GROUP BY F.Name";

                            using (var cmd = new System.Data.SQLite.SQLiteCommand(sqlCol, con))
                            using (var dr = cmd.ExecuteReader())
                            {
                                while (dr.Read())
                                {
                                    string sahaAdi = dr["SahaAdi"] != DBNull.Value ? dr["SahaAdi"].ToString() : "Silinmiş";
                                    seriCol.Points.AddXY(sahaAdi, dr["Adet"]);
                                }
                            }
                            lblOzet.Text = "Hangi sahanın ne kadar tercih edildiği.";
                            break;

                        // ---------------------------------------------------------
                        // 4. YOĞUN SAATLER (Çizgi Grafik)
                        // ---------------------------------------------------------
                        case "Saatlere Göre Yoğunluk":
                            GorunumAyarla(grafikAcik: true);

                            // Çizgi grafik için legend
                            chart1.Legends.Add(new Legend("Default"));

                            var seriLine = chart1.Series.Add("Yogunluk");
                            seriLine.ChartType = SeriesChartType.Line;
                            seriLine.BorderWidth = 3;
                            chart1.Titles.Add("Saatlik Yoğunluk");

                            string sqlTime = "SELECT Time, COUNT(*) as Sayi FROM Bookings GROUP BY Time ORDER BY Time";

                            using (var cmd = new System.Data.SQLite.SQLiteCommand(sqlTime, con))
                            using (var dr = cmd.ExecuteReader())
                            {
                                while (dr.Read())
                                {
                                    seriLine.Points.AddXY(dr["Time"].ToString(), dr["Sayi"]);
                                }
                            }
                            lblOzet.Text = "En yoğun saatler.";
                            break;

                        // ---------------------------------------------------------
                        // 5. SON İŞLEMLER (Tablo)
                        // ---------------------------------------------------------
                        case "Son 20 İşlem Kaydı (Tablo)":
                            GorunumAyarla(grafikAcik: false);
                            string sqlList = @"SELECT B.ID, U.NameSurname as 'Üye', F.Name as 'Saha', 
                                       T.NameSurname as 'Eğitmen', B.Date, B.Time, B.TotalPrice 
                                       FROM Bookings B
                                       LEFT JOIN Users U ON B.UserID = U.Id
                                       LEFT JOIN Fields F ON B.FieldID = F.Id
                                       LEFT JOIN Trainers T ON B.TrainerID = T.Id
                                       ORDER BY B.ID DESC LIMIT 20";
                            VeriyiTabloyaDok(con, sqlList);
                            lblOzet.Text = "Son işlemler.";
                            break;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("HATA: " + ex.Message);
                }
            }
        }
        // Bu kodları Class'ın içine, diğer metodların (void'lerin) bittiği yere yapıştır.
        private void GrafikTemizle()
        {
            // 1. Serileri Temizle
            chart1.Series.Clear();

            // 2. Başlıkları Temizle
            chart1.Titles.Clear();

            // 3. Legendları (Açıklamaları) MUTLAKA Temizle
            // (Bunu yapmadığınız için yazılar karışıyordu)
            chart1.Legends.Clear();

            // 4. Alanları Temizle ve Varsayılanı Ekle
            chart1.ChartAreas.Clear();
            ChartArea area = new ChartArea();
            area.AxisY.Minimum = 0;
            area.AxisY.Maximum = double.NaN;
            area.AxisY.IntervalAutoMode = IntervalAutoMode.VariableCount;
            area.AxisY.IsStartedFromZero = true;
            chart1.ChartAreas.Add(area);
        }

        private void GorunumAyarla(bool grafikAcik)
        {
            // Grafiği gösterip tabloyu gizler, ya da tam tersini yapar.
            chart1.Visible = grafikAcik;
            dataGridView1.Visible = !grafikAcik;
        }

        private void VeriyiTabloyaDok(System.Data.SQLite.SQLiteConnection con, string sql)
        {
            // SQL sonucunu DataGridView tablosuna doldurur.
            using (var da = new System.Data.SQLite.SQLiteDataAdapter(sql, con))
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                // Sütun genişliğini içindeki en uzun yazıya göre ayarlar, sığmazsa altta kaydırma çubuğu çıkar.
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }
        }
        // Bu kodları Class'ın içine, diğer metodların (void'lerin) bittiği yere yapıştır.


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        //Ödeme Sayfası
        private void rdbKrediKarti_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbKrediKarti.Checked)
            {
                lblTaksit.Visible = true;
                cmbTaksit.Visible = true;
            }
        }

        private void rdbBankaKarti_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbBankaKarti.Checked)
            {
                lblTaksit.Visible = false;
                cmbTaksit.Visible = false;
            }
        }

        private void btnIptal_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
        "Ödemeyi iptal etmek istediğinize emin misiniz?",
        "İptal",
        MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {
                // ✔ FORMU TEMİZLE
                RezervasyonFormunuTemizle();

                // ✔ ANA SAYFAYA DÖN
                tabControl1.SelectedTab = tabAnaSayfa;
            }
        }

        private void btnOdemeYap_Click(object sender, EventArgs e)
        {
            if (txtKartSahibi.Text == "" ||
        mskKartNo.Text.Replace(" ", "").Length < 16 ||
        mskSKT.Text.Length < 5 ||
        mskCVV.Text.Length < 3)
            {
                MessageBox.Show("Kart bilgilerini eksiksiz giriniz!");
                return;
            }

            // 🔴 FARK ÖDEMESİ
            if (GuncellemeModu && farkTutari > 0)
            {
                RezervasyonuGuncelle(sahaUcret + egitmenUcret);

                MessageBox.Show("Rezervasyon güncellendi.\nFark başarıyla ödendi.");

                GuncellemeModu = false;
                farkTutari = 0;

                RezervasyonlariGetir();
                tabControl1.SelectedTab = tabAnaSayfa;
            }

            // 🟢 NORMAL REZERVASYON
            using (var con = new SQLiteConnection(dbPath))
            {
                con.Open();
                using (var cmd = new SQLiteCommand(
                    "INSERT INTO Bookings (UserId, SportID, FieldID, Date, Time, TotalPrice, TrainerID) " +
                    "VALUES (@u, @s, @f, @d, @t, @p, @tr)", con))
                {
                    cmd.Parameters.AddWithValue("@u", UserSession.CurrentUserId);                    cmd.Parameters.AddWithValue("@s", cmbSporTuru.SelectedValue);
                    cmd.Parameters.AddWithValue("@f", cmbSaha.SelectedValue);
                    cmd.Parameters.AddWithValue("@d", dtpTarih.Value.ToShortDateString());
                    cmd.Parameters.AddWithValue("@t", cmbSaat.Text);
                    double odenecekTutar;

                    if (GuncellemeModu && farkTutari > 0)
                        odenecekTutar = farkTutari;   // 🔥 SADECE FARK
                    else
                        odenecekTutar = sahaUcret + egitmenUcret;

                    cmd.Parameters.AddWithValue("@p", odenecekTutar);
                    cmd.Parameters.AddWithValue("@tr",
                        SecilenEgitmenID == 0 ? (object)DBNull.Value : SecilenEgitmenID);

                    cmd.ExecuteNonQuery();
                }
            }
            if (GuncellemeModu)
            {
                RezervasyonuGuncelle(sahaUcret + egitmenUcret);
                GuncellemeModu = false;
                farkTutari = 0;
            }


            MessageBox.Show("Ödeme başarılı! Rezervasyon oluşturuldu.");
            RezervasyonFormunuTemizle();
            tabControl1.SelectedTab = tabAnaSayfa;
        }
        private void RezervasyonFormunuTemizle()
        {
            cmbSaha.SelectedIndex = -1;
            cmbSporTuru.SelectedIndex = -1;
            cmbSaat.SelectedIndex = -1;
            chkEgitmen.Checked = false;
            flpEgitmenler.Controls.Clear();

            dtpTarih.Value = DateTime.Today;

            sahaUcret = 0;
            egitmenUcret = 0;

            lblToplamValue.Text = "0 ₺";

            //OdemeButonlariniAyarla(false); // 🔥 ÇOK ÖNEMLİ
        }
        private void OdemeSayfasiBilgileriDoldur()
        {
            lblSahaValue.Text = cmbSaha.Text;
            lblTarihValue.Text = dtpTarih.Value.ToShortDateString();
            lblSaatValue.Text = cmbSaat.Text;
            lblToplamValue.Text = (sahaUcret + egitmenUcret) + " ₺";

            // Varsayılan ödeme tipi kredi kartı
            rdbKrediKarti.Checked = true;

            cmbTaksit.Items.Clear();
            cmbTaksit.Items.Add("3 Taksit");
            cmbTaksit.Items.Add("6 Taksit");
            cmbTaksit.SelectedIndex = 0;

            lblTaksit.Visible = true;
            cmbTaksit.Visible = true;

            // ---------------------------
            //  EĞİTMEN BİLGİLERİ EKLEME
            // ---------------------------

            if (SecilenEgitmenID != 0)
            {
                // Eğitmen adı ve ücretini çekelim
                using (var con = new SQLiteConnection(dbPath))
                {
                    con.Open();
                    using (var cmd = new SQLiteCommand(
                        "SELECT NameSurname, PricePerHour FROM Trainers WHERE ID=@id", con))
                    {
                        cmd.Parameters.AddWithValue("@id", SecilenEgitmenID);
                        var rdr = cmd.ExecuteReader();

                        if (rdr.Read())
                        {
                            lblEgitmen.Visible = true;
                            lblEgitmenValue.Visible = true;
                            lblEgitmenValue.Text = $"{rdr["NameSurname"]} ({rdr["PricePerHour"]} ₺)";
                        }
                    }
                }
            }
            else
            {
                // Eğitmen seçilmediyse tamamen gizle
                lblEgitmen.Visible = false;
                lblEgitmenValue.Visible = false;
            }
            //  btnFarkOde.Visible = false;
            btnOdemeYap.Visible = true;
            //  OdemeButonlariniAyarla(false);
            if (GuncellemeModu && farkTutari > 0)
            {
                lblToplamValue.Text = (sahaUcret + egitmenUcret) + " ₺";
                lblFarkText.Visible = true;
                lblFarkValue.Visible = true;
                lblFarkValue.Text = farkTutari + " ₺";
            }


        }

        private void lvRezervasyonlar_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool seciliMi = lvRezervasyonlar.SelectedItems.Count > 0;

            btnGuncelle.Enabled = seciliMi;
            btnIptalg.Enabled = seciliMi;
        }

        private void btnRezervasyonlarim_Click(object sender, EventArgs e)
        {
            // 1. KONTROL: Eğer liste zaten açıksa -> GİZLE ve ÇIK
            if (lvRezervasyonlar.Visible == true)
            {
                lvRezervasyonlar.Visible = false;
                btnGuncelle.Visible = false;
                btnIptalg.Visible = false;
                return; // İşlemi burada bitir, aşağıya inme.
            }

            // --- BURADAN AŞAĞISI LİSTE KAPALIYSA ÇALIŞIR (AÇMA İŞLEMİ) ---

            // 2. GÜVENLİK KONTROLÜ
            if (UserSession.CurrentUserId == 0)
            {
                MessageBox.Show("Rezervasyonlarınızı görmek için lütfen önce giriş yapınız.",
                                "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tabControl1.SelectedTab = tabGiris;
                return;
            }

            // 3. VERİLERİ GETİR VE GÖRÜNÜR YAP
            try
            {
                RezervasyonlariGetir();

                // Nesneleri görünür yap
                lvRezervasyonlar.Visible = true;
                btnGuncelle.Visible = true;
                btnIptalg.Visible = true;

                // Sekmeyi aç
                tabControl1.SelectedTab = tabAnaSayfa;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Listeleme hatası: " + ex.Message);
            }

        }
        private void RezervasyonlariGetir()
        {
            lvRezervasyonlar.Items.Clear();

            using (var con = new SQLiteConnection(dbPath))
            {
                con.Open();

                using (var cmd = new SQLiteCommand(@"
SELECT 
    B.ID,
    F.SportType AS SporTuru,
    F.Name AS Saha,
    B.Date,
    B.Time,
    IFNULL(T.NameSurname, '-') AS Egitmen,
    B.TotalPrice
FROM Bookings B
JOIN Fields F ON F.ID = B.FieldID
LEFT JOIN Trainers T ON T.ID = B.TrainerID
WHERE B.UserId = @u", con))
                {
                    cmd.Parameters.AddWithValue("@u", UserSession.CurrentUserId);

                    using (var dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            ListViewItem item = new ListViewItem(dr["ID"].ToString());
                            item.SubItems.Add(dr["SporTuru"].ToString());   // 1
                            item.SubItems.Add(dr["Saha"].ToString());       // 2
                            item.SubItems.Add(dr["Date"].ToString());       // 3
                            item.SubItems.Add(dr["Time"].ToString());       // 4
                            item.SubItems.Add(dr["Egitmen"].ToString());    // 5
                            item.SubItems.Add(dr["TotalPrice"] + " ₺");     // 6

                            lvRezervasyonlar.Items.Add(item);
                        }
                    }
                }
            }
        }



        private void ListViewKolonlariHazirla()
        {
            lvRezervasyonlar.Clear();
            lvRezervasyonlar.View = View.Details;
            lvRezervasyonlar.FullRowSelect = true;
            lvRezervasyonlar.GridLines = true;

            lvRezervasyonlar.Columns.Add("ID", 0); // 0
            lvRezervasyonlar.Columns.Add("Spor Türü", 120); // 1
            lvRezervasyonlar.Columns.Add("Saha", 140); // 2
            lvRezervasyonlar.Columns.Add("Tarih", 100); // 3
            lvRezervasyonlar.Columns.Add("Saat", 80); // 4
            lvRezervasyonlar.Columns.Add("Eğitmen", 130); // 5
            lvRezervasyonlar.Columns.Add("Toplam", 100); // 6
        }



        private void btnGuncelle_Click(object sender, EventArgs e)
        {
            if (lvRezervasyonlar.SelectedItems.Count == 0)
            {
                MessageBox.Show("Lütfen bir rezervasyon seçiniz!");
                return;
            }

            ListViewItem secilen = lvRezervasyonlar.SelectedItems[0];

            // -------------------------------------------------------------
            // 🔥 YENİ KONTROL: TARİH VE SAAT BİRLİKTE KONTROL EDİLİYOR
            // -------------------------------------------------------------
            string tarihStr = secilen.SubItems[3].Text; // Örn: 16.12.2025
            string saatStr = secilen.SubItems[4].Text;  // Örn: 10:00

            // Tarih ve saati birleştirip tam zamanı buluyoruz
            DateTime rezervasyonZamani = Convert.ToDateTime(tarihStr + " " + saatStr);

            // Şimdiki zamanla kıyaslıyoruz (Tarih aynı olsa bile saat geçmişse hata verir)
            if (rezervasyonZamani < DateTime.Now)
            {
                MessageBox.Show("Saati veya tarihi geçmiş bir rezervasyonu güncelleyemezsiniz!",
                                "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // İşlemi durdur
            }
            // -------------------------------------------------------------


            // ID'yi al
            GuncellenenRezervasyonID = Convert.ToInt32(secilen.SubItems[0].Text);

            // Verileri kutulara doldur
            cmbSporTuru.Text = secilen.SubItems[1].Text;
            cmbSaha.Text = secilen.SubItems[2].Text;
            dtpTarih.Value = Convert.ToDateTime(tarihStr); // Tarih kontrolü geçtiği için artık hata vermez
            cmbSaat.Text = saatStr;

            // Eğitmen İşlemleri
            string egitmenAdi = secilen.SubItems[5].Text;

            if (egitmenAdi != "-")
            {
                chkEgitmen.Checked = true;

                using (var con = new SQLiteConnection(dbPath))
                {
                    con.Open();
                    using (var cmd = new SQLiteCommand(
                        "SELECT ID, PricePerHour FROM Trainers WHERE NameSurname=@n", con))
                    {
                        cmd.Parameters.AddWithValue("@n", egitmenAdi);
                        using (var dr = cmd.ExecuteReader())
                        {
                            if (dr.Read())
                            {
                                SecilenEgitmenID = Convert.ToInt32(dr["ID"]);
                                egitmenUcret = Convert.ToDouble(dr["PricePerHour"]);
                            }
                        }
                    }
                }
            }
            else
            {
                chkEgitmen.Checked = false;
                SecilenEgitmenID = 0;
                egitmenUcret = 0;
            }

            // Saha Ücreti Hesapla
            using (var con = new SQLiteConnection(dbPath))
            {
                con.Open();
                using (var cmd = new SQLiteCommand(
                    "SELECT PricePerHour FROM Fields WHERE Name=@n", con))
                {
                    cmd.Parameters.AddWithValue("@n", cmbSaha.Text);
                    sahaUcret = Convert.ToDouble(cmd.ExecuteScalar());
                }
            }

            eskiOdenenTutar = Convert.ToDouble(secilen.SubItems[6].Text.Replace(" ₺", ""));
            GuncellemeModu = true;

            Hesapla();
        }

        private void btnIptalg_Click(object sender, EventArgs e)
        {
            if (lvRezervasyonlar.SelectedItems.Count == 0)
            {
                MessageBox.Show("Lütfen iptal edilecek rezervasyonu seçin.");
                return;
            }

            ListViewItem secilen = lvRezervasyonlar.SelectedItems[0];

            // -------------------------------------------------------------
            // 🔥 YENİ KONTROL: TARİH VE SAAT BİRLİKTE KONTROL EDİLİYOR
            // -------------------------------------------------------------
            string tarihStr = secilen.SubItems[3].Text;
            string saatStr = secilen.SubItems[4].Text;

            DateTime rezervasyonZamani = Convert.ToDateTime(tarihStr + " " + saatStr);

            if (rezervasyonZamani < DateTime.Now)
            {
                MessageBox.Show("Saati veya tarihi geçmiş bir rezervasyon iptal edilemez!",
                                "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            // -------------------------------------------------------------

            if (MessageBox.Show("Rezervasyon iptal edilsin mi?",
                "İptal", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            int id = Convert.ToInt32(secilen.SubItems[0].Text);

            using (var con = new SQLiteConnection(dbPath))
            {
                con.Open();
                using (var cmd = new SQLiteCommand(
                    "DELETE FROM Bookings WHERE ID=@id", con))
                {
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }
            }

            MessageBox.Show("İptal edildi.\nEn yakın zamanda kartınıza iade edilecektir.");
            RezervasyonlariGetir();
        }

        private void RezervasyonuGuncelle(double yeniToplam)
        {
            using (var con = new SQLiteConnection(dbPath))
            {
                con.Open();

                using (var cmd = new SQLiteCommand(@"
        UPDATE Bookings 
        SET 
            SportID = @sport,
            FieldID = @field,
            Date = @date,
            Time = @time,
            TotalPrice = @price,
            TrainerID = @trainer
        WHERE ID = @id AND UserID = @userId", con))
                {
                    cmd.Parameters.AddWithValue("@userId", UserSession.CurrentUserId);
                    cmd.Parameters.AddWithValue("@sport", cmbSporTuru.SelectedValue);
                    cmd.Parameters.AddWithValue("@field", cmbSaha.SelectedValue);
                    cmd.Parameters.AddWithValue("@date", dtpTarih.Value.ToShortDateString());
                    cmd.Parameters.AddWithValue("@time", cmbSaat.Text);
                    cmd.Parameters.AddWithValue("@price", yeniToplam);
                    cmd.Parameters.AddWithValue("@trainer",
                        SecilenEgitmenID == 0 ? (object)DBNull.Value : SecilenEgitmenID);
                    cmd.Parameters.AddWithValue("@id", GuncellenenRezervasyonID);

                    cmd.ExecuteNonQuery();
                }
            }
        }

       

        private void btnGenelGuncelle_Click(object sender, EventArgs e)
        {
            // --- YENİ EKLENEN KISIM: GİRİŞ KONTROLÜ ---
            if (UserSession.CurrentUserId == 0) // Kullanıcı ID'si 0 ise giriş yapmamış demektir.
            {
                MessageBox.Show("Rezervasyon güncellemek için lütfen önce üye girişi yapınız.",
                                "Yetkisiz İşlem", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                // İstersen burada giriş formunu otomatik açtırabilirsin:
                // FrmGiris giris = new FrmGiris();
                // giris.ShowDialog();

                return; // İşlemi durdur, aşağıya inme.
            }
            // ------------------------------------------

            double guncelToplam = sahaUcret + egitmenUcret;

            // 🔴 FARK VAR → ÖDEME SAYFASINA GİT
            if (guncelToplam > eskiOdenenTutar)
            {
                farkTutari = guncelToplam - eskiOdenenTutar;

                TemizleOdemeAlanlari();
                OdemeSayfasiBilgileriDoldur();

                // 🔥 EN SONA KOY
                lblFarkText.Visible = true;
                lblFarkValue.Visible = true;
                lblFarkValue.Text = farkTutari + " ₺";

                tabControl1.SelectedTab = tabOdeme;
                return;
            }

            // 🟢 İADE VAR VEYA FİYAT AYNI → DİREKT GÜNCELLE
            if (guncelToplam < eskiOdenenTutar)
            {
                RezervasyonuGuncelle(guncelToplam);
                MessageBox.Show("Güncellendi.\nFazla tutar kartınıza iade edilecektir.");
            }
            else
            {
                RezervasyonuGuncelle(guncelToplam);
                MessageBox.Show("Rezervasyon güncellendi.");
            }

            GuncellemeModu = false;
            RezervasyonlariGetir(); // Listeyi yenile
        }
        private void TemizleOdemeAlanlari()
        {
            txtKartSahibi.Clear();
            mskKartNo.Clear();
            mskSKT.Clear();
            mskCVV.Clear();

            lblFarkText.Visible = false;
            lblFarkValue.Visible = false;

        }

        private void tabKayit_Click(object sender, EventArgs e)
        {

        }

        private void cmbTaksit_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tabIstatistik_Click(object sender, EventArgs e)
        {

        }
    }
}


