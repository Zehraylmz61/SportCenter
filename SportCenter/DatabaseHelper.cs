using System;
using System.Data;
using System.Data.SQLite;
using System.Windows.Forms;

namespace SportCenter
{
    public class DatabaseHelper
    {
        // GÜNCELLEME: Veritabanı adı SportCenter.db oldu
        // Veritabanı dosyasının tam adresini (Mutlak Yol) veriyoruz:
        private string connectionString = @"Data Source=C:\Users\zeroo\Desktop\vss-istatistik(düzeltme olucak)\vss-istatistik(düzeltme olucak)\vss\SportCenter\SporCenter.db";






        // 1. Ekleme / Silme / Güncelleme
        public bool ExecuteQuery(string query, SQLiteParameter[] p = null)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                    {
                        if (p != null) cmd.Parameters.AddRange(p);
                        cmd.ExecuteNonQuery();
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("İşlem Hatası: " + ex.Message);
                    return false;
                }
            }
        }

        // 2. Veri Çekme (Tablo Olarak - ComboBox vb. için)
        public DataTable GetData(string query, SQLiteParameter[] p = null)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                    {
                        if (p != null) cmd.Parameters.AddRange(p);
                        using (SQLiteDataAdapter da = new SQLiteDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            da.Fill(dt);
                            return dt;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Veri Çekme Hatası: " + ex.Message);
                    return null;
                }
            }
        }

        // 3. Tek Değer Çekme (ID veya Sayı alma)
        public object GetScalar(string query, SQLiteParameter[] p = null)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                try
                {
                    conn.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
                    {
                        if (p != null) cmd.Parameters.AddRange(p);
                        return cmd.ExecuteScalar();
                    }
                }
                catch { return null; }
            }
        }
    }
}