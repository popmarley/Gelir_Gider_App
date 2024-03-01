using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace Gelir_Gider_App
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
            // ComboBox'a Ay Listesi Ekleyin
            for (int i = 1; i <= 12; i++)
            {
                cmbAy.Items.Add(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i));
            }
        }

        private void uygulamaHakkindaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Hakkinda hakkinda = new Hakkinda();
            hakkinda.ShowDialog();
        }

        private void btnKaydet_Click(object sender, EventArgs e)
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Gelir Gider App.xlsx");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Geliştirme amacıyla kullanım için lisans matrisini ayarladı
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Ay isimlerini sırayla al ve kontrol et
                var months = Enumerable.Range(1, 12).Select(i => System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i).ToUpper()).ToList();

                // Excel dosyasında her bir ay için bir sayfa oluştur
                foreach (var month in months)
                {
                    // Eğer sayfa zaten varsa atla
                    if (package.Workbook.Worksheets.Any(x => x.Name.ToUpper() == month))
                        continue;

                    // Sayfa yoksa, yeni bir sayfa ekle
                    package.Workbook.Worksheets.Add(month);
                }

                string selectedMonth = cmbAy.SelectedItem.ToString().ToUpper();
                var worksheet = package.Workbook.Worksheets[selectedMonth] ?? package.Workbook.Worksheets.Add(selectedMonth);

                // Başlıkları Ekle
                worksheet.Cells["A1"].Value = "Harcama Adı";
                worksheet.Cells["B1"].Value = "Tarih";
                worksheet.Cells["C1"].Value = "Tutar";
                worksheet.Cells["D1"].Value = "İşlem";

                // Sadece A-D sütunları için son dolu satırı bul
                int lastRowForABCD = 1;
                for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
                {
                    if (worksheet.Cells[i, 1].Value != null || worksheet.Cells[i, 2].Value != null || worksheet.Cells[i, 3].Value != null || worksheet.Cells[i, 4].Value != null)
                    {
                        lastRowForABCD = i;
                    }
                }

                int row = lastRowForABCD + 1; // Yeni veri eklenecek satır

                // Verileri Ekle
                worksheet.Cells[$"A{row}"].Value = txtHarcama.Text;
                worksheet.Cells[$"B{row}"].Value = dtTarih.Value.ToString("dd/MM/yyyy");
                worksheet.Cells[$"C{row}"].Value = decimal.Parse(txtTutar.Text);
                // Tutar + veya -'ye göre renk ve metin ayarlama
                if (decimal.Parse(txtTutar.Text) < 0)
                {
                    worksheet.Cells[$"D{row}"].Value = "Para Çıkış";
                    worksheet.Cells[$"D{row}"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                }
                else
                {
                    worksheet.Cells[$"D{row}"].Value = "Para Giriş";
                    worksheet.Cells[$"D{row}"].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                }

                // Toplam Formülleri Ekle
                // Bu kısımlar sabit kalır
                worksheet.Cells["F7"].Value = "Gelir";
                worksheet.Cells["F8"].Value = "Gider";
                worksheet.Cells["F9"].Value = "Kalan";
                worksheet.Cells["G7"].Formula = "SUMIF(C:C,\">0\")";
                worksheet.Cells["G8"].Formula = "SUMIF(C:C,\"<0\")";
                worksheet.Cells["G9"].Formula = "G7+G8"; // Daha doğru formülleme
                worksheet.Cells["G7"].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                worksheet.Cells["G8"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                worksheet.Cells["G9"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);

                // Değişiklikleri kaydet
                package.Save();
            }

            MessageBox.Show("Kayıt Başarılı!");
            txtHarcama.Clear();
            txtTutar.Clear();
        }


    }
}

