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

            // Excel Dosyasını Kontrol Et ve Varsa Ayları Oluştur
            CreateOrInitializeExcelFile();

            // Mevcut ayı otomatik olarak seç
            SelectCurrentMonth();


        }

        private void SelectCurrentMonth()
        {
            // Mevcut ayın adını al
            string currentMonthName = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(DateTime.Now.Month);

            // ComboBox'ta mevcut aya karşılık gelen değeri seç
            cmbAy.SelectedIndex = cmbAy.Items.IndexOf(currentMonthName);
        }


        private void CreateOrInitializeExcelFile()
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Gelir Gider App.xlsx");
            FileInfo fileInfo = new FileInfo(filePath);

            // Dosya yoksa veya boşsa yeni bir dosya oluştur
            if (!fileInfo.Exists || fileInfo.Length == 0)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(fileInfo))
                {
                    // Ay isimlerini sırayla al ve kontrol et
                    var months = Enumerable.Range(1, 12).Select(i => System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i).ToUpper()).ToList();

                    // Excel dosyasında her bir ay için bir sayfa oluştur, başlıkları ve toplam formülleri ekle
                    foreach (var month in months)
                    {
                        var worksheet = package.Workbook.Worksheets.Add(month);

                        // Başlıkları Ekle
                        worksheet.Cells["A1"].Value = "Harcama Adı";
                        worksheet.Cells["B1"].Value = "Tarih";
                        worksheet.Cells["C1"].Value = "Tutar";
                        worksheet.Cells["D1"].Value = "İşlem";

                        // Gelir, Gider ve Kalan için bölümleri ve formülleri ekle
                        worksheet.Cells["F7"].Value = "Gelir";
                        worksheet.Cells["F8"].Value = "Gider";
                        worksheet.Cells["F9"].Value = "Kalan";
                        worksheet.Cells["G7"].Formula = "SUMIF(C:C,\">0\")";
                        worksheet.Cells["G8"].Formula = "SUMIF(C:C,\"<0\")";
                        worksheet.Cells["G9"].Formula = "G7+G8"; // Daha doğru formülleme
                        worksheet.Cells["G7"].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                        worksheet.Cells["G8"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                        worksheet.Cells["G9"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                    }

                    // Değişiklikleri kaydet
                    package.Save();
                }
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
                // Kaydetme işlemi tamamlandıktan sonra, mevcut 'selectedMonth' değişkenini kullan
                LoadExcelDataToDataGridView(selectedMonth);
                UpdateLabelsFromExcel(selectedMonth);
            }

            MessageBox.Show("Kayıt Başarılı!");
           
            txtHarcama.Clear();
            txtTutar.Clear();
           

        }


        public void LoadExcelDataToDataGridView(string selectedMonth)
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Gelir Gider App.xlsx");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Geliştirme amacıyla kullanım için lisans matrisini ayarladı
            DataTable dataTable = new DataTable();

            using(var package = new ExcelPackage(new FileInfo(filePath)))
    {
                var worksheet = package.Workbook.Worksheets[selectedMonth];
                if (worksheet == null) return; // Seçilen ay için bir sayfa yoksa dön

                // Sütun başlıklarını ve sınırlı sütunları DataTable'a ekle
                dataTable.Columns.Add("Harcama Adı");
                dataTable.Columns.Add("Tarih");
                dataTable.Columns.Add("Tutar");
                dataTable.Columns.Add("İşlem");

                // A2:D2 aralığındaki verileri DataTable'a aktar
                for (int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                {
                    var row = dataTable.NewRow();

                    // Sadece A, B, C, D sütunlarındaki verileri al
                    row["Harcama Adı"] = worksheet.Cells[rowNum, 1].Text;
                    row["Tarih"] = worksheet.Cells[rowNum, 2].Text;
                    row["Tutar"] = worksheet.Cells[rowNum, 3].Text;
                    row["İşlem"] = worksheet.Cells[rowNum, 4].Text;

                    dataTable.Rows.Add(row);
                }
            }

            // DataGridView'i DataTable ile doldur
            dataGridView1.DataSource = dataTable;
            
        }

        private void UpdateLabelsFromExcel(string selectedMonth)
        {
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Gelir Gider App.xlsx");
            FileInfo fileInfo = new FileInfo(filePath);

            if (fileInfo.Exists)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(fileInfo))
                {
                    var worksheet = package.Workbook.Worksheets[selectedMonth];

                    if (worksheet != null)
                    {
                        // Hesaplamaları zorla
                        worksheet.Calculate();
                        // Excel'den değerleri oku
                        var gelirValue = worksheet.Cells["G7"].Value?.ToString() ?? "0";
                        var giderValue = worksheet.Cells["G8"].Value?.ToString() ?? "0";
                        var kalanValue = worksheet.Cells["G9"].Value?.ToString() ?? "0";

                        // UI thread üzerinde label'ları güncelle
                        Action updateUI = () =>
                        {
                            lblGelir.Text = gelirValue;
                            lblGider.Text = giderValue;
                            lblKalan.Text = kalanValue;

                            lblGelir.ForeColor = System.Drawing.Color.Green;
                            lblGider.ForeColor = System.Drawing.Color.Red;
                            lblKalan.ForeColor = System.Drawing.Color.Blue;
                        };

                        // UI thread kontrolü
                        if (this.InvokeRequired)
                        {
                            this.Invoke(updateUI);
                        }
                        else
                        {
                            updateUI();
                        }
                    }
                }
            }
        }

        private void cmbAy_SelectedIndexChanged(object sender, EventArgs e)
        {

            // Seçilen ayın adını al (ComboBox'dan)
            string selectedMonth = cmbAy.SelectedItem.ToString();
            LoadExcelDataToDataGridView(selectedMonth.ToUpper());

            // Seçili aya göre Excel'den verileri oku ve label'lara ata
            UpdateLabelsFromExcel(selectedMonth.ToUpper());
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // İşlem sütununun indeksini kontrol et (örneğin 3 varsayalım)
            if (dataGridView1.Columns[e.ColumnIndex].Name == "İşlem" && e.Value != null)
            {
                // Hücre değerine göre renk ataması yap
                if (e.Value.ToString() == "Para Giriş")
                {
                    e.CellStyle.ForeColor = Color.Green; // Metin rengi
                                                         // Alternatif olarak, arka plan rengini de ayarlayabilirsiniz:
                                                         // e.CellStyle.BackColor = Color.LightGreen;
                }
                else if (e.Value.ToString() == "Para Çıkış")
                {
                    e.CellStyle.ForeColor = Color.Red; // Metin rengi
                                                       // Alternatif olarak, arka plan rengini de ayarlayabilirsiniz:
                                                       // e.CellStyle.BackColor = Color.LightCoral;
                }
            }
        }
    }
}

