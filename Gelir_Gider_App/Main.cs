using OfficeOpenXml;
using OfficeOpenXml.Style;
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

            TabloSilGuncelle();

            CustomizeDataGridViewStyle();
        }

        private void TabloSilGuncelle()
        {
            // Güncelle butonu sütunu
            DataGridViewButtonColumn btnGuncelle = new DataGridViewButtonColumn();
            btnGuncelle.HeaderText = "";
            btnGuncelle.Text = "Güncelle";
            btnGuncelle.Name = "btnGuncelle";
            btnGuncelle.UseColumnTextForButtonValue = true; // Butonun üzerindeki yazıyı ayarlar
            dataGridView1.Columns.Add(btnGuncelle);

            // Sil butonu sütunu
            DataGridViewButtonColumn btnSil = new DataGridViewButtonColumn();
            btnSil.HeaderText = "";
            btnSil.Text = "Sil";
            btnSil.Name = "btnSil";
            btnSil.UseColumnTextForButtonValue = true; // Butonun üzerindeki yazıyı ayarlar
            dataGridView1.Columns.Add(btnSil);
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

                        // Başlıkları Ekle ve Stillerini Ayarla
                        var headerRange = "A1:D1"; // Başlıkların olduğu hücre aralığı
                        worksheet.Cells[headerRange].LoadFromArrays(new string[][] {
                        new string[] { "Harcama Adı", "Tarih", "Tutar", "İşlem" }
                        });

                        // Başlıklar için stili ayarla (kalın ve ortalı)
                        worksheet.Cells[headerRange].Style.Font.Bold = true; // Başlık fontunu kalın yap
                        worksheet.Cells[headerRange].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Başlıkları ortalı yap

                        // Opsiyonel: Başlık arka plan rengini ayarla
                        //worksheet.Cells[headerRange].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        //worksheet.Cells[headerRange].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);


                        // Gelir, Gider ve Kalan için bölümleri ve formülleri ekle
                        worksheet.Cells["F7"].Value = "Gelen Para";
                        worksheet.Cells["F8"].Value = "Giden Para";
                        worksheet.Cells["F9"].Value = "Durum(K/Z)";
                        worksheet.Cells["G7"].Formula = "SUMIF(C:C,\">0\")";
                        worksheet.Cells["G8"].Formula = "SUMIF(C:C,\"<0\")";
                        worksheet.Cells["G9"].Formula = "G7+G8"; // Daha doğru formülleme

                        // Gelir, Gider ve Kalan hücrelerinin formatını ayarla, yanında TL simgesi ile
                        worksheet.Cells["G7"].Style.Numberformat.Format = "#,##0 \"₺\"";
                        worksheet.Cells["G8"].Style.Numberformat.Format = "#,##0 \"₺\"";
                        worksheet.Cells["G9"].Style.Numberformat.Format = "#,##0 \"₺\"";

                        worksheet.Cells["G7"].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                        worksheet.Cells["G8"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                        worksheet.Cells["G9"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);

                        // Başlıklar için stili ayarla (kalın ve ortalı)
                        worksheet.Cells["F7"].Style.Font.Bold = true; // Başlık fontunu kalın yap
                        worksheet.Cells["F7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left; // Başlıkları sola yasla 
                        worksheet.Cells["F8"].Style.Font.Bold = true; // Başlık fontunu kalın yap
                        worksheet.Cells["F8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left; // Başlıkları sola yasla 
                        worksheet.Cells["F9"].Style.Font.Bold = true; // Başlık fontunu kalın yap
                        worksheet.Cells["F9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left; // Başlıkları sola yasla 
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

                // Başlıkları Ekle ve Stillerini Ayarla
                var headerRange = "A1:D1"; // Başlıkların olduğu hücre aralığı
                worksheet.Cells[headerRange].LoadFromArrays(new string[][] {
                new string[] { "Harcama Adı", "Tarih", "Tutar", "İşlem" }
                });

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
                // Tutarı sayısal olarak ekle ve formatını ayarla
                worksheet.Cells[$"C{row}"].Value = decimal.Parse(txtTutar.Text);
                worksheet.Cells[$"C{row}"].Style.Numberformat.Format = "#,##0 \"₺\""; // Binlik ayırıcı ekler

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

            using (var package = new ExcelPackage(new FileInfo(filePath)))
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
                        // Excel'den değerleri oku ve " TL" ekleyerek formatla
                        var gelirValue = double.TryParse(worksheet.Cells["G7"].Value?.ToString(), out double gelir) ? $"{gelir.ToString("#,##0")} ₺" : "0 ₺";
                        var giderValue = double.TryParse(worksheet.Cells["G8"].Value?.ToString(), out double gider) ? $"{gider.ToString("#,##0")} ₺" : "0 ₺";
                        var kalanValue = double.TryParse(worksheet.Cells["G9"].Value?.ToString(), out double kalan) ? $"{kalan.ToString("#,##0")} ₺" : "0 ₺";


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

        public void CustomizeDataGridViewStyle()
        {
            // Sütun başlıklarının stilini ayarla
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font(dataGridView1.Font, FontStyle.Bold);
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            // Opsiyonel: Sütun başlıklarının arka plan ve yazı rengini ayarla
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.EnableHeadersVisualStyles = false; // Özel stilin uygulanmasını sağlar
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // "Sil" butonunun sütun indeksini kontrol et
            if (dataGridView1.Columns[e.ColumnIndex].Name == "btnSil" && e.RowIndex >= 0)
            {
                // Kullanıcıdan silme işlemini onaylamasını iste
                var confirmResult = MessageBox.Show("Bu kaydı silmek istediğinize emin misiniz?", "Kaydı Sil", MessageBoxButtons.YesNo);
                if (confirmResult == DialogResult.Yes)
                {
                    // Excel dosyasını aç
                    string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Gelir Gider App.xlsx");
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        string selectedMonth = cmbAy.SelectedItem.ToString().ToUpper();
                        var worksheet = package.Workbook.Worksheets[selectedMonth];
                        if (worksheet != null)
                        {
                            // dataGridView1'den seçilen satırın indeksine karşılık gelen Excel satırını bul (Excel'de satır indeksleri 1'den başlar)
                            int excelRowIndex = e.RowIndex + 2; // dataGridView'deki indeks 0'dan başladığı için ve başlık satırı olduğu için +2
                            worksheet.DeleteRow(excelRowIndex);

                            // Excel dosyasındaki değişiklikleri kaydet
                            package.Save();
                        }
                    }

                    // DataGridView'den seçilen satırı sil
                    dataGridView1.Rows.RemoveAt(e.RowIndex);


                    // Gerekirse burada DataGridView'i yeniden yükleyin ve güncelleyin
                    //LoadExcelDataToDataGridView(selectedMonth);
                }
            }
        }
    }
}

