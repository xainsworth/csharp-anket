using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using Newtonsoft.Json;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Header;

namespace Anket
{

    public partial class Form1 : Form
    {
        public List<Kullanici> kullanicilar = new List<Kullanici>();
        public List<Soru> sorular = new List<Soru>();
        public List<Cevap> cevaplar = new List<Cevap>();
        public List<RadioButton> radioButtonlar = new List<RadioButton>();
        public List<CheckBox> checkBoxlar = new List<CheckBox>();
        public List<Soru> yenianket = new List<Soru>();
        public ComboBox combo = new ComboBox();
        public Soru yenisoru = new Soru();
        public ListBox lb = new ListBox();
        int id, soruid, yeniAnket_id;
        public Cevap dogru_cevap;
        string adi, soyadi, kullanici_adi, sifre;
        public string yol = Path.Combine(Path.GetDirectoryName(path: System.Reflection.Assembly.GetExecutingAssembly().Location), path2: "kullanicilar.json");
        public string yol2 = Path.Combine(Path.GetDirectoryName(path: System.Reflection.Assembly.GetExecutingAssembly().Location), path2: "sorular.json");
        public string yol3 = Path.Combine(Path.GetDirectoryName(path: System.Reflection.Assembly.GetExecutingAssembly().Location), path2: "cevaplar.json");
        public Kullanici giris;
        public Cevap verilen_cevaplar = new Cevap();
        bool yaptiMi = false;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            panel2.Hide();
            panel3.Hide();
            panel4.Hide();
            panel5.Hide();
            Size = panel1.Size;
            string json_yazi = File.ReadAllText(yol);
            kullanicilar = JsonConvert.DeserializeObject<List<Kullanici>>(json_yazi);
            string json_yazi2 = File.ReadAllText(yol2);
            sorular = JsonConvert.DeserializeObject<List<Soru>>(json_yazi2);
            string json_yazi3 = File.ReadAllText(yol3);
            cevaplar = JsonConvert.DeserializeObject<List<Cevap>>(json_yazi3);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            giris = kullanicilar.Find(x => x.kullanici_adi == textBox1.Text && x.sifre == textBox2.Text);
            if (giris != null)
            {
                panel1.Hide();
                textBox1.Clear();
                textBox2.Clear();
                panel3.Location = new Point(0, 0);
                panel3.Show();
                Size = panel3.Size;
                label12.Text = giris.adi + " " + giris.soyadi;
                button6.Visible = giris.admin;
                soru_Goster(sorular[0], label10, label11, groupBox1.Controls);
                if (cevaplar != null && cevaplar.Exists(x => x.user_id == giris.id))
                {
                    yaptiMi = true;
                    dogru_cevap = cevaplar.Find(x => x.user_id == giris.id);
                    
                    
                    if (sorular[soruid].multiple == 0 && dogru_cevap.cevaplar[0] != "-1")
                    {
                        radioButtonlar[Convert.ToInt32(dogru_cevap.cevaplar[0])].Checked = true;
                    }
                    if (sorular[soruid].multiple == 1)
                    {
                        for (int i = 0; i < dogru_cevap.cevaplar[0].Split(';').Length; i++)
                        {
                            if (dogru_cevap.cevaplar[0].Split(";")[i] == "1")
                            {
                                checkBoxlar[i].Checked = true;
                            }
                        }
                    }

                    if (sorular[soruid].multiple == 2 && dogru_cevap.cevaplar[0] != "-1")
                    {
                        combo.SelectedIndex = Convert.ToInt32(dogru_cevap.cevaplar[0]);
                    }
                    if (sorular[soruid].multiple == 3 && dogru_cevap.cevaplar[0] != "-1")
                    {
                        lb.SelectedValue = Convert.ToInt32(dogru_cevap.cevaplar[0]);
                    }
                }
                verilen_cevaplar.id = cevaplar.Count > 0 ? cevaplar.Max(x => x.id) + 1 : 0;
                verilen_cevaplar.user_id = giris.id;
                verilen_cevaplar.cevaplar = new List<string>( new string[sorular.Count] );
                soruid++;
            }
            else
            {
                MessageBox.Show("Kullanıcı adı veya şifre yanlış.", "Giriş Başarısız", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            panel1.Hide();
            textBox1.Clear(); textBox2.Clear();
            panel2.Location = new Point(0, 0);
            panel2.Show();
            Size = panel2.Size;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (radioButtonlar.Count > 0)
            {
                verilen_cevaplar.cevaplar[soruid - 1] = "-1";
                for (int i = 0; i < radioButtonlar.Count; i++)
                {
                    if (radioButtonlar[i].Checked) verilen_cevaplar.cevaplar[soruid - 1] = i.ToString();
                }
            }
            else if (checkBoxlar.Count > 0)
            {
                verilen_cevaplar.cevaplar[soruid - 1] = "";
                for (int i = 0; i < checkBoxlar.Count; i++)
                {
                    if (checkBoxlar[i].Checked) verilen_cevaplar.cevaplar[soruid - 1] += "1";
                    else { verilen_cevaplar.cevaplar[soruid - 1] += "-1"; }
                    if (i != checkBoxlar.Count - 1) verilen_cevaplar.cevaplar[soruid - 1] += ";";
                }
            }
            else if (combo.Items.Count > 0)
            {
                if (combo.SelectedItem != null) verilen_cevaplar.cevaplar[soruid - 1] = (combo.SelectedIndex).ToString();
                else verilen_cevaplar.cevaplar[soruid - 1] = "-1";
            }
            else if (lb.Items.Count > 0)
            {
                if (lb.SelectedItem != null) verilen_cevaplar.cevaplar[soruid - 1] = (lb.SelectedIndex).ToString();
                else verilen_cevaplar.cevaplar[soruid - 1] = "-1";
            }
            int max_id = sorular.Max(x => x.id);
            if (soruid <= max_id)
            {
                soru_Goster(sorular[soruid], label10, label11, groupBox1.Controls);
                if (dogru_cevap != null)
                {
                    if (sorular[soruid].multiple == 0 && dogru_cevap.cevaplar[soruid] != "-1")
                    {
                        radioButtonlar[Convert.ToInt32(dogru_cevap.cevaplar[soruid])].Checked = true;
                    }
                    if (sorular[soruid].multiple == 1)
                    {
                        for (int i = 0; i < dogru_cevap.cevaplar[soruid].Split(';').Length; i++)
                        {
                            if (dogru_cevap.cevaplar[soruid].Split(";")[i] == "1")
                            {
                                checkBoxlar[i].Checked = true;
                            }
                        }
                    }

                    if (sorular[soruid].multiple == 2 && dogru_cevap.cevaplar[soruid] != "-1")
                    {
                        combo.SelectedIndex = Convert.ToInt32(dogru_cevap.cevaplar[soruid]);
                    }
                    if (sorular[soruid].multiple == 3 && dogru_cevap.cevaplar[soruid] != "-1")
                    {
                        lb.SelectedIndex = Convert.ToInt32(dogru_cevap.cevaplar[soruid]);
                    }
                }
                
                soruid++;
            }
            else {
                if (soruid == max_id+1)
                {
                    radioButtonlar.Clear();
                    checkBoxlar.Clear();
                    combo.Items.Clear();
                    lb.Items.Clear();
                    groupBox1.Hide();
                    label10.Text = "Test Bitti.";
                    label11.Text = "Çıkmak için alttaki butona basın.";
                    button5.Text = "Çıkıþ";
                    button5.Location = new Point(32, 200);
                    panel3.Size = new Size(516, 300);
                    Size = new Size(516, 300);
                    soruid++;
                }
                else
                {
                    if (yaptiMi)
                    {
                        cevaplar[cevaplar.FindIndex(x => x.user_id == giris.id)] = verilen_cevaplar;
                    }
                    else cevaplar.Add(verilen_cevaplar);
                    string jsobj = JsonConvert.SerializeObject(cevaplar);
                    File.WriteAllText(yol3, jsobj);
                    Environment.Exit(0);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            panel2.Hide();
            textBox3.Clear(); textBox4.Clear(); textBox5.Clear(); textBox6.Clear();
            Size = panel1.Size;
            panel1.Show();
        }


        
        private void button3_Click(object sender, EventArgs e)
        {
            id = kullanicilar.Count > 0 ? (kullanicilar.Max(x => x.id))+1 : 0;
            adi = textBox6.Text; soyadi = textBox5.Text; kullanici_adi = textBox3.Text; sifre = textBox4.Text;
            if (!kullanicilar.Exists(x => x.kullanici_adi == kullanici_adi))
            {
                var yenikullanici = new Kullanici();
                yenikullanici.id = id; yenikullanici.adi = adi; yenikullanici.admin = false;
                yenikullanici.soyadi = soyadi; yenikullanici.kullanici_adi = kullanici_adi;
                yenikullanici.sifre = sifre;
                kullanicilar.Add(yenikullanici);
                string jsonobj = JsonConvert.SerializeObject(kullanicilar);
                File.WriteAllText(yol, jsonobj);
                MessageBox.Show("Kullanıcı başarıyla oluşturuldu. Giriş ekranına gidiliyor.", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
                panel2.Hide();
                Size = panel1.Size;
                panel1.Show();

            }
            else
            {
                MessageBox.Show("Bu kullanıcı adı mevcut. Lütfen başka kullanıcı adı seçin.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            panel3.Hide();
            panel4.Location = new Point(0, 0);
            panel4.Show();
            Size = panel4.Size;
            label14.Text = cevaplar.Count.ToString();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            panel4.Hide();
            Size = panel3.Size;
            panel3.Show();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var workbook = new XLWorkbook();
            workbook.AddWorksheet("Anket");
            var ws = workbook.Worksheet("Anket");
            int x=1, y=1;
            ws.Row(1).Height = 40;
            ws.Row(1).Style.Alignment.WrapText = true;
            ws.Row(1).Style.Font.Bold = true;
            ws.Cell(x, y).Style.Fill.BackgroundColor = XLColor.FromArgb(0xffd7d7);
            ws.Column(y).Width = 21;
            ws.Cell(x, y).Value = "İsim";
            for (int i = 0; i < sorular.Count; i++)
            {
                ws.Cell(x, y + i + 1).Style.Fill.BackgroundColor = XLColor.FromArgb(0xffd7d7);
                ws.Column(y + i + 1).Width = 21;
                ws.Cell(x, y + i + 1).Value = sorular[i].soru;
            }
            x++;
            foreach (Cevap item in cevaplar)
            {
                Kullanici p = kullanicilar.Find(i => i.id == item.user_id) ?? new Kullanici();
                ws.Cell(x,y).Value = p.adi + " " + p.soyadi;
                for (int i = 0; i < item.cevaplar.Count; i++)
                {
                    if (item.cevaplar[i] != "-1")
                    {
                        if (!item.cevaplar[i].Contains(';'))
                        {
                            ws.Cell(x, y + i + 1).Value = sorular[i].cevaplar[int.Parse(item.cevaplar[i])];
                        }
                        else
                        {
                            string[] cbcevaplar = item.cevaplar[i].Split(';');
                            for (int j = 0; j < cbcevaplar.Length; j++)
                            {
                                if (cbcevaplar[j] == "1")
                                    ws.Cell(x, y + i + 1).Value += sorular[i].cevaplar[j] + ", ";
                            }
                        }
                    }
                }
                x++;
            }

            workbook.SaveAs("anket.xlsx");
            MessageBox.Show("Excel dosyası başarıyla oluþturuldu.", "Başarılı", MessageBoxButtons.OK);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            panel4.Hide();
            panel5.Location = new Point(0, 0);
            Size = panel5.Size;
            comboBox1.SelectedIndex = 0;
            yenianket.Clear();
            label15.Text = "Soru 1";
            yeniAnket_id = 0;
            textBox7.Clear();
            textBox8.Clear();
            listBox1.Items.Clear();
            panel5.Show();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            listBox1.Items.Add(textBox8.Text);
            textBox8.Clear();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            listBox1.Items.Remove(listBox1.SelectedItem);
        }

        private void button13_Click(object sender, EventArgs e)
        {
            yenisoru.id = yeniAnket_id;
            yenisoru.multiple = comboBox1.SelectedIndex;
            yenisoru.soru = textBox7.Text;
            yenisoru.cevaplar = new List<string>();
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                yenisoru.cevaplar.Add(listBox1.Items[i]+"");
            }
            yenianket.Add(yenisoru);
            yenisoru = new Soru();
            comboBox1.SelectedIndex = 0;
            yeniAnket_id++;
            textBox7.Clear();
            textBox8.Clear();
            label15.Text = "Soru " + (yeniAnket_id + 1);
            listBox1.Items.Clear();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            File.Move(yol2, Path.Combine(Path.GetDirectoryName(path: System.Reflection.Assembly.GetExecutingAssembly().Location), path2: "sorular.old.json"));
            File.Move(yol3, Path.Combine(Path.GetDirectoryName(path: System.Reflection.Assembly.GetExecutingAssembly().Location), path2: "cevaplar.old.json"));
            string yeniAnket_json = JsonConvert.SerializeObject(yenianket);
            File.WriteAllText(yol2, yeniAnket_json);
            File.WriteAllText(yol3, "[]");
            MessageBox.Show("Yeni anketiniz başarıyla oluşturuldu. Eskisinin yedeği alındı. Programdan çıkış yapılıyor.", "Başarılı", MessageBoxButtons.OK);
            Environment.Exit(0);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            panel5.Hide();
            panel4.Show();
        }

        public class Kullanici
        {
            public int id { get; set; }

            public bool admin { get; set; }
            public string adi { get; set; }
            public string soyadi { get; set; }
            public string kullanici_adi { get; set; }
            public string sifre { get; set; }
        }

        public class Soru
        {
            public int id { get; set; }
            public int multiple { get; set; }
            public string soru { get; set; }
            public List<string> cevaplar { get; set; }
        }

        public class Cevap
        {
            public int id { get; set; }
            public int user_id { get; set; }
            public List<string> cevaplar { get; set; }
        }

        public void soru_Goster(Soru soru, Label label, Label soru_label, Control.ControlCollection controls)
        {
            label.Text = "Soru "+(soru.id + 1);
            soru_label.Text = soru.soru;
            radioButtonlar.Clear();
            checkBoxlar.Clear();
            combo.Items.Clear();
            lb.Items.Clear();
            groupBox1.Controls.Clear();

            if (soru.multiple == 0) // Radiobutton
            {
                for (int i = 0; i < soru.cevaplar.Count; i++)
                {
                    radioButtonlar.Add(new RadioButton());
                    radioButtonlar[i].AutoSize = true;
                    radioButtonlar[i].Location = new Point(19, (29 + (i * 30)));
                    radioButtonlar[i].Text = soru.cevaplar[i];
                    controls.Add(radioButtonlar[i]);
                }
            }
            else if (soru.multiple == 1) // Checkbox
            {
               
                for (int i = 0; i < soru.cevaplar.Count; i++)
                {
                    checkBoxlar.Add(new CheckBox());
                    checkBoxlar[i].AutoSize = true;
                    checkBoxlar[i].Location = new Point(19, (29 + (i * 30)));
                    checkBoxlar[i].Text = soru.cevaplar[i];
                    checkBoxlar[i].Visible = true;
                    controls.Add(checkBoxlar[i]);
                }
            }
            else if (soru.multiple == 2) // Combobox
            {
                combo.Items.AddRange(soru.cevaplar.ToArray());
                combo.Location = new Point(19, 29);
                combo.DropDownStyle = ComboBoxStyle.DropDownList;
                controls.Add(combo);
            }
            else if (soru.multiple == 3) // Combobox
            {
                lb.Items.AddRange(soru.cevaplar.ToArray());
                lb.Location = new Point(19, 29);
                controls.Add(lb);
            }

            button5.Location = new Point(32, 141 + groupBox1.Height + 20);
            panel3.Height = 32 + groupBox1.Height + 20 + 200;
            Size = panel3.Size;
        }
    }
}
