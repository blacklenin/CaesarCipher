using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace CaesarCipher
{
    public partial class Form1 : Form
    {
        OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
        SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
        private static int key = 2;
        private static bool isRight = true;

        public static void SetKey(int _key)
        {
            key = _key;
        }

        public static void setIsRight(bool _isRight)
        {
            isRight = _isRight;
        }

        private static Dictionary<char, int> alphabet = new Dictionary<char, int>()
        {
            {'А', 0}, {'Б', 1}, {'В', 2}, {'Г', 3}, {'Д', 4}, {'Е', 5}, {'Ё', 6}, {'Ж', 7}, {'З', 8},
            {'И', 9}, {'Й', 10}, {'К', 11}, {'Л', 12}, {'М', 13}, {'Н', 14}, {'О', 15}, {'П', 16},
            {'Р', 17}, {'С', 18}, {'Т', 19}, {'У', 20}, {'Ф', 21}, {'Х', 22}, {'Ц', 23}, {'Ч', 24},
            {'Ш', 25}, {'Щ', 26}, {'Ъ', 27}, {'Ы', 28}, {'Ь', 29}, {'Э', 30}, {'Ю', 31}, {'Я', 32},
            {'а', 33}, {'б', 34}, {'в', 35}, {'г', 36}, {'д', 37}, {'е', 38}, {'ё', 39}, {'ж', 40}, {'з', 41},
            {'и', 42}, {'й', 43}, {'к', 44}, {'л', 45}, {'м', 46}, {'н', 47}, {'о', 48}, {'п', 49},
            {'р', 50}, {'с', 51}, {'т', 52}, {'у', 53}, {'ф', 54}, {'х', 55}, {'ц', 56}, {'ч', 57},
            {'ш', 58}, {'щ', 59}, {'ъ', 60}, {'ы', 61}, {'ь', 62}, {'э', 63}, {'ю', 64}, {'я', 65},
            {'0', 66}, {'1', 67}, {'2', 68}, {'3', 69}, {'4', 70}, {'5', 71}, {'6', 72}, {'7', 73}, {'8', 74}, {'9', 75}
        };

        private static Dictionary<int, char> reverseAlphabet = new Dictionary<int, char>()
        {
            {0, 'А'}, {1, 'Б'}, {2, 'В'}, {3, 'Г'}, {4, 'Д'}, {5, 'Е'}, {6, 'Ё'}, {7, 'Ж'}, {8, 'З'},
            {9, 'И'}, {10, 'Й'}, {11, 'К'}, {12, 'Л'}, {13, 'М'}, {14, 'Н'},{15, 'О'}, {16, 'П'},
            {17, 'Р'}, {18, 'С'}, {19, 'Т'}, {20, 'У'}, {21, 'Ф'}, {22, 'Х'}, {23, 'Ц'}, {24, 'Ч'},
            {25, 'Ш'}, {26, 'Щ'}, {27, 'Ъ'}, {28, 'Ы'}, {29, 'Ь'}, {30, 'Э'}, {31, 'Ю'}, {32, 'Я'},
            {33, 'а'}, {34, 'б'}, {35, 'в'}, {36, 'г'}, {37, 'д'}, {38, 'е'}, {39, 'ё'}, {40, 'ж'}, {41, 'з'},
            {42, 'и'}, {43, 'й'}, {44, 'к'}, {45, 'л'}, {46, 'м'}, {47, 'н'},{48, 'о'}, {49, 'п'},
            {50, 'р'}, {51, 'с'}, {52, 'т'}, {53, 'у'}, {54, 'ф'}, {55, 'х'}, {56, 'ц'}, {57, 'ч'},
            {58, 'ш'}, {59, 'щ'}, {60, 'ъ'}, {61, 'ы'}, {62, 'ь'}, {63, 'э'}, {64, 'ю'}, {65, 'я'},
            {66, '0'}, {67, '1'}, {68, '2'}, {69, '3'}, {70, '4'}, {71, '5'}, {72, '6'}, {73, '7'}, {74, '8'}, {75, '9'}
        };

        public Form1()
        {
            InitializeComponent();

            OpenFileDialog1.Filter = "Text Files(*.txt)|*.txt|Document Files(*.docx)|*.docx";
            SaveFileDialog1.Filter = "Text Files(*.txt)|*.txt|Document Files(*.docx)|*.docx";

            richTextBox2.Text = "Текущий шаг сдвига: 2\nСдвиг: Вправо";
        }

        private void openToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (OpenFileDialog1.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }
            string filename = OpenFileDialog1.FileName;

            if (filename.EndsWith("txt"))
            {
                richTextBox1.Text = File.ReadAllText(filename, Encoding.UTF8);
            }
            else if (filename.EndsWith("docx"))
            {
                try
                {
                    Word.Application app = new Word.Application();
                    Word.Document doc = app.Documents.Open(filename);
                    richTextBox1.Text = doc.Content.Text;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    doc.Close();
                    Marshal.ReleaseComObject(doc);
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }
                catch(Exception exception)
                {
                    MessageBox.Show("Во время исполнения произошла ошибка.");
                }
            }
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (SaveFileDialog1.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }
            string filename = SaveFileDialog1.FileName;

            if (filename.EndsWith("txt"))
            {
                richTextBox1.SaveFile(filename, RichTextBoxStreamType.PlainText);
            }
            else if (filename.EndsWith("docx"))
            {
                try
                {
                    Word.Application app = new Word.Application();
                    Word.Document doc = app.Documents.Add();
                    doc.Content.Text = richTextBox1.Text;
                    doc.SaveAs2(filename);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    doc.Close();
                    Marshal.ReleaseComObject(doc);
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                }
                catch
                {
                    MessageBox.Show("Во время исполнения произошла ошибка.");
                }
            }
            MessageBox.Show("Файл был успешно сохранен!");
        }

        public static string Decrypt(string text)
        {
            string result = "";
            for (int i = 0; i < text.Length; i++)
            {
                if (alphabet.ContainsKey(text[i]))
                {
                    if (alphabet[text[i]] < 33)
                    {
                        if (alphabet[text[i]] - key < 0)
                            result += reverseAlphabet[alphabet[text[i]] - key + 33];
                        else if (alphabet[text[i]] - key > 32)
                            result += reverseAlphabet[alphabet[text[i]] - key - 33];
                        else
                            result += reverseAlphabet[alphabet[text[i]] - key];
                    }
                    else if (alphabet[text[i]] >= 33 && alphabet[text[i]] < 66)
                    {
                        if (alphabet[text[i]] - key < 33)
                            result += reverseAlphabet[alphabet[text[i]] - key + 33];
                        else if (alphabet[text[i]] - key > 65)
                            result += reverseAlphabet[alphabet[text[i]] - key - 33];
                        else
                            result += reverseAlphabet[alphabet[text[i]] - key];
                    }
                    else
                    {
                        if (alphabet[text[i]] - key < 66)
                            result += reverseAlphabet[alphabet[text[i]] - key + 10];
                        else if (alphabet[text[i]] - key > 75)
                            result += reverseAlphabet[alphabet[text[i]] - key - 10];
                        else
                            result += reverseAlphabet[alphabet[text[i]] - key];
                    }
                }
                else
                {
                    result += text[i];
                }
            }
            return result;
        }

        private void decryptToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string text = richTextBox1.Text;
            richTextBox1.Text = Decrypt(text);
        }

        public static string Encrypt(string text)
        {
            string result = "";
            for (int i = 0; i < text.Length; i++)
            {
                if (alphabet.ContainsKey(text[i]))
                {
                    if (alphabet[text[i]] < 33)
                    {
                        if (alphabet[text[i]] + key > 32)
                            result += reverseAlphabet[alphabet[text[i]] + key - 33];
                        else if (alphabet[text[i]] + key < 0)
                            result += reverseAlphabet[alphabet[text[i]] + key + 33];
                        else
                            result += reverseAlphabet[alphabet[text[i]] + key];
                    }
                    else if (alphabet[text[i]] >= 33 && alphabet[text[i]] < 66)
                    {
                        if (alphabet[text[i]] + key > 65)
                            result += reverseAlphabet[alphabet[text[i]] + key - 33];
                        else if (alphabet[text[i]] + key < 33)
                            result += reverseAlphabet[alphabet[text[i]] + key + 33];
                        else
                            result += reverseAlphabet[alphabet[text[i]] + key];
                    }
                    else
                    {
                        if (alphabet[text[i]] + key > 75)
                            result += reverseAlphabet[alphabet[text[i]] + key - 10];
                        else if (alphabet[text[i]] + key < 66)
                            result += reverseAlphabet[alphabet[text[i]] + key + 10];
                        else
                            result += reverseAlphabet[alphabet[text[i]] + key];
                    }
                }
                else
                {
                    result += text[i];
                }
            }
            return result;
        }

        private void encryptToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string text = richTextBox1.Text;
            richTextBox1.Text = Encrypt(text);
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string text = "Текущий шаг сдвига: ";
            if (!maskedTextBox1.Text.Equals(""))
                key = Convert.ToInt32(maskedTextBox1.Text);
            text += Math.Abs(key);
            if (radioButton1.Checked)
            {
                isRight = true;
                text += "\nСдвиг: Вправо";
                radioButton1.Checked = false;
            }
            else if (radioButton2.Checked)
            {
                isRight = false;
                text += "\nСдвиг: Влево";
                radioButton2.Checked = false;
            }
            else
            {
                if(!isRight)
                    text += "\nСдвиг: Влево";
                else
                    text += "\nСдвиг: Вправо";
            }
            if (isRight)
                key = Math.Abs(key);
            else
                key *= -1;
            richTextBox2.Text = text;
            maskedTextBox1.Clear();
        }
    }
}
