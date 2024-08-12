using System.Diagnostics.Metrics;
using System.Drawing.Printing;
using System.Reflection.Metadata;
using System.Text;
using System.Windows.Forms;
using static System.Collections.Specialized.BitVector32;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;






namespace SemesterProject_MsWord_
{
    public partial class Form1 : Form
    {
        public static string FindText;
        public static string ReplaceText;
        public static Boolean matchcase;
        int d;
        
        
        public Form1()
        {
            InitializeComponent();
        }

        private void uNDOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Undo();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Redo();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Copy();
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Paste();
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Cut();
        }

        private void boldToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionFont = new Font(richTextBox1.Font, FontStyle.Bold);
        }

        private void italicToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionFont = new Font(richTextBox1.Font, FontStyle.Italic);
        }

        private void unToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionFont = new Font(richTextBox1.Font, FontStyle.Underline);
        }

        private void increaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            float currentSize;
            currentSize = richTextBox1.Font.Size;
            currentSize += 12.0F;
            richTextBox1.Font = new Font(richTextBox1.Font.Name, currentSize,
            richTextBox1.Font.Style, richTextBox1.Font.Unit);
            

        }

        private void wordWrapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(richTextBox1.WordWrap = true)
            {
                richTextBox1.WordWrap = false;
            }
            else
            {
                richTextBox1.WordWrap = true;
            }
        }

        private void decreaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            float currentSize;
            currentSize = richTextBox1.Font.Size;
            currentSize -= 2.0F;
            richTextBox1.Font = new Font(richTextBox1.Font.Name, currentSize,
            richTextBox1.Font.Style, richTextBox1.Font.Unit);
        }

        private void sentenceCaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //string Text = richTextBox1.SelectedText;
            //richTextBox1.SelectedText = Text.();
        }

        private void lowerCaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Text = richTextBox1.SelectedText;
            richTextBox1.SelectedText = Text.ToLower();
        }

        private void uPPERCASEToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Text = richTextBox1.SelectedText;
            richTextBox1.SelectedText = Text.ToUpper();
        }

        private void cAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Text = richTextBox1.SelectedText;
            richTextBox1.SelectedText = Text.ToUpper();
        }

        private void paragraphToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //left
            richTextBox1.SelectionAlignment = HorizontalAlignment.Left;
        }

        private void centerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Center;
        }

        private void rightToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectionAlignment = HorizontalAlignment.Right;
        }

        private void aboutmswordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About ab=new About();
            ab.Show();
            
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Microsoft 365 Help & Learning",
               "MsWord Help", MessageBoxButtons.OK, MessageBoxIcon.Question);
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog open1 = new OpenFileDialog();
            open1.Title = "Save";
            open1.Filter = "TextDocument(*.txt)|*.txt|All Files(*.*)|*.*";
            if (open1.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.LoadFile(open1.FileName, RichTextBoxStreamType.PlainText);
                this.Text = open1.FileName;
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Save";
            save.Filter = "TextDocument(*.txt)|*.txt|All Files(*.*)|*.*";
            if (save.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.SaveFile(save.FileName, RichTextBoxStreamType.PlainText);
                this.Text = save.FileName;
            }
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Title = "Save";
            save.Filter = "TextDocument(*.txt)|*.txt|All Files(*.*)|*.*";
            if (save.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.SaveFile(save.FileName, RichTextBoxStreamType.PlainText);
                this.Text = save.FileName;
            }
        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintDialog pnt = new PrintDialog();
            if (pnt.ShowDialog() == DialogResult.OK)
            {

            }

        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectedText = "";
        }

        private void deleteAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        private void dateAndTimeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            richTextBox1.Text = dt.ToString();
        }

        private void tableToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StringBuilder strTable = new StringBuilder();
            strTable.Append(@"{\rtf1 ");
            for (int i = 0; i < 5; i++)
            {
                strTable.Append(@"\trowd");
                strTable.Append(@"\cellx1000");
                strTable.Append(@"\cellx2000");
                strTable.Append(@"\cellx3000");
                strTable.Append(@"\cellx4000");
                strTable.Append(@"\intbl \cell \row");
            }
            strTable.Append(@"\pard");
            strTable.Append(@"}");
            this.richTextBox1.Rtf = strTable.ToString();
        }

        private void picturesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog op = new OpenFileDialog())
            {
                op.Filter = "Image Files(*.jpg; *.jpeg; *.gif; *.png; *.bmp)|*.jpg;" +
                    " *.jpeg; *.gif; *.png; *.bmp";
                if (op.ShowDialog() == DialogResult.OK)
                {
                    Clipboard.SetImage(Image.FromFile(op.FileName));
                    richTextBox1.Paste();
                }
            }
        }

        private void lineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //nothing
        }

        private void rectangleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //nothing
        }

        private void marginsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.BackColor = Color.LightGray;
            richTextBox1.Padding = new Padding(6);
            richTextBox1.Location = new Point(100, 50);
            richTextBox1.AutoSize = true;
            Font = new Font("French Script MT", 15);
            richTextBox1.Margin = new Padding(12, 12, 12, 12);
            this.Controls.Add(richTextBox1);


        }

        private void zoomInToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.ZoomFactor += 2.0f;
        }

        private void zoomOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            float currentSize;
            currentSize = richTextBox1.Font.Size;
            currentSize -= 2.0F;
            richTextBox1.Font = new Font(richTextBox1.Font.Name, currentSize,
            richTextBox1.Font.Style, richTextBox1.Font.Unit);
        }

        private void landspaceToolStripMenuItem_Click(object sender, EventArgs e)
        {

            //nothing

        }

        private void fontStyleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FontDialog fnt = new FontDialog();
            if (fnt.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.Font = fnt.Font;
            }
        }

        private void fontSizeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FontDialog fnt = new FontDialog();
            if (fnt.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.Font = fnt.Font;
            }
        }

        private void fontColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ColorDialog fnt = new ColorDialog();
            if (fnt.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.ForeColor = fnt.Color;
            }
        }

        private void searchAndHighlightTextToolStripMenuItem_Click(object sender, EventArgs e)
        {

            //nothing
        }

        private void searchToolStripMenuItem1_Click(object sender, EventArgs e)
        {

            //nothing
        }

        private void highlightToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //nothing
        }

        private void findToolStripMenuItem_Click(object sender, EventArgs e)
        {

            Find f = new Find();
            f.ShowDialog();
            if (FindText != "")
            {
                d = richTextBox1.Find(FindText);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            //nothing
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //For Search
            string[] words = txtSearch.Text.Split(',');
            foreach (string word in words)
            {
                int startIndex = 0;
                while (startIndex < richTextBox1.TextLength)
                {

                    int wordStartIndex = richTextBox1.Find(word, startIndex, RichTextBoxFinds.None);
                    if (wordStartIndex != -1)
                    {

                        richTextBox1.SelectionStart = wordStartIndex;
                        richTextBox1.SelectionLength = word.Length;
                        richTextBox1.SelectionBackColor = Color.Red;
                    }
                    else
                        break;
                    startIndex += wordStartIndex + word.Length;
                }
            }
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectAll();
        }

        private void portraitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //nothing
        }

        private void findNextToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (FindText != "")
            {
                if (matchcase == true)
                {
                    d = richTextBox1.Find(FindText, (d + 1), richTextBox1.Text.Length, RichTextBoxFinds.MatchCase);
                }
                else
                {
                    d = richTextBox1.Find(FindText, (d + 1), richTextBox1.Text.Length, RichTextBoxFinds.None);
                }
            }
        }

        private void searchToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Replace r = new Replace();
            r.ShowDialog();
            richTextBox1.Find(FindText);
            richTextBox1.SelectedText = ReplaceText;
        }

        private void backGroundColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
          
            ColorDialog col = new ColorDialog();
            col.ShowDialog();
            richTextBox1.BackColor=col.Color;
        }

        private void selectFontColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ColorDialog col1 = new ColorDialog();
            col1.ShowDialog();
            richTextBox1.SelectionColor=col1.Color;
        }

        private void selectBackGroundColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ColorDialog col2 = new ColorDialog();
            col2.ShowDialog();
            richTextBox1.SelectionBackColor = col2.Color;
        }

        private void circleToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //nothing
        }

        private void bulletsToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
            richTextBox1.SelectionIndent = 15;
            richTextBox1.BulletIndent = 5;
            richTextBox1.SelectionBullet = true;
        }

        private void bordersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            richTextBox1.Text = "";
            richTextBox1.Location=new Point(236,90);
            richTextBox1.Size = new Size(500, 200);
            richTextBox1.BorderStyle= BorderStyle.Fixed3D;
            this.Controls.Add(richTextBox1);
           
        }

        private void lineNumberStripToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //nothing
        }

        private void mathEquationsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //nothing

        }

        private void sortingToolStripMenuItem_Click(object sender, EventArgs e)
        {
           //nothing
        }

        private void pageSetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //nothing
        }

        private void paintToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Paint p=new Paint();
            p.ShowDialog();
        }

        private void playSoundToolStripMenuItem_Click(object sender, EventArgs e)
        {
           //nothing
        }
    }
}
    
