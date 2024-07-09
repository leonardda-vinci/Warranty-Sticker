using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media;
using Syncfusion.Windows.Forms.Barcode;
using Color = System.Drawing.Color;
using System.Diagnostics;
using Font = System.Drawing.Font;
using System.Linq;
using System.Drawing.Text;
using System.Text;
using System.Xml.Linq;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Printing;

namespace WarrantySticker
{
    public partial class WarrantyForm : Form
    {
        string strINIPath = Application.StartupPath + "\\\\ipadd.ini";
        string successPath = Application.StartupPath + @"\SuccessFile";
        string ImagePath = Application.StartupPath + @"\Generated Image";
        public WarrantyForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (InitClass.CheckApp() == true)
            {
                MessageBox.Show("Application is already running.", "Get Data From Weighing Scale Integration", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
                Application.Exit();
            }
            Initialize();
            backgroundWorker1.RunWorkerAsync();
            CheckForIllegalCrossThreadCalls = false;
        }

        public void Initialize()
        {
            try
            {
                SetLabelText(" Setting up your files and folders, please wait ...");
                if (SystemSetup.DefaultFolders() == false)
                {
                    SetLabelText("Error occured in getting users data. Please check Error Log File.");
                    this.Close();
                }

                SetLabelText(" Set up completed...");
            }
            catch (Exception ex)
            {
                SystemSetup.ErrorAppend(string.Format("Error Code: {0} " + "Error Desc: {1} " + "-999 " + ex.Message));
                GC.Collect();
            }
        }
        private void SetLabelText(string val)
        {
            try
            {
                if (string.IsNullOrEmpty(txtStat.Text))
                {
                    txtStat.AppendText(val);
                }
                else
                {
                    txtStat.AppendText(Environment.NewLine + val);
                }
            }
            catch (Exception ex)
            {
                SystemSetup.ErrorAppend(string.Format("Error Code: {0} " + "Error Desc: {1} " + "-999 " + ex.Message));
            }
        }

        private bool WarrantySticker()
        {
            try
            {
                var IniGet = new INIFile(strINIPath);
                var txtPath = IniGet.Read("txtPath");

                //for printing xml file
                string xmlFiles = "*.xml";
                string[] GetXMLFiles = Directory.GetFiles(txtPath, xmlFiles);

                foreach (var XML in GetXMLFiles)
                {
                    using (var reader = new StreamReader(XML))
                    {
                        string value = reader.ReadToEnd();
                        string[] spliter = value.Split('\n');

                        //Get SEARCH INDEX
                        string Search1 = ">";
                        string Search2 = "<";
                        string Positionx = " x=\"";
                        string Positiony = " y=\"";
                        string Fontsearch = " name=\"";
                        string Sizesearch = "size=\"";
                        string Widthsearch = "width=\"";
                        string Heightsearch = "height=\"";
                        string Colorsearch = "color=\"";
                        string Multipliersearch = "multiplier=\"";
                        string Positionlast = "\"";
                        string isBold = " bold=\"";
                        string CFont = "font_id=\"";
                        //For Line
                        string x1 = " x1=\"";
                        string x2 = " x2=\"";
                        string y1 = " y1=\"";
                        string y2 = " y2=\"";

                        //Get Fonts
                        //ID1 fontname
                        string IDF1Line = spliter[3].ToString();
                        int IDF1start1 = IDF1Line.IndexOf(Fontsearch) + Fontsearch.Length;
                        int IDF1end1 = IDF1Line.IndexOf(Positionlast, IDF1start1);
                        string IDF1 = IDF1Line.ToUpper().Substring(IDF1start1, IDF1end1 - IDF1start1);
                        //ID1 Size
                        int IDSstart1 = IDF1Line.IndexOf(Sizesearch) + Sizesearch.Length;
                        int IDSend1 = IDF1Line.IndexOf(Positionlast, IDSstart1);
                        string IDSS1 = IDF1Line.ToUpper().Substring(IDSstart1, IDSend1 - IDSstart1);
                        int IDS1 = Int32.Parse(IDSS1);
                        // ID1 isBold
                        int IDB1start1 = IDF1Line.IndexOf(isBold) + isBold.Length;
                        int IDB1end1 = IDF1Line.IndexOf(Positionlast, IDB1start1);
                        string IDB1 = IDF1Line.ToUpper().Substring(IDB1start1, IDB1end1 - IDB1start1);

                        string IDF2Line = spliter[4].ToString();
                        int IDF2start2 = IDF2Line.IndexOf(Fontsearch) + Fontsearch.Length;
                        int IDF2end2 = IDF2Line.IndexOf(Positionlast, IDF2start2);
                        string IDF2 = IDF2Line.ToUpper().Substring(IDF2start2, IDF2end2 - IDF2start2);
                        //ID2 Size
                        int IDSstart2 = IDF2Line.IndexOf(Sizesearch) + Sizesearch.Length;
                        int IDSend2 = IDF2Line.IndexOf(Positionlast, IDSstart2);
                        string IDSS2 = IDF2Line.ToUpper().Substring(IDSstart2, IDSend2 - IDSstart2);
                        int IDS2 = Int32.Parse(IDSS2);
                        // ID2 isBold
                        int IDB2start2 = IDF2Line.IndexOf(isBold) + isBold.Length;
                        int IDB2end2 = IDF2Line.IndexOf(Positionlast, IDB2start2);
                        string IDB2 = IDF2Line.ToUpper().Substring(IDB2start2, IDB2end2 - IDB2start2);

                        string IDF3Line = spliter[5].ToString();
                        int IDF3start3 = IDF3Line.IndexOf(Fontsearch) + Fontsearch.Length;
                        int IDF3end3 = IDF3Line.IndexOf(Positionlast, IDF3start3);
                        string IDF3 = IDF3Line.ToUpper().Substring(IDF3start3, IDF3end3 - IDF3start3);
                        //ID3 Size
                        int IDSstart3 = IDF3Line.IndexOf(Sizesearch) + Sizesearch.Length;
                        int IDSend3 = IDF3Line.IndexOf(Positionlast, IDSstart3);
                        string IDSS3 = IDF3Line.ToUpper().Substring(IDSstart3, IDSend3 - IDSstart3);
                        int IDS3 = Int32.Parse(IDSS3);
                        // ID3 isBold
                        int IDB3start3 = IDF3Line.IndexOf(isBold) + isBold.Length;
                        int IDB3end3 = IDF3Line.IndexOf(Positionlast, IDB3start3);
                        string IDB3 = IDF3Line.ToUpper().Substring(IDB3start3, IDB3end3 - IDB3start3);

                        string IDF4Line = spliter[6].ToString();
                        int IDF4start4 = IDF4Line.IndexOf(Fontsearch) + Fontsearch.Length;
                        int IDF4end4 = IDF4Line.IndexOf(Positionlast, IDF4start4);
                        string IDF4 = IDF4Line.ToUpper().Substring(IDF4start4, IDF4end4 - IDF4start4);
                        //ID4 Size
                        int IDSstart4 = IDF4Line.IndexOf(Sizesearch) + Sizesearch.Length;
                        int IDSend4 = IDF4Line.IndexOf(Positionlast, IDSstart4);
                        string IDSS4 = IDF4Line.ToUpper().Substring(IDSstart4, IDSend4 - IDSstart4);
                        int IDS4 = Int32.Parse(IDSS4);
                        // ID4 isBold
                        int IDB4start4 = IDF4Line.IndexOf(isBold) + isBold.Length;
                        int IDB4end4 = IDF4Line.IndexOf(Positionlast, IDB4start4);
                        string IDB4 = IDF4Line.ToUpper().Substring(IDB4start4, IDB4end4 - IDB4start4);

                        //Header
                        //find Header
                        string HeaderLine = spliter[14].ToString();
                        int Headerstart1 = HeaderLine.IndexOf(Search1) + Search1.Length;
                        int Headerend1 = HeaderLine.IndexOf(Search2, Headerstart1);
                        string Header = HeaderLine.ToUpper().Substring(Headerstart1, Headerend1 - Headerstart1);
                        string FinalHeader = String.Join(" ", Header.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find Header locX
                        int HeaderXstart = HeaderLine.IndexOf(Positionx) + Positionx.Length;
                        int HeaderXend = HeaderLine.IndexOf(Positionlast, HeaderXstart);
                        string HeaderaX = HeaderLine.ToUpper().Substring(HeaderXstart, HeaderXend - HeaderXstart);
                        int HeaderX = Int32.Parse(HeaderaX);
                        //find Header locXY
                        int HeaderYstart = HeaderLine.IndexOf(Positiony) + Positiony.Length;
                        int HeaderYend = HeaderLine.IndexOf(Positionlast, HeaderYstart);
                        string HeaderaY = HeaderLine.ToUpper().Substring(HeaderYstart, HeaderYend - HeaderYstart);
                        int HeaderY = Int32.Parse(HeaderaY);
                        //Find Header width
                        int HeaderWstart = HeaderLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int HeaderWend = HeaderLine.IndexOf(Positionlast, HeaderWstart);
                        string HeaderaW = HeaderLine.ToUpper().Substring(HeaderWstart, HeaderWend - HeaderWstart);
                        int HeaderW = Int32.Parse(HeaderaW);
                        float HeaderFW = float.Parse(HeaderaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find Header height
                        int HeaderHstart = HeaderLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int HeaderHend = HeaderLine.IndexOf(Positionlast, HeaderHstart);
                        string HeaderaH = HeaderLine.ToUpper().Substring(HeaderHstart, HeaderHend - HeaderHstart);
                        int HeaderH = Int32.Parse(HeaderaH);
                        float HeaderFH = float.Parse(HeaderaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find Header Color
                        int HeaderCstart1 = HeaderLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int HeaderCend1 = HeaderLine.IndexOf(Positionlast, HeaderCstart1);
                        string HeaderCH = HeaderLine.ToUpper().Substring(HeaderCstart1, HeaderCend1 - HeaderCstart1);
                        string HeaderCF = HeaderCH.Replace("0x", "#");
                        Color HeaderC = ColorTranslator.FromHtml(HeaderCF);
                        if (HeaderCF == "000000")
                        {
                            HeaderC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var Headerbrush = new SolidBrush(Color.FromArgb(HeaderC.A, HeaderC.R, HeaderC.G, HeaderC.B));
                        //find Header Font
                        int HeaderFontstart = HeaderLine.IndexOf(CFont) + CFont.Length;
                        int HeaderFontend = HeaderLine.IndexOf(Positionlast, HeaderFontstart);
                        string HeaderaFont = HeaderLine.ToUpper().Substring(HeaderFontstart, HeaderFontend - HeaderFontstart);
                        Font HeaderFont;
                        if (HeaderaFont == "1")
                        {
                            HeaderFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                       else if (HeaderaFont == "2")
                        {
                            HeaderFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (HeaderaFont == "3")
                        {
                            HeaderFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            HeaderFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //Lot_txt
                        //find Lot_txt
                        string Lot_txtLine = spliter[15].ToString();
                        int Lot_txtstart1 = Lot_txtLine.IndexOf(Search1) + Search1.Length;
                        int Lot_txtend1 = Lot_txtLine.IndexOf(Search2, Lot_txtstart1);
                        string Lot_txt = Lot_txtLine.ToUpper().Substring(Lot_txtstart1, Lot_txtend1 - Lot_txtstart1);
                        string FinalLot_txt = String.Join(" ", Lot_txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find Lot_txt locX
                        int Lot_txtXstart = Lot_txtLine.IndexOf(Positionx) + Positionx.Length;
                        int Lot_txtXend = Lot_txtLine.IndexOf(Positionlast, Lot_txtXstart);
                        string Lot_txtaX = Lot_txtLine.ToUpper().Substring(Lot_txtXstart, Lot_txtXend - Lot_txtXstart);
                        int Lot_txtX = Int32.Parse(Lot_txtaX);
                        //find Lot_txt locXY
                        int Lot_txtYstart = Lot_txtLine.IndexOf(Positiony) + Positiony.Length;
                        int Lot_txtYend = Lot_txtLine.IndexOf(Positionlast, Lot_txtYstart);
                        string Lot_txtaY = Lot_txtLine.ToUpper().Substring(Lot_txtYstart, Lot_txtYend - Lot_txtYstart);
                        int Lot_txtY = Int32.Parse(Lot_txtaY);
                        //find Lot_txt Color
                        int Lot_txtCstart1 = Lot_txtLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int Lot_txtCend1 = Lot_txtLine.IndexOf(Positionlast, Lot_txtCstart1);
                        string Lot_txtCH = Lot_txtLine.ToUpper().Substring(Lot_txtCstart1, Lot_txtCend1 - Lot_txtCstart1);
                        string Lot_txtCF = Lot_txtCH.Replace("0x", "#");
                        Color Lot_txtC = ColorTranslator.FromHtml(Lot_txtCF);
                        if (Lot_txtCF == "000000")
                        {
                            Lot_txtC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var Lot_txtbrush = new SolidBrush(Color.FromArgb(Lot_txtC.A, Lot_txtC.R, Lot_txtC.G, Lot_txtC.B));
                        //find Lot_txt Font
                        int Lot_txtFontstart = Lot_txtLine.IndexOf(CFont) + CFont.Length;
                        int Lot_txtFontend = Lot_txtLine.IndexOf(Positionlast, Lot_txtFontstart);
                        string Lot_txtaFont = Lot_txtLine.ToUpper().Substring(Lot_txtFontstart, Lot_txtFontend - Lot_txtFontstart);
                        Font Lot_txtFont;
                        if (Lot_txtaFont == "1")
                        {
                            Lot_txtFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (Lot_txtaFont == "2")
                        {
                            Lot_txtFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (Lot_txtaFont == "3")
                        {
                            Lot_txtFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            Lot_txtFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //Lot
                        //find Lot
                        string LotLine = spliter[16].ToString();
                        int Lotstart1 = LotLine.IndexOf(Search1) + Search1.Length;
                        int Lotend1 = LotLine.IndexOf(Search2, Lotstart1);
                        string Lot = LotLine.ToUpper().Substring(Lotstart1, Lotend1 - Lotstart1);
                        string FinalLot = String.Join(" ", Lot.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find Lot locX
                        int LotXstart = LotLine.IndexOf(Positionx) + Positionx.Length;
                        int LotXend = LotLine.IndexOf(Positionlast, LotXstart);
                        string LotaX = LotLine.ToUpper().Substring(LotXstart, LotXend - LotXstart);
                        int LotX = Int32.Parse(LotaX);
                        //find Lot locXY
                        int LotYstart = LotLine.IndexOf(Positiony) + Positiony.Length;
                        int LotYend = LotLine.IndexOf(Positionlast, LotYstart);
                        string LotaY = LotLine.ToUpper().Substring(LotYstart, LotYend - LotYstart);
                        int LotY = Int32.Parse(LotaY);
                        //find Lot Color
                        int LotCstart1 = LotLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int LotCend1 = LotLine.IndexOf(Positionlast, LotCstart1);
                        string LotCH = LotLine.ToUpper().Substring(LotCstart1, LotCend1 - LotCstart1);
                        string LotCF = LotCH.Replace("0x", "#");
                        Color LotC = ColorTranslator.FromHtml(LotCF);
                        if (LotCF == "000000")
                        {
                            LotC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var Lotbrush = new SolidBrush(Color.FromArgb(LotC.A, LotC.R, LotC.G, LotC.B));
                        //find Lot Font
                        int LotFontstart = LotLine.IndexOf(CFont) + CFont.Length;
                        int LotFontend = LotLine.IndexOf(Positionlast, LotFontstart);
                        string LotaFont = LotLine.ToUpper().Substring(LotFontstart, LotFontend - LotFontstart);
                        Font LotFont;
                        if (LotaFont == "1")
                        {
                            LotFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (LotaFont == "2")
                        {
                            LotFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (LotaFont == "3")
                        {
                            LotFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            LotFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //Description_txt
                        //find Description_txt
                        string Description_txtLine = spliter[17].ToString();
                        int Description_txtstart1 = Description_txtLine.IndexOf(Search1) + Search1.Length;
                        int Description_txtend1 = Description_txtLine.IndexOf(Search2, Description_txtstart1);
                        string Description_txt = Description_txtLine.ToUpper().Substring(Description_txtstart1, Description_txtend1 - Description_txtstart1);
                        string FinalDescription_txt = String.Join(" ", Description_txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find Description_txt locX
                        int Description_txtXstart = Description_txtLine.IndexOf(Positionx) + Positionx.Length;
                        int Description_txtXend = Description_txtLine.IndexOf(Positionlast, Description_txtXstart);
                        string Description_txtaX = Description_txtLine.ToUpper().Substring(Description_txtXstart, Description_txtXend - Description_txtXstart);
                        int Description_txtX = Int32.Parse(Description_txtaX);
                        //find Description_txt locXY
                        int Description_txtYstart = Description_txtLine.IndexOf(Positiony) + Positiony.Length;
                        int Description_txtYend = Description_txtLine.IndexOf(Positionlast, Description_txtYstart);
                        string Description_txtaY = Description_txtLine.ToUpper().Substring(Description_txtYstart, Description_txtYend - Description_txtYstart);
                        int Description_txtY = Int32.Parse(Description_txtaY);
                        //find Description_txt Color
                        int Description_txtCstart1 = Description_txtLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int Description_txtCend1 = Description_txtLine.IndexOf(Positionlast, Description_txtCstart1);
                        string Description_txtCH = Description_txtLine.ToUpper().Substring(Description_txtCstart1, Description_txtCend1 - Description_txtCstart1);
                        string Description_txtCF = Description_txtCH.Replace("0x", "#");
                        Color Description_txtC = ColorTranslator.FromHtml(Description_txtCF);
                        if (Description_txtCF == "000000")
                        {
                            Description_txtC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var Description_txtbrush = new SolidBrush(Color.FromArgb(Description_txtC.A, Description_txtC.R, Description_txtC.G, Description_txtC.B));
                        //find Description_txt Font
                        int Description_txtFontstart = Description_txtLine.IndexOf(CFont) + CFont.Length;
                        int Description_txtFontend = Description_txtLine.IndexOf(Positionlast, Description_txtFontstart);
                        string Description_txtaFont = Description_txtLine.ToUpper().Substring(Description_txtFontstart, Description_txtFontend - Description_txtFontstart);
                        Font Description_txtFont;
                        if (Description_txtaFont == "1")
                        {
                            Description_txtFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (Description_txtaFont == "2")
                        {
                            Description_txtFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (Description_txtaFont == "3")
                        {
                            Description_txtFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            Description_txtFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //Description
                        //find Description
                        string DescriptionLine = spliter[18].ToString();
                        int Descriptionstart1 = DescriptionLine.IndexOf(Search1) + Search1.Length;
                        int Descriptionend1 = DescriptionLine.IndexOf(Search2, Descriptionstart1);
                        string Description = DescriptionLine.ToUpper().Substring(Descriptionstart1, Descriptionend1 - Descriptionstart1);
                        string FinalDescription = String.Join(" ", Description.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find Description locX
                        int DescriptionXstart = DescriptionLine.IndexOf(Positionx) + Positionx.Length;
                        int DescriptionXend = DescriptionLine.IndexOf(Positionlast, DescriptionXstart);
                        string DescriptionaX = DescriptionLine.ToUpper().Substring(DescriptionXstart, DescriptionXend - DescriptionXstart);
                        int DescriptionX = Int32.Parse(DescriptionaX);
                        //find Description locXY
                        int DescriptionYstart = DescriptionLine.IndexOf(Positiony) + Positiony.Length;
                        int DescriptionYend = DescriptionLine.IndexOf(Positionlast, DescriptionYstart);
                        string DescriptionaY = DescriptionLine.ToUpper().Substring(DescriptionYstart, DescriptionYend - DescriptionYstart);
                        int DescriptionY = Int32.Parse(DescriptionaY);
                        //find Description Color
                        int DescriptionCstart1 = DescriptionLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int DescriptionCend1 = DescriptionLine.IndexOf(Positionlast, DescriptionCstart1);
                        string DescriptionCH = DescriptionLine.ToUpper().Substring(DescriptionCstart1, DescriptionCend1 - DescriptionCstart1);
                        string DescriptionCF = DescriptionCH.Replace("0x", "#");
                        Color DescriptionC = ColorTranslator.FromHtml(DescriptionCF);
                        if (DescriptionCF == "000000")
                        {
                            DescriptionC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var Descriptionbrush = new SolidBrush(Color.FromArgb(DescriptionC.A, DescriptionC.R, DescriptionC.G, DescriptionC.B));
                        //find Description Font
                        int DescriptionFontstart = DescriptionLine.IndexOf(CFont) + CFont.Length;
                        int DescriptionFontend = DescriptionLine.IndexOf(Positionlast, DescriptionFontstart);
                        string DescriptionaFont = DescriptionLine.ToUpper().Substring(DescriptionFontstart, DescriptionFontend - DescriptionFontstart);
                        Font DescriptionFont;
                        if (DescriptionaFont == "1")
                        {
                            DescriptionFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (DescriptionaFont == "2")
                        {
                            DescriptionFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (DescriptionaFont == "3")
                        {
                            DescriptionFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            DescriptionFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //MFGSN_txt
                        //find MFGSN_txt
                        string MFGSN_txtLine = spliter[19].ToString();
                        int MFGSN_txtstart1 = MFGSN_txtLine.IndexOf(Search1) + Search1.Length;
                        int MFGSN_txtend1 = MFGSN_txtLine.IndexOf(Search2, MFGSN_txtstart1);
                        string MFGSN_txt = MFGSN_txtLine.ToUpper().Substring(MFGSN_txtstart1, MFGSN_txtend1 - MFGSN_txtstart1);
                        string FinalMFGSN_txt = String.Join(" ", MFGSN_txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find MFGSN_txt locX
                        int MFGSN_txtXstart = MFGSN_txtLine.IndexOf(Positionx) + Positionx.Length;
                        int MFGSN_txtXend = MFGSN_txtLine.IndexOf(Positionlast, MFGSN_txtXstart);
                        string MFGSN_txtaX = MFGSN_txtLine.ToUpper().Substring(MFGSN_txtXstart, MFGSN_txtXend - MFGSN_txtXstart);
                        int MFGSN_txtX = Int32.Parse(MFGSN_txtaX);
                        //find MFGSN_txt locXY
                        int MFGSN_txtYstart = MFGSN_txtLine.IndexOf(Positiony) + Positiony.Length;
                        int MFGSN_txtYend = MFGSN_txtLine.IndexOf(Positionlast, MFGSN_txtYstart);
                        string MFGSN_txtaY = MFGSN_txtLine.ToUpper().Substring(MFGSN_txtYstart, MFGSN_txtYend - MFGSN_txtYstart);
                        int MFGSN_txtY = Int32.Parse(MFGSN_txtaY);
                        //find MFGSN_txt Color
                        int MFGSN_txtCstart1 = MFGSN_txtLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int MFGSN_txtCend1 = MFGSN_txtLine.IndexOf(Positionlast, MFGSN_txtCstart1);
                        string MFGSN_txtCH = MFGSN_txtLine.ToUpper().Substring(MFGSN_txtCstart1, MFGSN_txtCend1 - MFGSN_txtCstart1);
                        string MFGSN_txtCF = MFGSN_txtCH.Replace("0x", "#");
                        Color MFGSN_txtC = ColorTranslator.FromHtml(MFGSN_txtCF);
                        if (MFGSN_txtCF == "000000")
                        {
                            MFGSN_txtC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var MFGSN_txtbrush = new SolidBrush(Color.FromArgb(MFGSN_txtC.A, MFGSN_txtC.R, MFGSN_txtC.G, MFGSN_txtC.B));
                        //find MFGSN_txt Font
                        int MFGSN_txtFontstart = MFGSN_txtLine.IndexOf(CFont) + CFont.Length;
                        int MFGSN_txtFontend = MFGSN_txtLine.IndexOf(Positionlast, MFGSN_txtFontstart);
                        string MFGSN_txtaFont = MFGSN_txtLine.ToUpper().Substring(MFGSN_txtFontstart, MFGSN_txtFontend - MFGSN_txtFontstart);
                        Font MFGSN_txtFont;
                        if (MFGSN_txtaFont == "1")
                        {
                            MFGSN_txtFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (MFGSN_txtaFont == "2")
                        {
                            MFGSN_txtFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (MFGSN_txtaFont == "3")
                        {
                            MFGSN_txtFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            MFGSN_txtFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //MFGSN
                        //find MFGSN
                        string MFGSNLine = spliter[20].ToString();
                        int MFGSNstart1 = MFGSNLine.IndexOf(Search1) + Search1.Length;
                        int MFGSNend1 = MFGSNLine.IndexOf(Search2, MFGSNstart1);
                        string MFGSN = MFGSNLine.ToUpper().Substring(MFGSNstart1, MFGSNend1 - MFGSNstart1);
                        string FinalMFGSN = String.Join(" ", MFGSN.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find MFGSN locX
                        int MFGSNXstart = MFGSNLine.IndexOf(Positionx) + Positionx.Length;
                        int MFGSNXend = MFGSNLine.IndexOf(Positionlast, MFGSNXstart);
                        string MFGSNaX = MFGSNLine.ToUpper().Substring(MFGSNXstart, MFGSNXend - MFGSNXstart);
                        int MFGSNX = Int32.Parse(MFGSNaX);
                        //find MFGSN locXY
                        int MFGSNYstart = MFGSNLine.IndexOf(Positiony) + Positiony.Length;
                        int MFGSNYend = MFGSNLine.IndexOf(Positionlast, MFGSNYstart);
                        string MFGSNaY = MFGSNLine.ToUpper().Substring(MFGSNYstart, MFGSNYend - MFGSNYstart);
                        int MFGSNY = Int32.Parse(MFGSNaY);
                        //find MFGSN Color
                        int MFGSNCstart1 = MFGSNLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int MFGSNCend1 = MFGSNLine.IndexOf(Positionlast, MFGSNCstart1);
                        string MFGSNCH = MFGSNLine.ToUpper().Substring(MFGSNCstart1, MFGSNCend1 - MFGSNCstart1);
                        string MFGSNCF = MFGSNCH.Replace("0x", "#");
                        Color MFGSNC = ColorTranslator.FromHtml(MFGSNCF);
                        if (MFGSNCF == "000000")
                        {
                            MFGSNC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var MFGSNbrush = new SolidBrush(Color.FromArgb(MFGSNC.A, MFGSNC.R, MFGSNC.G, MFGSNC.B));
                        //find MFGSN Font
                        int MFGSNFontstart = MFGSNLine.IndexOf(CFont) + CFont.Length;
                        int MFGSNFontend = MFGSNLine.IndexOf(Positionlast, MFGSNFontstart);
                        string MFGSNaFont = MFGSNLine.ToUpper().Substring(MFGSNFontstart, MFGSNFontend - MFGSNFontstart);
                        Font MFGSNFont;
                        if (MFGSNaFont == "1")
                        {
                            MFGSNFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (MFGSNaFont == "2")
                        {
                            MFGSNFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (MFGSNaFont == "3")
                        {
                            MFGSNFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            MFGSNFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //WarehouseLocation_txt
                        //find WarehouseLocation_txt
                        string WarehouseLocation_txtLine = spliter[21].ToString();
                        int WarehouseLocation_txtstart1 = WarehouseLocation_txtLine.IndexOf(Search1) + Search1.Length;
                        int WarehouseLocation_txtend1 = WarehouseLocation_txtLine.IndexOf(Search2, WarehouseLocation_txtstart1);
                        string WarehouseLocation_txt = WarehouseLocation_txtLine.ToUpper().Substring(WarehouseLocation_txtstart1, WarehouseLocation_txtend1 - WarehouseLocation_txtstart1);
                        string FinalWarehouseLocation_txt = String.Join(" ", WarehouseLocation_txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find WarehouseLocation_txt locX
                        int WarehouseLocation_txtXstart = WarehouseLocation_txtLine.IndexOf(Positionx) + Positionx.Length;
                        int WarehouseLocation_txtXend = WarehouseLocation_txtLine.IndexOf(Positionlast, WarehouseLocation_txtXstart);
                        string WarehouseLocation_txtaX = WarehouseLocation_txtLine.ToUpper().Substring(WarehouseLocation_txtXstart, WarehouseLocation_txtXend - WarehouseLocation_txtXstart);
                        int WarehouseLocation_txtX = Int32.Parse(WarehouseLocation_txtaX);
                        //find WarehouseLocation_txt locXY
                        int WarehouseLocation_txtYstart = WarehouseLocation_txtLine.IndexOf(Positiony) + Positiony.Length;
                        int WarehouseLocation_txtYend = WarehouseLocation_txtLine.IndexOf(Positionlast, WarehouseLocation_txtYstart);
                        string WarehouseLocation_txtaY = WarehouseLocation_txtLine.ToUpper().Substring(WarehouseLocation_txtYstart, WarehouseLocation_txtYend - WarehouseLocation_txtYstart);
                        int WarehouseLocation_txtY = Int32.Parse(WarehouseLocation_txtaY);
                        //Find WarehouseLocation_txt width
                        int WarehouseLocation_txtWstart = WarehouseLocation_txtLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int WarehouseLocation_txtWend = WarehouseLocation_txtLine.IndexOf(Positionlast, WarehouseLocation_txtWstart);
                        string WarehouseLocation_txtaW = WarehouseLocation_txtLine.ToUpper().Substring(WarehouseLocation_txtWstart, WarehouseLocation_txtWend - WarehouseLocation_txtWstart);
                        int WarehouseLocation_txtW = Int32.Parse(WarehouseLocation_txtaW);
                        float WarehouseLocation_txtFW = float.Parse(WarehouseLocation_txtaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find WarehouseLocation_txt height
                        int WarehouseLocation_txtHstart = WarehouseLocation_txtLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int WarehouseLocation_txtHend = WarehouseLocation_txtLine.IndexOf(Positionlast, WarehouseLocation_txtHstart);
                        string WarehouseLocation_txtaH = WarehouseLocation_txtLine.ToUpper().Substring(WarehouseLocation_txtHstart, WarehouseLocation_txtHend - WarehouseLocation_txtHstart);
                        int WarehouseLocation_txtH = Int32.Parse(WarehouseLocation_txtaH);
                        float WarehouseLocation_txtFH = float.Parse(WarehouseLocation_txtaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find WarehouseLocation_txt Color
                        int WarehouseLocation_txtCstart1 = WarehouseLocation_txtLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int WarehouseLocation_txtCend1 = WarehouseLocation_txtLine.IndexOf(Positionlast, WarehouseLocation_txtCstart1);
                        string WarehouseLocation_txtCH = WarehouseLocation_txtLine.ToUpper().Substring(WarehouseLocation_txtCstart1, WarehouseLocation_txtCend1 - WarehouseLocation_txtCstart1);
                        string WarehouseLocation_txtCF = WarehouseLocation_txtCH.Replace("0x", "#");
                        Color WarehouseLocation_txtC = ColorTranslator.FromHtml(WarehouseLocation_txtCF);
                        if (WarehouseLocation_txtCF == "000000")
                        {
                            WarehouseLocation_txtC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var WarehouseLocation_txtbrush = new SolidBrush(Color.FromArgb(WarehouseLocation_txtC.A, WarehouseLocation_txtC.R, WarehouseLocation_txtC.G, WarehouseLocation_txtC.B));
                        //find WarehouseLocation_txt Font
                        int WarehouseLocation_txtFontstart = WarehouseLocation_txtLine.IndexOf(CFont) + CFont.Length;
                        int WarehouseLocation_txtFontend = WarehouseLocation_txtLine.IndexOf(Positionlast, WarehouseLocation_txtFontstart);
                        string WarehouseLocation_txtaFont = WarehouseLocation_txtLine.ToUpper().Substring(WarehouseLocation_txtFontstart, WarehouseLocation_txtFontend - WarehouseLocation_txtFontstart);
                        Font WarehouseLocation_txtFont;
                        if (WarehouseLocation_txtaFont == "1")
                        {
                            WarehouseLocation_txtFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (WarehouseLocation_txtaFont == "2")
                        {
                            WarehouseLocation_txtFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (WarehouseLocation_txtaFont == "3")
                        {
                            WarehouseLocation_txtFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            WarehouseLocation_txtFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //WarehouseLocation
                        //find WarehouseLocation
                        string WarehouseLocationLine = spliter[22].ToString();
                        int WarehouseLocationstart1 = WarehouseLocationLine.IndexOf(Search1) + Search1.Length;
                        int WarehouseLocationend1 = WarehouseLocationLine.IndexOf(Search2, WarehouseLocationstart1);
                        string WarehouseLocation = WarehouseLocationLine.ToUpper().Substring(WarehouseLocationstart1, WarehouseLocationend1 - WarehouseLocationstart1);
                        string FinalWarehouseLocation = String.Join(" ", WarehouseLocation.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find WarehouseLocation locX
                        int WarehouseLocationXstart = WarehouseLocationLine.IndexOf(Positionx) + Positionx.Length;
                        int WarehouseLocationXend = WarehouseLocationLine.IndexOf(Positionlast, WarehouseLocationXstart);
                        string WarehouseLocationaX = WarehouseLocationLine.ToUpper().Substring(WarehouseLocationXstart, WarehouseLocationXend - WarehouseLocationXstart);
                        int WarehouseLocationX = Int32.Parse(WarehouseLocationaX);
                        //find WarehouseLocation locXY
                        int WarehouseLocationYstart = WarehouseLocationLine.IndexOf(Positiony) + Positiony.Length;
                        int WarehouseLocationYend = WarehouseLocationLine.IndexOf(Positionlast, WarehouseLocationYstart);
                        string WarehouseLocationaY = WarehouseLocationLine.ToUpper().Substring(WarehouseLocationYstart, WarehouseLocationYend - WarehouseLocationYstart);
                        int WarehouseLocationY = Int32.Parse(WarehouseLocationaY);
                        //Find WarehouseLocation width
                        int WarehouseLocationWstart = WarehouseLocationLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int WarehouseLocationWend = WarehouseLocationLine.IndexOf(Positionlast, WarehouseLocationWstart);
                        string WarehouseLocationaW = WarehouseLocationLine.ToUpper().Substring(WarehouseLocationWstart, WarehouseLocationWend - WarehouseLocationWstart);
                        int WarehouseLocationW = Int32.Parse(WarehouseLocationaW);
                        float WarehouseLocationFW = float.Parse(WarehouseLocationaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find WarehouseLocation height
                        int WarehouseLocationHstart = WarehouseLocationLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int WarehouseLocationHend = WarehouseLocationLine.IndexOf(Positionlast, WarehouseLocationHstart);
                        string WarehouseLocationaH = WarehouseLocationLine.ToUpper().Substring(WarehouseLocationHstart, WarehouseLocationHend - WarehouseLocationHstart);
                        int WarehouseLocationH = Int32.Parse(WarehouseLocationaH);
                        float WarehouseLocationFH = float.Parse(WarehouseLocationaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find WarehouseLocation Color
                        int WarehouseLocationCstart1 = WarehouseLocationLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int WarehouseLocationCend1 = WarehouseLocationLine.IndexOf(Positionlast, WarehouseLocationCstart1);
                        string WarehouseLocationCH = WarehouseLocationLine.ToUpper().Substring(WarehouseLocationCstart1, WarehouseLocationCend1 - WarehouseLocationCstart1);
                        string WarehouseLocationCF = WarehouseLocationCH.Replace("0x", "#");
                        Color WarehouseLocationC = ColorTranslator.FromHtml(WarehouseLocationCF);
                        if (WarehouseLocationCF == "000000")
                        {
                            WarehouseLocationC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var WarehouseLocationbrush = new SolidBrush(Color.FromArgb(WarehouseLocationC.A, WarehouseLocationC.R, WarehouseLocationC.G, WarehouseLocationC.B));
                        //find WarehouseLocation Font
                        int WarehouseLocationFontstart = WarehouseLocationLine.IndexOf(CFont) + CFont.Length;
                        int WarehouseLocationFontend = WarehouseLocationLine.IndexOf(Positionlast, WarehouseLocationFontstart);
                        string WarehouseLocationaFont = WarehouseLocationLine.ToUpper().Substring(WarehouseLocationFontstart, WarehouseLocationFontend - WarehouseLocationFontstart);
                        Font WarehouseLocationFont;
                        if (WarehouseLocationaFont == "1")
                        {
                            WarehouseLocationFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (WarehouseLocationaFont == "2")
                        {
                            WarehouseLocationFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (WarehouseLocationaFont == "3")
                        {
                            WarehouseLocationFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            WarehouseLocationFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //FixedAssetLocation_txt
                        //find FixedAssetLocation_txt
                        string FixedAssetLocation_txtLine = spliter[23].ToString();
                        int FixedAssetLocation_txtstart1 = FixedAssetLocation_txtLine.IndexOf(Search1) + Search1.Length;
                        int FixedAssetLocation_txtend1 = FixedAssetLocation_txtLine.IndexOf(Search2, FixedAssetLocation_txtstart1);
                        string FixedAssetLocation_txt = FixedAssetLocation_txtLine.ToUpper().Substring(FixedAssetLocation_txtstart1, FixedAssetLocation_txtend1 - FixedAssetLocation_txtstart1);
                        string FinalFixedAssetLocation_txt = String.Join(" ", FixedAssetLocation_txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find FixedAssetLocation_txt locX
                        int FixedAssetLocation_txtXstart = FixedAssetLocation_txtLine.IndexOf(Positionx) + Positionx.Length;
                        int FixedAssetLocation_txtXend = FixedAssetLocation_txtLine.IndexOf(Positionlast, FixedAssetLocation_txtXstart);
                        string FixedAssetLocation_txtaX = FixedAssetLocation_txtLine.ToUpper().Substring(FixedAssetLocation_txtXstart, FixedAssetLocation_txtXend - FixedAssetLocation_txtXstart);
                        int FixedAssetLocation_txtX = Int32.Parse(FixedAssetLocation_txtaX);
                        //find FixedAssetLocation_txt locXY
                        int FixedAssetLocation_txtYstart = FixedAssetLocation_txtLine.IndexOf(Positiony) + Positiony.Length;
                        int FixedAssetLocation_txtYend = FixedAssetLocation_txtLine.IndexOf(Positionlast, FixedAssetLocation_txtYstart);
                        string FixedAssetLocation_txtaY = FixedAssetLocation_txtLine.ToUpper().Substring(FixedAssetLocation_txtYstart, FixedAssetLocation_txtYend - FixedAssetLocation_txtYstart);
                        int FixedAssetLocation_txtY = Int32.Parse(FixedAssetLocation_txtaY);
                        //Find FixedAssetLocation_txt width
                        int FixedAssetLocation_txtWstart = FixedAssetLocation_txtLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int FixedAssetLocation_txtWend = FixedAssetLocation_txtLine.IndexOf(Positionlast, FixedAssetLocation_txtWstart);
                        string FixedAssetLocation_txtaW = FixedAssetLocation_txtLine.ToUpper().Substring(FixedAssetLocation_txtWstart, FixedAssetLocation_txtWend - FixedAssetLocation_txtWstart);
                        int FixedAssetLocation_txtW = Int32.Parse(FixedAssetLocation_txtaW);
                        float FixedAssetLocation_txtFW = float.Parse(FixedAssetLocation_txtaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find FixedAssetLocation_txt height
                        int FixedAssetLocation_txtHstart = FixedAssetLocation_txtLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int FixedAssetLocation_txtHend = FixedAssetLocation_txtLine.IndexOf(Positionlast, FixedAssetLocation_txtHstart);
                        string FixedAssetLocation_txtaH = FixedAssetLocation_txtLine.ToUpper().Substring(FixedAssetLocation_txtHstart, FixedAssetLocation_txtHend - FixedAssetLocation_txtHstart);
                        int FixedAssetLocation_txtH = Int32.Parse(FixedAssetLocation_txtaH);
                        float FixedAssetLocation_txtFH = float.Parse(FixedAssetLocation_txtaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find FixedAssetLocation_txt Color
                        int FixedAssetLocation_txtCstart1 = FixedAssetLocation_txtLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int FixedAssetLocation_txtCend1 = FixedAssetLocation_txtLine.IndexOf(Positionlast, FixedAssetLocation_txtCstart1);
                        string FixedAssetLocation_txtCH = FixedAssetLocation_txtLine.ToUpper().Substring(FixedAssetLocation_txtCstart1, FixedAssetLocation_txtCend1 - FixedAssetLocation_txtCstart1);
                        string FixedAssetLocation_txtCF = FixedAssetLocation_txtCH.Replace("0x", "#");
                        Color FixedAssetLocation_txtC = ColorTranslator.FromHtml(FixedAssetLocation_txtCF);
                        if (FixedAssetLocation_txtCF == "000000")
                        {
                            FixedAssetLocation_txtC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var FixedAssetLocation_txtbrush = new SolidBrush(Color.FromArgb(FixedAssetLocation_txtC.A, FixedAssetLocation_txtC.R, FixedAssetLocation_txtC.G, FixedAssetLocation_txtC.B));
                        //find FixedAssetLocation_txt Font
                        int FixedAssetLocation_txtFontstart = FixedAssetLocation_txtLine.IndexOf(CFont) + CFont.Length;
                        int FixedAssetLocation_txtFontend = FixedAssetLocation_txtLine.IndexOf(Positionlast, FixedAssetLocation_txtFontstart);
                        string FixedAssetLocation_txtaFont = FixedAssetLocation_txtLine.ToUpper().Substring(FixedAssetLocation_txtFontstart, FixedAssetLocation_txtFontend - FixedAssetLocation_txtFontstart);
                        Font FixedAssetLocation_txtFont;
                        if (FixedAssetLocation_txtaFont == "1")
                        {
                            FixedAssetLocation_txtFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (FixedAssetLocation_txtaFont == "2")
                        {
                            FixedAssetLocation_txtFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (FixedAssetLocation_txtaFont == "3")
                        {
                            FixedAssetLocation_txtFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            FixedAssetLocation_txtFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //FixedAssetLocation
                        //find FixedAssetLocation
                        string FixedAssetLocationLine = spliter[24].ToString();
                        int FixedAssetLocationstart1 = FixedAssetLocationLine.IndexOf(Search1) + Search1.Length;
                        int FixedAssetLocationend1 = FixedAssetLocationLine.IndexOf(Search2, FixedAssetLocationstart1);
                        string FixedAssetLocation = FixedAssetLocationLine.ToUpper().Substring(FixedAssetLocationstart1, FixedAssetLocationend1 - FixedAssetLocationstart1);
                        string FinalFixedAssetLocation = String.Join(" ", FixedAssetLocation.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find FixedAssetLocation locX
                        int FixedAssetLocationXstart = FixedAssetLocationLine.IndexOf(Positionx) + Positionx.Length;
                        int FixedAssetLocationXend = FixedAssetLocationLine.IndexOf(Positionlast, FixedAssetLocationXstart);
                        string FixedAssetLocationaX = FixedAssetLocationLine.ToUpper().Substring(FixedAssetLocationXstart, FixedAssetLocationXend - FixedAssetLocationXstart);
                        int FixedAssetLocationX = Int32.Parse(FixedAssetLocationaX);
                        //find FixedAssetLocation locXY
                        int FixedAssetLocationYstart = FixedAssetLocationLine.IndexOf(Positiony) + Positiony.Length;
                        int FixedAssetLocationYend = FixedAssetLocationLine.IndexOf(Positionlast, FixedAssetLocationYstart);
                        string FixedAssetLocationaY = FixedAssetLocationLine.ToUpper().Substring(FixedAssetLocationYstart, FixedAssetLocationYend - FixedAssetLocationYstart);
                        int FixedAssetLocationY = Int32.Parse(FixedAssetLocationaY);
                        //Find FixedAssetLocation width
                        int FixedAssetLocationWstart = FixedAssetLocationLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int FixedAssetLocationWend = FixedAssetLocationLine.IndexOf(Positionlast, FixedAssetLocationWstart);
                        string FixedAssetLocationaW = FixedAssetLocationLine.ToUpper().Substring(FixedAssetLocationWstart, FixedAssetLocationWend - FixedAssetLocationWstart);
                        int FixedAssetLocationW = Int32.Parse(FixedAssetLocationaW);
                        float FixedAssetLocationFW = float.Parse(FixedAssetLocationaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find FixedAssetLocation height
                        int FixedAssetLocationHstart = FixedAssetLocationLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int FixedAssetLocationHend = FixedAssetLocationLine.IndexOf(Positionlast, FixedAssetLocationHstart);
                        string FixedAssetLocationaH = FixedAssetLocationLine.ToUpper().Substring(FixedAssetLocationHstart, FixedAssetLocationHend - FixedAssetLocationHstart);
                        int FixedAssetLocationH = Int32.Parse(FixedAssetLocationaH);
                        float FixedAssetLocationFH = float.Parse(FixedAssetLocationaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find FixedAssetLocation Color
                        int FixedAssetLocationCstart1 = FixedAssetLocationLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int FixedAssetLocationCend1 = FixedAssetLocationLine.IndexOf(Positionlast, FixedAssetLocationCstart1);
                        string FixedAssetLocationCH = FixedAssetLocationLine.ToUpper().Substring(FixedAssetLocationCstart1, FixedAssetLocationCend1 - FixedAssetLocationCstart1);
                        string FixedAssetLocationCF = FixedAssetLocationCH.Replace("0x", "#");
                        Color FixedAssetLocationC = ColorTranslator.FromHtml(FixedAssetLocationCF);
                        if (FixedAssetLocationCF == "000000")
                        {
                            FixedAssetLocationC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var FixedAssetLocationbrush = new SolidBrush(Color.FromArgb(FixedAssetLocationC.A, FixedAssetLocationC.R, FixedAssetLocationC.G, FixedAssetLocationC.B));
                        //find FixedAssetLocation Font
                        int FixedAssetLocationFontstart = FixedAssetLocationLine.IndexOf(CFont) + CFont.Length;
                        int FixedAssetLocationFontend = FixedAssetLocationLine.IndexOf(Positionlast, FixedAssetLocationFontstart);
                        string FixedAssetLocationaFont = FixedAssetLocationLine.ToUpper().Substring(FixedAssetLocationFontstart, FixedAssetLocationFontend - FixedAssetLocationFontstart);
                        Font FixedAssetLocationFont;
                        if (FixedAssetLocationaFont == "1")
                        {
                            FixedAssetLocationFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (FixedAssetLocationaFont == "2")
                        {
                            FixedAssetLocationFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (FixedAssetLocationaFont == "3")
                        {
                            FixedAssetLocationFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            FixedAssetLocationFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }


                        //MFGDate_txt
                        //find MFGDate_txt
                        string MFGDate_txtLine = spliter[25].ToString();
                        int MFGDate_txtstart1 = MFGDate_txtLine.IndexOf(Search1) + Search1.Length;
                        int MFGDate_txtend1 = MFGDate_txtLine.IndexOf(Search2, MFGDate_txtstart1);
                        string MFGDate_txt = MFGDate_txtLine.ToUpper().Substring(MFGDate_txtstart1, MFGDate_txtend1 - MFGDate_txtstart1);
                        string FinalMFGDate_txt = String.Join(" ", MFGDate_txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find MFGDate_txt locX
                        int MFGDate_txtXstart = MFGDate_txtLine.IndexOf(Positionx) + Positionx.Length;
                        int MFGDate_txtXend = MFGDate_txtLine.IndexOf(Positionlast, MFGDate_txtXstart);
                        string MFGDate_txtaX = MFGDate_txtLine.ToUpper().Substring(MFGDate_txtXstart, MFGDate_txtXend - MFGDate_txtXstart);
                        int MFGDate_txtX = Int32.Parse(MFGDate_txtaX);
                        //find MFGDate_txt locXY
                        int MFGDate_txtYstart = MFGDate_txtLine.IndexOf(Positiony) + Positiony.Length;
                        int MFGDate_txtYend = MFGDate_txtLine.IndexOf(Positionlast, MFGDate_txtYstart);
                        string MFGDate_txtaY = MFGDate_txtLine.ToUpper().Substring(MFGDate_txtYstart, MFGDate_txtYend - MFGDate_txtYstart);
                        int MFGDate_txtY = Int32.Parse(MFGDate_txtaY);
                        //Find MFGDate_txt width
                        int MFGDate_txtWstart = MFGDate_txtLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int MFGDate_txtWend = MFGDate_txtLine.IndexOf(Positionlast, MFGDate_txtWstart);
                        string MFGDate_txtaW = MFGDate_txtLine.ToUpper().Substring(MFGDate_txtWstart, MFGDate_txtWend - MFGDate_txtWstart);
                        int MFGDate_txtW = Int32.Parse(MFGDate_txtaW);
                        float MFGDate_txtFW = float.Parse(MFGDate_txtaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find MFGDate_txt height
                        int MFGDate_txtHstart = MFGDate_txtLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int MFGDate_txtHend = MFGDate_txtLine.IndexOf(Positionlast, MFGDate_txtHstart);
                        string MFGDate_txtaH = MFGDate_txtLine.ToUpper().Substring(MFGDate_txtHstart, MFGDate_txtHend - MFGDate_txtHstart);
                        int MFGDate_txtH = Int32.Parse(MFGDate_txtaH);
                        float MFGDate_txtFH = float.Parse(MFGDate_txtaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find MFGDate_txt Color
                        int MFGDate_txtCstart1 = MFGDate_txtLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int MFGDate_txtCend1 = MFGDate_txtLine.IndexOf(Positionlast, MFGDate_txtCstart1);
                        string MFGDate_txtCH = MFGDate_txtLine.ToUpper().Substring(MFGDate_txtCstart1, MFGDate_txtCend1 - MFGDate_txtCstart1);
                        string MFGDate_txtCF = MFGDate_txtCH.Replace("0x", "#");
                        Color MFGDate_txtC = ColorTranslator.FromHtml(MFGDate_txtCF);
                        if (MFGDate_txtCF == "000000")
                        {
                            MFGDate_txtC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var MFGDate_txtbrush = new SolidBrush(Color.FromArgb(MFGDate_txtC.A, MFGDate_txtC.R, MFGDate_txtC.G, MFGDate_txtC.B));
                        //find MFGDate_txt Font
                        int MFGDate_txtFontstart = MFGDate_txtLine.IndexOf(CFont) + CFont.Length;
                        int MFGDate_txtFontend = MFGDate_txtLine.IndexOf(Positionlast, MFGDate_txtFontstart);
                        string MFGDate_txtaFont = MFGDate_txtLine.ToUpper().Substring(MFGDate_txtFontstart, MFGDate_txtFontend - MFGDate_txtFontstart);
                        Font MFGDate_txtFont;
                        if (MFGDate_txtaFont == "1")
                        {
                            MFGDate_txtFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (MFGDate_txtaFont == "2")
                        {
                            MFGDate_txtFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (MFGDate_txtaFont == "3")
                        {
                            MFGDate_txtFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            MFGDate_txtFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }
                        //MFGDate
                        //find MFGDate
                        string MFGDateLine = spliter[26].ToString();
                        int MFGDatestart1 = MFGDateLine.IndexOf(Search1) + Search1.Length;
                        int MFGDateend1 = MFGDateLine.IndexOf(Search2, MFGDatestart1);
                        string MFGDate = MFGDateLine.ToUpper().Substring(MFGDatestart1, MFGDateend1 - MFGDatestart1);
                        string FinalMFGDate = String.Join(" ", MFGDate.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find MFGDate locX
                        int MFGDateXstart = MFGDateLine.IndexOf(Positionx) + Positionx.Length;
                        int MFGDateXend = MFGDateLine.IndexOf(Positionlast, MFGDateXstart);
                        string MFGDateaX = MFGDateLine.ToUpper().Substring(MFGDateXstart, MFGDateXend - MFGDateXstart);
                        int MFGDateX = Int32.Parse(MFGDateaX);
                        //find MFGDate locXY
                        int MFGDateYstart = MFGDateLine.IndexOf(Positiony) + Positiony.Length;
                        int MFGDateYend = MFGDateLine.IndexOf(Positionlast, MFGDateYstart);
                        string MFGDateaY = MFGDateLine.ToUpper().Substring(MFGDateYstart, MFGDateYend - MFGDateYstart);
                        int MFGDateY = Int32.Parse(MFGDateaY);
                        //Find MFGDate width
                        int MFGDateWstart = MFGDateLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int MFGDateWend = MFGDateLine.IndexOf(Positionlast, MFGDateWstart);
                        string MFGDateaW = MFGDateLine.ToUpper().Substring(MFGDateWstart, MFGDateWend - MFGDateWstart);
                        int MFGDateW = Int32.Parse(MFGDateaW);
                        float MFGDateFW = float.Parse(MFGDateaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find MFGDate height
                        int MFGDateHstart = MFGDateLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int MFGDateHend = MFGDateLine.IndexOf(Positionlast, MFGDateHstart);
                        string MFGDateaH = MFGDateLine.ToUpper().Substring(MFGDateHstart, MFGDateHend - MFGDateHstart);
                        int MFGDateH = Int32.Parse(MFGDateaH);
                        float MFGDateFH = float.Parse(MFGDateaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find MFGDate Color
                        int MFGDateCstart1 = MFGDateLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int MFGDateCend1 = MFGDateLine.IndexOf(Positionlast, MFGDateCstart1);
                        string MFGDateCH = MFGDateLine.ToUpper().Substring(MFGDateCstart1, MFGDateCend1 - MFGDateCstart1);
                        string MFGDateCF = MFGDateCH.Replace("0x", "#");
                        Color MFGDateC = ColorTranslator.FromHtml(MFGDateCF);
                        if (MFGDateCF == "000000")
                        {
                            MFGDateC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var MFGDatebrush = new SolidBrush(Color.FromArgb(MFGDateC.A, MFGDateC.R, MFGDateC.G, MFGDateC.B));
                        //find MFGDate Font
                        int MFGDateFontstart = MFGDateLine.IndexOf(CFont) + CFont.Length;
                        int MFGDateFontend = MFGDateLine.IndexOf(Positionlast, MFGDateFontstart);
                        string MFGDateaFont = MFGDateLine.ToUpper().Substring(MFGDateFontstart, MFGDateFontend - MFGDateFontstart);
                        Font MFGDateFont;
                        if (MFGDateaFont == "1")
                        {
                            MFGDateFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (MFGDateaFont == "2")
                        {
                            MFGDateFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (MFGDateaFont == "3")
                        {
                            MFGDateFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            MFGDateFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //AdmissionDate_txt
                        //find AdmissionDate_txt
                        string AdmissionDate_txtLine = spliter[27].ToString();
                        int AdmissionDate_txtstart1 = AdmissionDate_txtLine.IndexOf(Search1) + Search1.Length;
                        int AdmissionDate_txtend1 = AdmissionDate_txtLine.IndexOf(Search2, AdmissionDate_txtstart1);
                        string AdmissionDate_txt = AdmissionDate_txtLine.ToUpper().Substring(AdmissionDate_txtstart1, AdmissionDate_txtend1 - AdmissionDate_txtstart1);
                        string FinalAdmissionDate_txt = String.Join(" ", AdmissionDate_txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find AdmissionDate_txt locX
                        int AdmissionDate_txtXstart = AdmissionDate_txtLine.IndexOf(Positionx) + Positionx.Length;
                        int AdmissionDate_txtXend = AdmissionDate_txtLine.IndexOf(Positionlast, AdmissionDate_txtXstart);
                        string AdmissionDate_txtaX = AdmissionDate_txtLine.ToUpper().Substring(AdmissionDate_txtXstart, AdmissionDate_txtXend - AdmissionDate_txtXstart);
                        int AdmissionDate_txtX = Int32.Parse(AdmissionDate_txtaX);
                        //find AdmissionDate_txt locXY
                        int AdmissionDate_txtYstart = AdmissionDate_txtLine.IndexOf(Positiony) + Positiony.Length;
                        int AdmissionDate_txtYend = AdmissionDate_txtLine.IndexOf(Positionlast, AdmissionDate_txtYstart);
                        string AdmissionDate_txtaY = AdmissionDate_txtLine.ToUpper().Substring(AdmissionDate_txtYstart, AdmissionDate_txtYend - AdmissionDate_txtYstart);
                        int AdmissionDate_txtY = Int32.Parse(AdmissionDate_txtaY);
                        //Find AdmissionDate_txt width
                        int AdmissionDate_txtWstart = AdmissionDate_txtLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int AdmissionDate_txtWend = AdmissionDate_txtLine.IndexOf(Positionlast, AdmissionDate_txtWstart);
                        string AdmissionDate_txtaW = AdmissionDate_txtLine.ToUpper().Substring(AdmissionDate_txtWstart, AdmissionDate_txtWend - AdmissionDate_txtWstart);
                        int AdmissionDate_txtW = Int32.Parse(AdmissionDate_txtaW);
                        float AdmissionDate_txtFW = float.Parse(AdmissionDate_txtaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find AdmissionDate_txt height
                        int AdmissionDate_txtHstart = AdmissionDate_txtLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int AdmissionDate_txtHend = AdmissionDate_txtLine.IndexOf(Positionlast, AdmissionDate_txtHstart);
                        string AdmissionDate_txtaH = AdmissionDate_txtLine.ToUpper().Substring(AdmissionDate_txtHstart, AdmissionDate_txtHend - AdmissionDate_txtHstart);
                        int AdmissionDate_txtH = Int32.Parse(AdmissionDate_txtaH);
                        float AdmissionDate_txtFH = float.Parse(AdmissionDate_txtaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find AdmissionDate_txt Color
                        int AdmissionDate_txtCstart1 = AdmissionDate_txtLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int AdmissionDate_txtCend1 = AdmissionDate_txtLine.IndexOf(Positionlast, AdmissionDate_txtCstart1);
                        string AdmissionDate_txtCH = AdmissionDate_txtLine.ToUpper().Substring(AdmissionDate_txtCstart1, AdmissionDate_txtCend1 - AdmissionDate_txtCstart1);
                        string AdmissionDate_txtCF = AdmissionDate_txtCH.Replace("0x", "#");
                        Color AdmissionDate_txtC = ColorTranslator.FromHtml(AdmissionDate_txtCF);
                        if (AdmissionDate_txtCF == "000000")
                        {
                            AdmissionDate_txtC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var AdmissionDate_txtbrush = new SolidBrush(Color.FromArgb(AdmissionDate_txtC.A, AdmissionDate_txtC.R, AdmissionDate_txtC.G, AdmissionDate_txtC.B));
                        //find AdmissionDate_txt Font
                        int AdmissionDate_txtFontstart = AdmissionDate_txtLine.IndexOf(CFont) + CFont.Length;
                        int AdmissionDate_txtFontend = AdmissionDate_txtLine.IndexOf(Positionlast, AdmissionDate_txtFontstart);
                        string AdmissionDate_txtaFont = AdmissionDate_txtLine.ToUpper().Substring(AdmissionDate_txtFontstart, AdmissionDate_txtFontend - AdmissionDate_txtFontstart);
                        Font AdmissionDate_txtFont;
                        if (AdmissionDate_txtaFont == "1")
                        {
                            AdmissionDate_txtFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (AdmissionDate_txtaFont == "2")
                        {
                            AdmissionDate_txtFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (AdmissionDate_txtaFont == "3")
                        {
                            AdmissionDate_txtFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            AdmissionDate_txtFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //AdmissionDate
                        //find AdmissionDate
                        string AdmissionDateLine = spliter[28].ToString();
                        int AdmissionDatestart1 = AdmissionDateLine.IndexOf(Search1) + Search1.Length;
                        int AdmissionDateend1 = AdmissionDateLine.IndexOf(Search2, AdmissionDatestart1);
                        string AdmissionDate = AdmissionDateLine.ToUpper().Substring(AdmissionDatestart1, AdmissionDateend1 - AdmissionDatestart1);
                        string FinalAdmissionDate = String.Join(" ", AdmissionDate.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find AdmissionDate locX
                        int AdmissionDateXstart = AdmissionDateLine.IndexOf(Positionx) + Positionx.Length;
                        int AdmissionDateXend = AdmissionDateLine.IndexOf(Positionlast, AdmissionDateXstart);
                        string AdmissionDateaX = AdmissionDateLine.ToUpper().Substring(AdmissionDateXstart, AdmissionDateXend - AdmissionDateXstart);
                        int AdmissionDateX = Int32.Parse(AdmissionDateaX);
                        //find AdmissionDate locXY
                        int AdmissionDateYstart = AdmissionDateLine.IndexOf(Positiony) + Positiony.Length;
                        int AdmissionDateYend = AdmissionDateLine.IndexOf(Positionlast, AdmissionDateYstart);
                        string AdmissionDateaY = AdmissionDateLine.ToUpper().Substring(AdmissionDateYstart, AdmissionDateYend - AdmissionDateYstart);
                        int AdmissionDateY = Int32.Parse(AdmissionDateaY);
                        //Find AdmissionDate width
                        int AdmissionDateWstart = AdmissionDateLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int AdmissionDateWend = AdmissionDateLine.IndexOf(Positionlast, AdmissionDateWstart);
                        string AdmissionDateaW = AdmissionDateLine.ToUpper().Substring(AdmissionDateWstart, AdmissionDateWend - AdmissionDateWstart);
                        int AdmissionDateW = Int32.Parse(AdmissionDateaW);
                        float AdmissionDateFW = float.Parse(AdmissionDateaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find AdmissionDate height
                        int AdmissionDateHstart = AdmissionDateLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int AdmissionDateHend = AdmissionDateLine.IndexOf(Positionlast, AdmissionDateHstart);
                        string AdmissionDateaH = AdmissionDateLine.ToUpper().Substring(AdmissionDateHstart, AdmissionDateHend - AdmissionDateHstart);
                        int AdmissionDateH = Int32.Parse(AdmissionDateaH);
                        float AdmissionDateFH = float.Parse(AdmissionDateaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find AdmissionDate Color
                        int AdmissionDateCstart1 = AdmissionDateLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int AdmissionDateCend1 = AdmissionDateLine.IndexOf(Positionlast, AdmissionDateCstart1);
                        string AdmissionDateCH = AdmissionDateLine.ToUpper().Substring(AdmissionDateCstart1, AdmissionDateCend1 - AdmissionDateCstart1);
                        string AdmissionDateCF = AdmissionDateCH.Replace("0x", "#");
                        Color AdmissionDateC = ColorTranslator.FromHtml(AdmissionDateCF);
                        if (AdmissionDateCF == "000000")
                        {
                            AdmissionDateC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var AdmissionDatebrush = new SolidBrush(Color.FromArgb(AdmissionDateC.A, AdmissionDateC.R, AdmissionDateC.G, AdmissionDateC.B));
                        //find AdmissionDate Font
                        int AdmissionDateFontstart = AdmissionDateLine.IndexOf(CFont) + CFont.Length;
                        int AdmissionDateFontend = AdmissionDateLine.IndexOf(Positionlast, AdmissionDateFontstart);
                        string AdmissionDateaFont = AdmissionDateLine.ToUpper().Substring(AdmissionDateFontstart, AdmissionDateFontend - AdmissionDateFontstart);
                        Font AdmissionDateFont;
                        if (AdmissionDateaFont == "1")
                        {
                            AdmissionDateFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (AdmissionDateaFont == "2")
                        {
                            AdmissionDateFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (AdmissionDateaFont == "3")
                        {
                            AdmissionDateFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            AdmissionDateFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //Lifespan_txt
                        //find Lifespan_txt
                        string Lifespan_txtLine = spliter[29].ToString();
                        int Lifespan_txtstart1 = Lifespan_txtLine.IndexOf(Search1) + Search1.Length;
                        int Lifespan_txtend1 = Lifespan_txtLine.IndexOf(Search2, Lifespan_txtstart1);
                        string Lifespan_txt = Lifespan_txtLine.ToUpper().Substring(Lifespan_txtstart1, Lifespan_txtend1 - Lifespan_txtstart1);
                        string FinalLifespan_txt = String.Join(" ", Lifespan_txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find Lifespan_txt locX
                        int Lifespan_txtXstart = Lifespan_txtLine.IndexOf(Positionx) + Positionx.Length;
                        int Lifespan_txtXend = Lifespan_txtLine.IndexOf(Positionlast, Lifespan_txtXstart);
                        string Lifespan_txtaX = Lifespan_txtLine.ToUpper().Substring(Lifespan_txtXstart, Lifespan_txtXend - Lifespan_txtXstart);
                        int Lifespan_txtX = Int32.Parse(Lifespan_txtaX);
                        //find Lifespan_txt locXY
                        int Lifespan_txtYstart = Lifespan_txtLine.IndexOf(Positiony) + Positiony.Length;
                        int Lifespan_txtYend = Lifespan_txtLine.IndexOf(Positionlast, Lifespan_txtYstart);
                        string Lifespan_txtaY = Lifespan_txtLine.ToUpper().Substring(Lifespan_txtYstart, Lifespan_txtYend - Lifespan_txtYstart);
                        int Lifespan_txtY = Int32.Parse(Lifespan_txtaY);
                        //Find Lifespan_txt width
                        int Lifespan_txtWstart = Lifespan_txtLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int Lifespan_txtWend = Lifespan_txtLine.IndexOf(Positionlast, Lifespan_txtWstart);
                        string Lifespan_txtaW = Lifespan_txtLine.ToUpper().Substring(Lifespan_txtWstart, Lifespan_txtWend - Lifespan_txtWstart);
                        int Lifespan_txtW = Int32.Parse(Lifespan_txtaW);
                        float Lifespan_txtFW = float.Parse(Lifespan_txtaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find Lifespan_txt height
                        int Lifespan_txtHstart = Lifespan_txtLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int Lifespan_txtHend = Lifespan_txtLine.IndexOf(Positionlast, Lifespan_txtHstart);
                        string Lifespan_txtaH = Lifespan_txtLine.ToUpper().Substring(Lifespan_txtHstart, Lifespan_txtHend - Lifespan_txtHstart);
                        int Lifespan_txtH = Int32.Parse(Lifespan_txtaH);
                        float Lifespan_txtFH = float.Parse(Lifespan_txtaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find Lifespan_txt Color
                        int Lifespan_txtCstart1 = Lifespan_txtLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int Lifespan_txtCend1 = Lifespan_txtLine.IndexOf(Positionlast, Lifespan_txtCstart1);
                        string Lifespan_txtCH = Lifespan_txtLine.ToUpper().Substring(Lifespan_txtCstart1, Lifespan_txtCend1 - Lifespan_txtCstart1);
                        string Lifespan_txtCF = Lifespan_txtCH.Replace("0x", "#");
                        Color Lifespan_txtC = ColorTranslator.FromHtml(Lifespan_txtCF);
                        if (Lifespan_txtCF == "000000")
                        {
                            Lifespan_txtC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var Lifespan_txtbrush = new SolidBrush(Color.FromArgb(Lifespan_txtC.A, Lifespan_txtC.R, Lifespan_txtC.G, Lifespan_txtC.B));
                        //find Lifespan_txt Font
                        int Lifespan_txtFontstart = Lifespan_txtLine.IndexOf(CFont) + CFont.Length;
                        int Lifespan_txtFontend = Lifespan_txtLine.IndexOf(Positionlast, Lifespan_txtFontstart);
                        string Lifespan_txtaFont = Lifespan_txtLine.ToUpper().Substring(Lifespan_txtFontstart, Lifespan_txtFontend - Lifespan_txtFontstart);
                        Font Lifespan_txtFont;
                        if (Lifespan_txtaFont == "1")
                        {
                            Lifespan_txtFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (Lifespan_txtaFont == "2")
                        {
                            Lifespan_txtFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (Lifespan_txtaFont == "3")
                        {
                            Lifespan_txtFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            Lifespan_txtFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //Lifespan
                        //find Lifespan
                        string LifespanLine = spliter[30].ToString();
                        int Lifespanstart1 = LifespanLine.IndexOf(Search1) + Search1.Length;
                        int Lifespanend1 = LifespanLine.IndexOf(Search2, Lifespanstart1);
                        string Lifespan = LifespanLine.ToUpper().Substring(Lifespanstart1, Lifespanend1 - Lifespanstart1);
                        string FinalLifespan = String.Join(" ", Lifespan.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find Lifespan locX
                        int LifespanXstart = LifespanLine.IndexOf(Positionx) + Positionx.Length;
                        int LifespanXend = LifespanLine.IndexOf(Positionlast, LifespanXstart);
                        string LifespanaX = LifespanLine.ToUpper().Substring(LifespanXstart, LifespanXend - LifespanXstart);
                        int LifespanX = Int32.Parse(LifespanaX);
                        //find Lifespan locXY
                        int LifespanYstart = LifespanLine.IndexOf(Positiony) + Positiony.Length;
                        int LifespanYend = LifespanLine.IndexOf(Positionlast, LifespanYstart);
                        string LifespanaY = LifespanLine.ToUpper().Substring(LifespanYstart, LifespanYend - LifespanYstart);
                        int LifespanY = Int32.Parse(LifespanaY);
                        //Find Lifespan width
                        int LifespanWstart = LifespanLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int LifespanWend = LifespanLine.IndexOf(Positionlast, LifespanWstart);
                        string LifespanaW = LifespanLine.ToUpper().Substring(LifespanWstart, LifespanWend - LifespanWstart);
                        int LifespanW = Int32.Parse(LifespanaW);
                        float LifespanFW = float.Parse(LifespanaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find Lifespan height
                        int LifespanHstart = LifespanLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int LifespanHend = LifespanLine.IndexOf(Positionlast, LifespanHstart);
                        string LifespanaH = LifespanLine.ToUpper().Substring(LifespanHstart, LifespanHend - LifespanHstart);
                        int LifespanH = Int32.Parse(LifespanaH);
                        float LifespanFH = float.Parse(LifespanaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find Lifespan Color
                        int LifespanCstart1 = LifespanLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int LifespanCend1 = LifespanLine.IndexOf(Positionlast, LifespanCstart1);
                        string LifespanCH = LifespanLine.ToUpper().Substring(LifespanCstart1, LifespanCend1 - LifespanCstart1);
                        string LifespanCF = LifespanCH.Replace("0x", "#");
                        Color LifespanC = ColorTranslator.FromHtml(LifespanCF);
                        if (LifespanCF == "000000")
                        {
                            LifespanC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var Lifespanbrush = new SolidBrush(Color.FromArgb(LifespanC.A, LifespanC.R, LifespanC.G, LifespanC.B));
                        //find Lifespan Font
                        int LifespanFontstart = LifespanLine.IndexOf(CFont) + CFont.Length;
                        int LifespanFontend = LifespanLine.IndexOf(Positionlast, LifespanFontstart);
                        string LifespanaFont = LifespanLine.ToUpper().Substring(LifespanFontstart, LifespanFontend - LifespanFontstart);
                        Font LifespanFont;
                        if (LifespanaFont == "1")
                        {
                            LifespanFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (LifespanaFont == "2")
                        {
                            LifespanFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (LifespanaFont == "3")
                        {
                            LifespanFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            LifespanFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //EXPDate_txt
                        //find EXPDate_txt
                        string EXPDate_txtLine = spliter[31].ToString();
                        int EXPDate_txtstart1 = EXPDate_txtLine.IndexOf(Search1) + Search1.Length;
                        int EXPDate_txtend1 = EXPDate_txtLine.IndexOf(Search2, EXPDate_txtstart1);
                        string EXPDate_txt = EXPDate_txtLine.ToUpper().Substring(EXPDate_txtstart1, EXPDate_txtend1 - EXPDate_txtstart1);
                        string FinalEXPDate_txt = String.Join(" ", EXPDate_txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find EXPDate_txt locX
                        int EXPDate_txtXstart = EXPDate_txtLine.IndexOf(Positionx) + Positionx.Length;
                        int EXPDate_txtXend = EXPDate_txtLine.IndexOf(Positionlast, EXPDate_txtXstart);
                        string EXPDate_txtaX = EXPDate_txtLine.ToUpper().Substring(EXPDate_txtXstart, EXPDate_txtXend - EXPDate_txtXstart);
                        int EXPDate_txtX = Int32.Parse(EXPDate_txtaX);
                        //find EXPDate_txt locXY
                        int EXPDate_txtYstart = EXPDate_txtLine.IndexOf(Positiony) + Positiony.Length;
                        int EXPDate_txtYend = EXPDate_txtLine.IndexOf(Positionlast, EXPDate_txtYstart);
                        string EXPDate_txtaY = EXPDate_txtLine.ToUpper().Substring(EXPDate_txtYstart, EXPDate_txtYend - EXPDate_txtYstart);
                        int EXPDate_txtY = Int32.Parse(EXPDate_txtaY);
                        //Find EXPDate_txt width
                        int EXPDate_txtWstart = EXPDate_txtLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int EXPDate_txtWend = EXPDate_txtLine.IndexOf(Positionlast, EXPDate_txtWstart);
                        string EXPDate_txtaW = EXPDate_txtLine.ToUpper().Substring(EXPDate_txtWstart, EXPDate_txtWend - EXPDate_txtWstart);
                        int EXPDate_txtW = Int32.Parse(EXPDate_txtaW);
                        float EXPDate_txtFW = float.Parse(EXPDate_txtaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find EXPDate_txt height
                        int EXPDate_txtHstart = EXPDate_txtLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int EXPDate_txtHend = EXPDate_txtLine.IndexOf(Positionlast, EXPDate_txtHstart);
                        string EXPDate_txtaH = EXPDate_txtLine.ToUpper().Substring(EXPDate_txtHstart, EXPDate_txtHend - EXPDate_txtHstart);
                        int EXPDate_txtH = Int32.Parse(EXPDate_txtaH);
                        float EXPDate_txtFH = float.Parse(EXPDate_txtaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find EXPDate_txt Color
                        int EXPDate_txtCstart1 = EXPDate_txtLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int EXPDate_txtCend1 = EXPDate_txtLine.IndexOf(Positionlast, EXPDate_txtCstart1);
                        string EXPDate_txtCH = EXPDate_txtLine.ToUpper().Substring(EXPDate_txtCstart1, EXPDate_txtCend1 - EXPDate_txtCstart1);
                        string EXPDate_txtCF = EXPDate_txtCH.Replace("0x", "#");
                        Color EXPDate_txtC = ColorTranslator.FromHtml(EXPDate_txtCF);
                        if (EXPDate_txtCF == "000000")
                        {
                            EXPDate_txtC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var EXPDate_txtbrush = new SolidBrush(Color.FromArgb(EXPDate_txtC.A, EXPDate_txtC.R, EXPDate_txtC.G, EXPDate_txtC.B));
                        //find EXPDate_txt Font
                        int EXPDate_txtFontstart = EXPDate_txtLine.IndexOf(CFont) + CFont.Length;
                        int EXPDate_txtFontend = EXPDate_txtLine.IndexOf(Positionlast, EXPDate_txtFontstart);
                        string EXPDate_txtaFont = EXPDate_txtLine.ToUpper().Substring(EXPDate_txtFontstart, EXPDate_txtFontend - EXPDate_txtFontstart);
                        Font EXPDate_txtFont;
                        if (EXPDate_txtaFont == "1")
                        {
                            EXPDate_txtFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (EXPDate_txtaFont == "2")
                        {
                            EXPDate_txtFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (EXPDate_txtaFont == "3")
                        {
                            EXPDate_txtFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            EXPDate_txtFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //EXPDate
                        //find EXPDate
                        string EXPDateLine = spliter[32].ToString();
                        int EXPDatestart1 = EXPDateLine.IndexOf(Search1) + Search1.Length;
                        int EXPDateend1 = EXPDateLine.IndexOf(Search2, EXPDatestart1);
                        string EXPDate = EXPDateLine.ToUpper().Substring(EXPDatestart1, EXPDateend1 - EXPDatestart1);
                        string FinalEXPDate = String.Join(" ", EXPDate.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find EXPDate locX
                        int EXPDateXstart = EXPDateLine.IndexOf(Positionx) + Positionx.Length;
                        int EXPDateXend = EXPDateLine.IndexOf(Positionlast, EXPDateXstart);
                        string EXPDateaX = EXPDateLine.ToUpper().Substring(EXPDateXstart, EXPDateXend - EXPDateXstart);
                        int EXPDateX = Int32.Parse(EXPDateaX);
                        //find EXPDate locXY
                        int EXPDateYstart = EXPDateLine.IndexOf(Positiony) + Positiony.Length;
                        int EXPDateYend = EXPDateLine.IndexOf(Positionlast, EXPDateYstart);
                        string EXPDateaY = EXPDateLine.ToUpper().Substring(EXPDateYstart, EXPDateYend - EXPDateYstart);
                        int EXPDateY = Int32.Parse(EXPDateaY);
                        //Find EXPDate width
                        int EXPDateWstart = EXPDateLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int EXPDateWend = EXPDateLine.IndexOf(Positionlast, EXPDateWstart);
                        string EXPDateaW = EXPDateLine.ToUpper().Substring(EXPDateWstart, EXPDateWend - EXPDateWstart);
                        int EXPDateW = Int32.Parse(EXPDateaW);
                        float EXPDateFW = float.Parse(EXPDateaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find EXPDate height
                        int EXPDateHstart = EXPDateLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int EXPDateHend = EXPDateLine.IndexOf(Positionlast, EXPDateHstart);
                        string EXPDateaH = EXPDateLine.ToUpper().Substring(EXPDateHstart, EXPDateHend - EXPDateHstart);
                        int EXPDateH = Int32.Parse(EXPDateaH);
                        float EXPDateFH = float.Parse(EXPDateaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find EXPDate Color
                        int EXPDateCstart1 = EXPDateLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int EXPDateCend1 = EXPDateLine.IndexOf(Positionlast, EXPDateCstart1);
                        string EXPDateCH = EXPDateLine.ToUpper().Substring(EXPDateCstart1, EXPDateCend1 - EXPDateCstart1);
                        string EXPDateCF = EXPDateCH.Replace("0x", "#");
                        Color EXPDateC = ColorTranslator.FromHtml(EXPDateCF);
                        if (EXPDateCF == "000000")
                        {
                            EXPDateC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var EXPDatebrush = new SolidBrush(Color.FromArgb(EXPDateC.A, EXPDateC.R, EXPDateC.G, EXPDateC.B));
                        //find EXPDate Font
                        int EXPDateFontstart = EXPDateLine.IndexOf(CFont) + CFont.Length;
                        int EXPDateFontend = EXPDateLine.IndexOf(Positionlast, EXPDateFontstart);
                        string EXPDateaFont = EXPDateLine.ToUpper().Substring(EXPDateFontstart, EXPDateFontend - EXPDateFontstart);
                        Font EXPDateFont;
                        if (EXPDateaFont == "1")
                        {
                            EXPDateFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (EXPDateaFont == "2")
                        {
                            EXPDateFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (EXPDateaFont == "3")
                        {
                            EXPDateFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            EXPDateFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }


                        //StartWarranty_txt
                        //find StartWarranty_txt
                        string StartWarranty_txtLine = spliter[33].ToString();
                        int StartWarranty_txtstart1 = StartWarranty_txtLine.IndexOf(Search1) + Search1.Length;
                        int StartWarranty_txtend1 = StartWarranty_txtLine.IndexOf(Search2, StartWarranty_txtstart1);
                        string StartWarranty_txt = StartWarranty_txtLine.ToUpper().Substring(StartWarranty_txtstart1, StartWarranty_txtend1 - StartWarranty_txtstart1);
                        string FinalStartWarranty_txt = String.Join(" ", StartWarranty_txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find StartWarranty_txt locX
                        int StartWarranty_txtXstart = StartWarranty_txtLine.IndexOf(Positionx) + Positionx.Length;
                        int StartWarranty_txtXend = StartWarranty_txtLine.IndexOf(Positionlast, StartWarranty_txtXstart);
                        string StartWarranty_txtaX = StartWarranty_txtLine.ToUpper().Substring(StartWarranty_txtXstart, StartWarranty_txtXend - StartWarranty_txtXstart);
                        int StartWarranty_txtX = Int32.Parse(StartWarranty_txtaX);
                        //find StartWarranty_txt locXY
                        int StartWarranty_txtYstart = StartWarranty_txtLine.IndexOf(Positiony) + Positiony.Length;
                        int StartWarranty_txtYend = StartWarranty_txtLine.IndexOf(Positionlast, StartWarranty_txtYstart);
                        string StartWarranty_txtaY = StartWarranty_txtLine.ToUpper().Substring(StartWarranty_txtYstart, StartWarranty_txtYend - StartWarranty_txtYstart);
                        int StartWarranty_txtY = Int32.Parse(StartWarranty_txtaY);
                        //Find StartWarranty_txt width
                        int StartWarranty_txtWstart = StartWarranty_txtLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int StartWarranty_txtWend = StartWarranty_txtLine.IndexOf(Positionlast, StartWarranty_txtWstart);
                        string StartWarranty_txtaW = StartWarranty_txtLine.ToUpper().Substring(StartWarranty_txtWstart, StartWarranty_txtWend - StartWarranty_txtWstart);
                        int StartWarranty_txtW = Int32.Parse(StartWarranty_txtaW);
                        float StartWarranty_txtFW = float.Parse(StartWarranty_txtaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find StartWarranty_txt height
                        int StartWarranty_txtHstart = StartWarranty_txtLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int StartWarranty_txtHend = StartWarranty_txtLine.IndexOf(Positionlast, StartWarranty_txtHstart);
                        string StartWarranty_txtaH = StartWarranty_txtLine.ToUpper().Substring(StartWarranty_txtHstart, StartWarranty_txtHend - StartWarranty_txtHstart);
                        int StartWarranty_txtH = Int32.Parse(StartWarranty_txtaH);
                        float StartWarranty_txtFH = float.Parse(StartWarranty_txtaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find StartWarranty_txt Color
                        int StartWarranty_txtCstart1 = StartWarranty_txtLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int StartWarranty_txtCend1 = StartWarranty_txtLine.IndexOf(Positionlast, StartWarranty_txtCstart1);
                        string StartWarranty_txtCH = StartWarranty_txtLine.ToUpper().Substring(StartWarranty_txtCstart1, StartWarranty_txtCend1 - StartWarranty_txtCstart1);
                        string StartWarranty_txtCF = StartWarranty_txtCH.Replace("0x", "#");
                        Color StartWarranty_txtC = ColorTranslator.FromHtml(StartWarranty_txtCF);
                        if (StartWarranty_txtCF == "000000")
                        {
                            StartWarranty_txtC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var StartWarranty_txtbrush = new SolidBrush(Color.FromArgb(StartWarranty_txtC.A, StartWarranty_txtC.R, StartWarranty_txtC.G, StartWarranty_txtC.B));
                        //find StartWarranty_txt Font
                        int StartWarranty_txtFontstart = StartWarranty_txtLine.IndexOf(CFont) + CFont.Length;
                        int StartWarranty_txtFontend = StartWarranty_txtLine.IndexOf(Positionlast, StartWarranty_txtFontstart);
                        string StartWarranty_txtaFont = StartWarranty_txtLine.ToUpper().Substring(StartWarranty_txtFontstart, StartWarranty_txtFontend - StartWarranty_txtFontstart);
                        Font StartWarranty_txtFont;
                        if (StartWarranty_txtaFont == "1")
                        {
                            StartWarranty_txtFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (StartWarranty_txtaFont == "2")
                        {
                            StartWarranty_txtFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (StartWarranty_txtaFont == "3")
                        {
                            StartWarranty_txtFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            StartWarranty_txtFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //StartWarranty
                        //find StartWarranty
                        string StartWarrantyLine = spliter[34].ToString();
                        int StartWarrantystart1 = StartWarrantyLine.IndexOf(Search1) + Search1.Length;
                        int StartWarrantyend1 = StartWarrantyLine.IndexOf(Search2, StartWarrantystart1);
                        string StartWarranty = StartWarrantyLine.ToUpper().Substring(StartWarrantystart1, StartWarrantyend1 - StartWarrantystart1);
                        string FinalStartWarranty = String.Join(" ", StartWarranty.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find StartWarranty locX
                        int StartWarrantyXstart = StartWarrantyLine.IndexOf(Positionx) + Positionx.Length;
                        int StartWarrantyXend = StartWarrantyLine.IndexOf(Positionlast, StartWarrantyXstart);
                        string StartWarrantyaX = StartWarrantyLine.ToUpper().Substring(StartWarrantyXstart, StartWarrantyXend - StartWarrantyXstart);
                        int StartWarrantyX = Int32.Parse(StartWarrantyaX);
                        //find StartWarranty locXY
                        int StartWarrantyYstart = StartWarrantyLine.IndexOf(Positiony) + Positiony.Length;
                        int StartWarrantyYend = StartWarrantyLine.IndexOf(Positionlast, StartWarrantyYstart);
                        string StartWarrantyaY = StartWarrantyLine.ToUpper().Substring(StartWarrantyYstart, StartWarrantyYend - StartWarrantyYstart);
                        int StartWarrantyY = Int32.Parse(StartWarrantyaY);
                        //Find StartWarranty width
                        int StartWarrantyWstart = StartWarrantyLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int StartWarrantyWend = StartWarrantyLine.IndexOf(Positionlast, StartWarrantyWstart);
                        string StartWarrantyaW = StartWarrantyLine.ToUpper().Substring(StartWarrantyWstart, StartWarrantyWend - StartWarrantyWstart);
                        int StartWarrantyW = Int32.Parse(StartWarrantyaW);
                        float StartWarrantyFW = float.Parse(StartWarrantyaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find StartWarranty height
                        int StartWarrantyHstart = StartWarrantyLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int StartWarrantyHend = StartWarrantyLine.IndexOf(Positionlast, StartWarrantyHstart);
                        string StartWarrantyaH = StartWarrantyLine.ToUpper().Substring(StartWarrantyHstart, StartWarrantyHend - StartWarrantyHstart);
                        int StartWarrantyH = Int32.Parse(StartWarrantyaH);
                        float StartWarrantyFH = float.Parse(StartWarrantyaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find StartWarranty Color
                        int StartWarrantyCstart1 = StartWarrantyLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int StartWarrantyCend1 = StartWarrantyLine.IndexOf(Positionlast, StartWarrantyCstart1);
                        string StartWarrantyCH = StartWarrantyLine.ToUpper().Substring(StartWarrantyCstart1, StartWarrantyCend1 - StartWarrantyCstart1);
                        string StartWarrantyCF = StartWarrantyCH.Replace("0x", "#");
                        Color StartWarrantyC = ColorTranslator.FromHtml(StartWarrantyCF);
                        if (StartWarrantyCF == "000000")
                        {
                            StartWarrantyC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var StartWarrantybrush = new SolidBrush(Color.FromArgb(StartWarrantyC.A, StartWarrantyC.R, StartWarrantyC.G, StartWarrantyC.B));
                        //find StartWarranty Font
                        int StartWarrantyFontstart = StartWarrantyLine.IndexOf(CFont) + CFont.Length;
                        int StartWarrantyFontend = StartWarrantyLine.IndexOf(Positionlast, StartWarrantyFontstart);
                        string StartWarrantyaFont = StartWarrantyLine.ToUpper().Substring(StartWarrantyFontstart, StartWarrantyFontend - StartWarrantyFontstart);
                        Font StartWarrantyFont;
                        if (StartWarrantyaFont == "1")
                        {
                            StartWarrantyFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (StartWarrantyaFont == "2")
                        {
                            StartWarrantyFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (StartWarrantyaFont == "3")
                        {
                            StartWarrantyFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            StartWarrantyFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //EndWarranty_txt
                        //find EndWarranty_txt
                        string EndWarranty_txtLine = spliter[35].ToString();
                        int EndWarranty_txtstart1 = EndWarranty_txtLine.IndexOf(Search1) + Search1.Length;
                        int EndWarranty_txtend1 = EndWarranty_txtLine.IndexOf(Search2, EndWarranty_txtstart1);
                        string EndWarranty_txt = EndWarranty_txtLine.ToUpper().Substring(EndWarranty_txtstart1, EndWarranty_txtend1 - EndWarranty_txtstart1);
                        string FinalEndWarranty_txt = String.Join(" ", EndWarranty_txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find EndWarranty_txt locX
                        int EndWarranty_txtXstart = EndWarranty_txtLine.IndexOf(Positionx) + Positionx.Length;
                        int EndWarranty_txtXend = EndWarranty_txtLine.IndexOf(Positionlast, EndWarranty_txtXstart);
                        string EndWarranty_txtaX = EndWarranty_txtLine.ToUpper().Substring(EndWarranty_txtXstart, EndWarranty_txtXend - EndWarranty_txtXstart);
                        int EndWarranty_txtX = Int32.Parse(EndWarranty_txtaX);
                        //find EndWarranty_txt locXY
                        int EndWarranty_txtYstart = EndWarranty_txtLine.IndexOf(Positiony) + Positiony.Length;
                        int EndWarranty_txtYend = EndWarranty_txtLine.IndexOf(Positionlast, EndWarranty_txtYstart);
                        string EndWarranty_txtaY = EndWarranty_txtLine.ToUpper().Substring(EndWarranty_txtYstart, EndWarranty_txtYend - EndWarranty_txtYstart);
                        int EndWarranty_txtY = Int32.Parse(EndWarranty_txtaY);
                        //Find EndWarranty_txt width
                        int EndWarranty_txtWstart = EndWarranty_txtLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int EndWarranty_txtWend = EndWarranty_txtLine.IndexOf(Positionlast, EndWarranty_txtWstart);
                        string EndWarranty_txtaW = EndWarranty_txtLine.ToUpper().Substring(EndWarranty_txtWstart, EndWarranty_txtWend - EndWarranty_txtWstart);
                        int EndWarranty_txtW = Int32.Parse(EndWarranty_txtaW);
                        float EndWarranty_txtFW = float.Parse(EndWarranty_txtaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find EndWarranty_txt height
                        int EndWarranty_txtHstart = EndWarranty_txtLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int EndWarranty_txtHend = EndWarranty_txtLine.IndexOf(Positionlast, EndWarranty_txtHstart);
                        string EndWarranty_txtaH = EndWarranty_txtLine.ToUpper().Substring(EndWarranty_txtHstart, EndWarranty_txtHend - EndWarranty_txtHstart);
                        int EndWarranty_txtH = Int32.Parse(EndWarranty_txtaH);
                        float EndWarranty_txtFH = float.Parse(EndWarranty_txtaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find EndWarranty_txt Color
                        int EndWarranty_txtCstart1 = EndWarranty_txtLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int EndWarranty_txtCend1 = EndWarranty_txtLine.IndexOf(Positionlast, EndWarranty_txtCstart1);
                        string EndWarranty_txtCH = EndWarranty_txtLine.ToUpper().Substring(EndWarranty_txtCstart1, EndWarranty_txtCend1 - EndWarranty_txtCstart1);
                        string EndWarranty_txtCF = EndWarranty_txtCH.Replace("0x", "#");
                        Color EndWarranty_txtC = ColorTranslator.FromHtml(EndWarranty_txtCF);
                        if (EndWarranty_txtCF == "000000")
                        {
                            EndWarranty_txtC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var EndWarranty_txtbrush = new SolidBrush(Color.FromArgb(EndWarranty_txtC.A, EndWarranty_txtC.R, EndWarranty_txtC.G, EndWarranty_txtC.B));
                        //find EndWarranty_txt Font
                        int EndWarranty_txtFontstart = EndWarranty_txtLine.IndexOf(CFont) + CFont.Length;
                        int EndWarranty_txtFontend = EndWarranty_txtLine.IndexOf(Positionlast, EndWarranty_txtFontstart);
                        string EndWarranty_txtaFont = EndWarranty_txtLine.ToUpper().Substring(EndWarranty_txtFontstart, EndWarranty_txtFontend - EndWarranty_txtFontstart);
                        Font EndWarranty_txtFont;
                        if (EndWarranty_txtaFont == "1")
                        {
                            EndWarranty_txtFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (EndWarranty_txtaFont == "2")
                        {
                            EndWarranty_txtFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (EndWarranty_txtaFont == "3")
                        {
                            EndWarranty_txtFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            EndWarranty_txtFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //EndWarranty
                        //find EndWarranty
                        string EndWarrantyLine = spliter[36].ToString();
                        int EndWarrantystart1 = EndWarrantyLine.IndexOf(Search1) + Search1.Length;
                        int EndWarrantyend1 = EndWarrantyLine.IndexOf(Search2, EndWarrantystart1);
                        string EndWarranty = EndWarrantyLine.ToUpper().Substring(EndWarrantystart1, EndWarrantyend1 - EndWarrantystart1);
                        string FinalEndWarranty = String.Join(" ", EndWarranty.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find EndWarranty locX
                        int EndWarrantyXstart = EndWarrantyLine.IndexOf(Positionx) + Positionx.Length;
                        int EndWarrantyXend = EndWarrantyLine.IndexOf(Positionlast, EndWarrantyXstart);
                        string EndWarrantyaX = EndWarrantyLine.ToUpper().Substring(EndWarrantyXstart, EndWarrantyXend - EndWarrantyXstart);
                        int EndWarrantyX = Int32.Parse(EndWarrantyaX);
                        //find EndWarranty locXY
                        int EndWarrantyYstart = EndWarrantyLine.IndexOf(Positiony) + Positiony.Length;
                        int EndWarrantyYend = EndWarrantyLine.IndexOf(Positionlast, EndWarrantyYstart);
                        string EndWarrantyaY = EndWarrantyLine.ToUpper().Substring(EndWarrantyYstart, EndWarrantyYend - EndWarrantyYstart);
                        int EndWarrantyY = Int32.Parse(EndWarrantyaY);
                        //Find EndWarranty width
                        int EndWarrantyWstart = EndWarrantyLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int EndWarrantyWend = EndWarrantyLine.IndexOf(Positionlast, EndWarrantyWstart);
                        string EndWarrantyaW = EndWarrantyLine.ToUpper().Substring(EndWarrantyWstart, EndWarrantyWend - EndWarrantyWstart);
                        int EndWarrantyW = Int32.Parse(EndWarrantyaW);
                        float EndWarrantyFW = float.Parse(EndWarrantyaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find EndWarranty height
                        int EndWarrantyHstart = EndWarrantyLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int EndWarrantyHend = EndWarrantyLine.IndexOf(Positionlast, EndWarrantyHstart);
                        string EndWarrantyaH = EndWarrantyLine.ToUpper().Substring(EndWarrantyHstart, EndWarrantyHend - EndWarrantyHstart);
                        int EndWarrantyH = Int32.Parse(EndWarrantyaH);
                        float EndWarrantyFH = float.Parse(EndWarrantyaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find EndWarranty Color
                        int EndWarrantyCstart1 = EndWarrantyLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int EndWarrantyCend1 = EndWarrantyLine.IndexOf(Positionlast, EndWarrantyCstart1);
                        string EndWarrantyCH = EndWarrantyLine.ToUpper().Substring(EndWarrantyCstart1, EndWarrantyCend1 - EndWarrantyCstart1);
                        string EndWarrantyCF = EndWarrantyCH.Replace("0x", "#");
                        Color EndWarrantyC = ColorTranslator.FromHtml(EndWarrantyCF);
                        if (EndWarrantyCF == "000000")
                        {
                            EndWarrantyC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var EndWarrantybrush = new SolidBrush(Color.FromArgb(EndWarrantyC.A, EndWarrantyC.R, EndWarrantyC.G, EndWarrantyC.B));
                        //find EndWarranty Font
                        int EndWarrantyFontstart = EndWarrantyLine.IndexOf(CFont) + CFont.Length;
                        int EndWarrantyFontend = EndWarrantyLine.IndexOf(Positionlast, EndWarrantyFontstart);
                        string EndWarrantyaFont = EndWarrantyLine.ToUpper().Substring(EndWarrantyFontstart, EndWarrantyFontend - EndWarrantyFontstart);
                        Font EndWarrantyFont;
                        if (EndWarrantyaFont == "1")
                        {
                            EndWarrantyFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (EndWarrantyaFont == "2")
                        {
                            EndWarrantyFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (EndWarrantyaFont == "3")
                        {
                            EndWarrantyFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            EndWarrantyFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //Status_txt
                        //find Status_txt
                        string Status_txtLine = spliter[37].ToString();
                        int Status_txtstart1 = Status_txtLine.IndexOf(Search1) + Search1.Length;
                        int Status_txtend1 = Status_txtLine.IndexOf(Search2, Status_txtstart1);
                        string Status_txt = Status_txtLine.ToUpper().Substring(Status_txtstart1, Status_txtend1 - Status_txtstart1);
                        string FinalStatus_txt = String.Join(" ", Status_txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find Status_txt locX
                        int Status_txtXstart = Status_txtLine.IndexOf(Positionx) + Positionx.Length;
                        int Status_txtXend = Status_txtLine.IndexOf(Positionlast, Status_txtXstart);
                        string Status_txtaX = Status_txtLine.ToUpper().Substring(Status_txtXstart, Status_txtXend - Status_txtXstart);
                        int Status_txtX = Int32.Parse(Status_txtaX);
                        //find Status_txt locXY
                        int Status_txtYstart = Status_txtLine.IndexOf(Positiony) + Positiony.Length;
                        int Status_txtYend = Status_txtLine.IndexOf(Positionlast, Status_txtYstart);
                        string Status_txtaY = Status_txtLine.ToUpper().Substring(Status_txtYstart, Status_txtYend - Status_txtYstart);
                        int Status_txtY = Int32.Parse(Status_txtaY);
                        //Find Status_txt width
                        int Status_txtWstart = Status_txtLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int Status_txtWend = Status_txtLine.IndexOf(Positionlast, Status_txtWstart);
                        string Status_txtaW = Status_txtLine.ToUpper().Substring(Status_txtWstart, Status_txtWend - Status_txtWstart);
                        int Status_txtW = Int32.Parse(Status_txtaW);
                        float Status_txtFW = float.Parse(Status_txtaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find Status_txt height
                        int Status_txtHstart = Status_txtLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int Status_txtHend = Status_txtLine.IndexOf(Positionlast, Status_txtHstart);
                        string Status_txtaH = Status_txtLine.ToUpper().Substring(Status_txtHstart, Status_txtHend - Status_txtHstart);
                        int Status_txtH = Int32.Parse(Status_txtaH);
                        float Status_txtFH = float.Parse(Status_txtaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find Status_txt Color
                        int Status_txtCstart1 = Status_txtLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int Status_txtCend1 = Status_txtLine.IndexOf(Positionlast, Status_txtCstart1);
                        string Status_txtCH = Status_txtLine.ToUpper().Substring(Status_txtCstart1, Status_txtCend1 - Status_txtCstart1);
                        string Status_txtCF = Status_txtCH.Replace("0x", "#");
                        Color Status_txtC = ColorTranslator.FromHtml(Status_txtCF);
                        if (Status_txtCF == "000000")
                        {
                            Status_txtC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var Status_txtbrush = new SolidBrush(Color.FromArgb(Status_txtC.A, Status_txtC.R, Status_txtC.G, Status_txtC.B));
                        //find Status_txt Font
                        int Status_txtFontstart = Status_txtLine.IndexOf(CFont) + CFont.Length;
                        int Status_txtFontend = Status_txtLine.IndexOf(Positionlast, Status_txtFontstart);
                        string Status_txtaFont = Status_txtLine.ToUpper().Substring(Status_txtFontstart, Status_txtFontend - Status_txtFontstart);
                        Font Status_txtFont;
                        if (Status_txtaFont == "1")
                        {
                            Status_txtFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (Status_txtaFont == "2")
                        {
                            Status_txtFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (Status_txtaFont == "3")
                        {
                            Status_txtFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            Status_txtFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }


                        //Status
                        //find Status
                        string StatusLine = spliter[38].ToString();
                        int Statusstart1 = StatusLine.IndexOf(Search1) + Search1.Length;
                        int Statusend1 = StatusLine.IndexOf(Search2, Statusstart1);
                        string Status = StatusLine.ToUpper().Substring(Statusstart1, Statusend1 - Statusstart1);
                        string FinalStatus = String.Join(" ", Status.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find Status locX
                        int StatusXstart = StatusLine.IndexOf(Positionx) + Positionx.Length;
                        int StatusXend = StatusLine.IndexOf(Positionlast, StatusXstart);
                        string StatusaX = StatusLine.ToUpper().Substring(StatusXstart, StatusXend - StatusXstart);
                        int StatusX = Int32.Parse(StatusaX);
                        //find Status locXY
                        int StatusYstart = StatusLine.IndexOf(Positiony) + Positiony.Length;
                        int StatusYend = StatusLine.IndexOf(Positionlast, StatusYstart);
                        string StatusaY = StatusLine.ToUpper().Substring(StatusYstart, StatusYend - StatusYstart);
                        int StatusY = Int32.Parse(StatusaY);
                        //Find Status width
                        int StatusWstart = StatusLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int StatusWend = StatusLine.IndexOf(Positionlast, StatusWstart);
                        string StatusaW = StatusLine.ToUpper().Substring(StatusWstart, StatusWend - StatusWstart);
                        int StatusW = Int32.Parse(StatusaW);
                        float StatusFW = float.Parse(StatusaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find Status height
                        int StatusHstart = StatusLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int StatusHend = StatusLine.IndexOf(Positionlast, StatusHstart);
                        string StatusaH = StatusLine.ToUpper().Substring(StatusHstart, StatusHend - StatusHstart);
                        int StatusH = Int32.Parse(StatusaH);
                        float StatusFH = float.Parse(StatusaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find Status Color
                        int StatusCstart1 = StatusLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int StatusCend1 = StatusLine.IndexOf(Positionlast, StatusCstart1);
                        string StatusCH = StatusLine.ToUpper().Substring(StatusCstart1, StatusCend1 - StatusCstart1);
                        string StatusCF = StatusCH.Replace("0x", "#");
                        Color StatusC = ColorTranslator.FromHtml(StatusCF);
                        if (StatusCF == "000000")
                        {
                            StatusC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var Statusbrush = new SolidBrush(Color.FromArgb(StatusC.A, StatusC.R, StatusC.G, StatusC.B));
                        //find Status Font
                        int StatusFontstart = StatusLine.IndexOf(CFont) + CFont.Length;
                        int StatusFontend = StatusLine.IndexOf(Positionlast, StatusFontstart);
                        string StatusaFont = StatusLine.ToUpper().Substring(StatusFontstart, StatusFontend - StatusFontstart);
                        Font StatusFont;
                        if (StatusaFont == "1")
                        {
                            StatusFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (StatusaFont == "2")
                        {
                            StatusFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (StatusaFont == "3")
                        {
                            StatusFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            StatusFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //WarrantySpan
                        //find WarrantySpan
                        string WarrantySpanLine = spliter[39].ToString();
                        int WarrantySpanstart1 = WarrantySpanLine.IndexOf(Search1) + Search1.Length;
                        int WarrantySpanend1 = WarrantySpanLine.IndexOf(Search2, WarrantySpanstart1);
                        string WarrantySpan = WarrantySpanLine.ToUpper().Substring(WarrantySpanstart1, WarrantySpanend1 - WarrantySpanstart1);
                        string FinalWarrantySpan = String.Join(" ", WarrantySpan.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find WarrantySpan locX
                        int WarrantySpanXstart = WarrantySpanLine.IndexOf(Positionx) + Positionx.Length;
                        int WarrantySpanXend = WarrantySpanLine.IndexOf(Positionlast, WarrantySpanXstart);
                        string WarrantySpanaX = WarrantySpanLine.ToUpper().Substring(WarrantySpanXstart, WarrantySpanXend - WarrantySpanXstart);
                        int WarrantySpanX = Int32.Parse(WarrantySpanaX);
                        //find WarrantySpan locXY
                        int WarrantySpanYstart = WarrantySpanLine.IndexOf(Positiony) + Positiony.Length;
                        int WarrantySpanYend = WarrantySpanLine.IndexOf(Positionlast, WarrantySpanYstart);
                        string WarrantySpanaY = WarrantySpanLine.ToUpper().Substring(WarrantySpanYstart, WarrantySpanYend - WarrantySpanYstart);
                        int WarrantySpanY = Int32.Parse(WarrantySpanaY);
                        //Find WarrantySpan width
                        int WarrantySpanWstart = WarrantySpanLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int WarrantySpanWend = WarrantySpanLine.IndexOf(Positionlast, WarrantySpanWstart);
                        string WarrantySpanaW = WarrantySpanLine.ToUpper().Substring(WarrantySpanWstart, WarrantySpanWend - WarrantySpanWstart);
                        int WarrantySpanW = Int32.Parse(WarrantySpanaW);
                        float WarrantySpanFW = float.Parse(WarrantySpanaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find WarrantySpan height
                        int WarrantySpanHstart = WarrantySpanLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int WarrantySpanHend = WarrantySpanLine.IndexOf(Positionlast, WarrantySpanHstart);
                        string WarrantySpanaH = WarrantySpanLine.ToUpper().Substring(WarrantySpanHstart, WarrantySpanHend - WarrantySpanHstart);
                        int WarrantySpanH = Int32.Parse(WarrantySpanaH);
                        float WarrantySpanFH = float.Parse(WarrantySpanaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find WarrantySpan Color
                        int WarrantySpanCstart1 = WarrantySpanLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int WarrantySpanCend1 = WarrantySpanLine.IndexOf(Positionlast, WarrantySpanCstart1);
                        string WarrantySpanCH = WarrantySpanLine.ToUpper().Substring(WarrantySpanCstart1, WarrantySpanCend1 - WarrantySpanCstart1);
                        string WarrantySpanCF = WarrantySpanCH.Replace("0x", "#");
                        Color WarrantySpanC = ColorTranslator.FromHtml(WarrantySpanCF);
                        if (WarrantySpanCF == "000000")
                        {
                            WarrantySpanC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var WarrantySpanbrush = new SolidBrush(Color.FromArgb(WarrantySpanC.A, WarrantySpanC.R, WarrantySpanC.G, WarrantySpanC.B));
                        //find WarrantySpan Font
                        int WarrantySpanFontstart = WarrantySpanLine.IndexOf(CFont) + CFont.Length;
                        int WarrantySpanFontend = WarrantySpanLine.IndexOf(Positionlast, WarrantySpanFontstart);
                        string WarrantySpanaFont = WarrantySpanLine.ToUpper().Substring(WarrantySpanFontstart, WarrantySpanFontend - WarrantySpanFontstart);
                        Font WarrantySpanFont;
                        if (WarrantySpanaFont == "1")
                        {
                            WarrantySpanFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (WarrantySpanaFont == "2")
                        {
                            WarrantySpanFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (WarrantySpanaFont == "3")
                        {
                            WarrantySpanFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            WarrantySpanFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //FinalExpiration_txt
                        //find FinalExpiration_txt
                        string FinalExpiration_txtLine = spliter[40].ToString();
                        int FinalExpiration_txtstart1 = FinalExpiration_txtLine.IndexOf(Search1) + Search1.Length;
                        int FinalExpiration_txtend1 = FinalExpiration_txtLine.IndexOf(Search2, FinalExpiration_txtstart1);
                        string FinalExpiration_txt = FinalExpiration_txtLine.ToUpper().Substring(FinalExpiration_txtstart1, FinalExpiration_txtend1 - FinalExpiration_txtstart1);
                        string FinalFinalExpiration_txt = String.Join(" ", FinalExpiration_txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find FinalExpiration_txt locX
                        int FinalExpiration_txtXstart = FinalExpiration_txtLine.IndexOf(Positionx) + Positionx.Length;
                        int FinalExpiration_txtXend = FinalExpiration_txtLine.IndexOf(Positionlast, FinalExpiration_txtXstart);
                        string FinalExpiration_txtaX = FinalExpiration_txtLine.ToUpper().Substring(FinalExpiration_txtXstart, FinalExpiration_txtXend - FinalExpiration_txtXstart);
                        int FinalExpiration_txtX = Int32.Parse(FinalExpiration_txtaX);
                        //find FinalExpiration_txt locXY
                        int FinalExpiration_txtYstart = FinalExpiration_txtLine.IndexOf(Positiony) + Positiony.Length;
                        int FinalExpiration_txtYend = FinalExpiration_txtLine.IndexOf(Positionlast, FinalExpiration_txtYstart);
                        string FinalExpiration_txtaY = FinalExpiration_txtLine.ToUpper().Substring(FinalExpiration_txtYstart, FinalExpiration_txtYend - FinalExpiration_txtYstart);
                        int FinalExpiration_txtY = Int32.Parse(FinalExpiration_txtaY);
                        //Find FinalExpiration_txt width
                        int FinalExpiration_txtWstart = FinalExpiration_txtLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int FinalExpiration_txtWend = FinalExpiration_txtLine.IndexOf(Positionlast, FinalExpiration_txtWstart);
                        string FinalExpiration_txtaW = FinalExpiration_txtLine.ToUpper().Substring(FinalExpiration_txtWstart, FinalExpiration_txtWend - FinalExpiration_txtWstart);
                        int FinalExpiration_txtW = Int32.Parse(FinalExpiration_txtaW);
                        float FinalExpiration_txtFW = float.Parse(FinalExpiration_txtaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find FinalExpiration_txt height
                        int FinalExpiration_txtHstart = FinalExpiration_txtLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int FinalExpiration_txtHend = FinalExpiration_txtLine.IndexOf(Positionlast, FinalExpiration_txtHstart);
                        string FinalExpiration_txtaH = FinalExpiration_txtLine.ToUpper().Substring(FinalExpiration_txtHstart, FinalExpiration_txtHend - FinalExpiration_txtHstart);
                        int FinalExpiration_txtH = Int32.Parse(FinalExpiration_txtaH);
                        float FinalExpiration_txtFH = float.Parse(FinalExpiration_txtaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find FinalExpiration_txt Color
                        int FinalExpiration_txtCstart1 = FinalExpiration_txtLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int FinalExpiration_txtCend1 = FinalExpiration_txtLine.IndexOf(Positionlast, FinalExpiration_txtCstart1);
                        string FinalExpiration_txtCH = FinalExpiration_txtLine.ToUpper().Substring(FinalExpiration_txtCstart1, FinalExpiration_txtCend1 - FinalExpiration_txtCstart1);
                        string FinalExpiration_txtCF = FinalExpiration_txtCH.Replace("0x", "#");
                        Color FinalExpiration_txtC = ColorTranslator.FromHtml(FinalExpiration_txtCF);
                        if (FinalExpiration_txtCF == "000000")
                        {
                            FinalExpiration_txtC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var FinalExpiration_txtbrush = new SolidBrush(Color.FromArgb(FinalExpiration_txtC.A, FinalExpiration_txtC.R, FinalExpiration_txtC.G, FinalExpiration_txtC.B));
                        //find FinalExpiration_txt Font
                        int FinalExpiration_txtFontstart = FinalExpiration_txtLine.IndexOf(CFont) + CFont.Length;
                        int FinalExpiration_txtFontend = FinalExpiration_txtLine.IndexOf(Positionlast, FinalExpiration_txtFontstart);
                        string FinalExpiration_txtaFont = FinalExpiration_txtLine.ToUpper().Substring(FinalExpiration_txtFontstart, FinalExpiration_txtFontend - FinalExpiration_txtFontstart);
                        Font FinalExpiration_txtFont;
                        if (FinalExpiration_txtaFont == "1")
                        {
                            FinalExpiration_txtFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (FinalExpiration_txtaFont == "2")
                        {
                            FinalExpiration_txtFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (FinalExpiration_txtaFont == "3")
                        {
                            FinalExpiration_txtFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            FinalExpiration_txtFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //FinalExpiration
                        //find FinalExpiration
                        string FinalExpirationLine = spliter[41].ToString();
                        int FinalExpirationstart1 = FinalExpirationLine.IndexOf(Search1) + Search1.Length;
                        int FinalExpirationend1 = FinalExpirationLine.IndexOf(Search2, FinalExpirationstart1);
                        string FinalExpiration = FinalExpirationLine.ToUpper().Substring(FinalExpirationstart1, FinalExpirationend1 - FinalExpirationstart1);
                        string FinalFinalExpiration = String.Join(" ", FinalExpiration.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find FinalExpiration locX
                        int FinalExpirationXstart = FinalExpirationLine.IndexOf(Positionx) + Positionx.Length;
                        int FinalExpirationXend = FinalExpirationLine.IndexOf(Positionlast, FinalExpirationXstart);
                        string FinalExpirationaX = FinalExpirationLine.ToUpper().Substring(FinalExpirationXstart, FinalExpirationXend - FinalExpirationXstart);
                        int FinalExpirationX = Int32.Parse(FinalExpirationaX);
                        //find FinalExpiration locXY
                        int FinalExpirationYstart = FinalExpirationLine.IndexOf(Positiony) + Positiony.Length;
                        int FinalExpirationYend = FinalExpirationLine.IndexOf(Positionlast, FinalExpirationYstart);
                        string FinalExpirationaY = FinalExpirationLine.ToUpper().Substring(FinalExpirationYstart, FinalExpirationYend - FinalExpirationYstart);
                        int FinalExpirationY = Int32.Parse(FinalExpirationaY);
                        //Find FinalExpiration width
                        int FinalExpirationWstart = FinalExpirationLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int FinalExpirationWend = FinalExpirationLine.IndexOf(Positionlast, FinalExpirationWstart);
                        string FinalExpirationaW = FinalExpirationLine.ToUpper().Substring(FinalExpirationWstart, FinalExpirationWend - FinalExpirationWstart);
                        int FinalExpirationW = Int32.Parse(FinalExpirationaW);
                        float FinalExpirationFW = float.Parse(FinalExpirationaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find FinalExpiration height
                        int FinalExpirationHstart = FinalExpirationLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int FinalExpirationHend = FinalExpirationLine.IndexOf(Positionlast, FinalExpirationHstart);
                        string FinalExpirationaH = FinalExpirationLine.ToUpper().Substring(FinalExpirationHstart, FinalExpirationHend - FinalExpirationHstart);
                        int FinalExpirationH = Int32.Parse(FinalExpirationaH);
                        float FinalExpirationFH = float.Parse(FinalExpirationaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find FinalExpiration Color
                        int FinalExpirationCstart1 = FinalExpirationLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int FinalExpirationCend1 = FinalExpirationLine.IndexOf(Positionlast, FinalExpirationCstart1);
                        string FinalExpirationCH = FinalExpirationLine.ToUpper().Substring(FinalExpirationCstart1, FinalExpirationCend1 - FinalExpirationCstart1);
                        string FinalExpirationCF = FinalExpirationCH.Replace("0x", "#");
                        Color FinalExpirationC = ColorTranslator.FromHtml(FinalExpirationCF);
                        if (FinalExpirationCF == "000000")
                        {
                            FinalExpirationC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var FinalExpirationbrush = new SolidBrush(Color.FromArgb(FinalExpirationC.A, FinalExpirationC.R, FinalExpirationC.G, FinalExpirationC.B));
                        //find FinalExpiration Font
                        int FinalExpirationFontstart = FinalExpirationLine.IndexOf(CFont) + CFont.Length;
                        int FinalExpirationFontend = FinalExpirationLine.IndexOf(Positionlast, FinalExpirationFontstart);
                        string FinalExpirationaFont = FinalExpirationLine.ToUpper().Substring(FinalExpirationFontstart, FinalExpirationFontend - FinalExpirationFontstart);
                        Font FinalExpirationFont;
                        if (FinalExpirationaFont == "1")
                        {
                            FinalExpirationFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (FinalExpirationaFont == "2")
                        {
                            FinalExpirationFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (FinalExpirationaFont == "3")
                        {
                            FinalExpirationFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            FinalExpirationFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //Details_txt
                        //find Details_txt
                        string Details_txtLine = spliter[42].ToString();
                        int Details_txtstart1 = Details_txtLine.IndexOf(Search1) + Search1.Length;
                        int Details_txtend1 = Details_txtLine.IndexOf(Search2, Details_txtstart1);
                        string Details_txt = Details_txtLine.ToUpper().Substring(Details_txtstart1, Details_txtend1 - Details_txtstart1);
                        string FinalDetails_txt = String.Join(" ", Details_txt.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find Details_txt locX
                        int Details_txtXstart = Details_txtLine.IndexOf(Positionx) + Positionx.Length;
                        int Details_txtXend = Details_txtLine.IndexOf(Positionlast, Details_txtXstart);
                        string Details_txtaX = Details_txtLine.ToUpper().Substring(Details_txtXstart, Details_txtXend - Details_txtXstart);
                        int Details_txtX = Int32.Parse(Details_txtaX);
                        //find Details_txt locXY
                        int Details_txtYstart = Details_txtLine.IndexOf(Positiony) + Positiony.Length;
                        int Details_txtYend = Details_txtLine.IndexOf(Positionlast, Details_txtYstart);
                        string Details_txtaY = Details_txtLine.ToUpper().Substring(Details_txtYstart, Details_txtYend - Details_txtYstart);
                        int Details_txtY = Int32.Parse(Details_txtaY);
                        //find Details_txt Color
                        int Details_txtCstart1 = Details_txtLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int Details_txtCend1 = Details_txtLine.IndexOf(Positionlast, Details_txtCstart1);
                        string Details_txtCH = Details_txtLine.ToUpper().Substring(Details_txtCstart1, Details_txtCend1 - Details_txtCstart1);
                        string Details_txtCF = Details_txtCH.Replace("0x", "#");
                        Color Details_txtC = ColorTranslator.FromHtml(Details_txtCF);
                        if (Details_txtCF == "000000")
                        {
                            Details_txtC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var Details_txtbrush = new SolidBrush(Color.FromArgb(Details_txtC.A, Details_txtC.R, Details_txtC.G, Details_txtC.B));
                        //find Details_txt Font
                        int Details_txtFontstart = Details_txtLine.IndexOf(CFont) + CFont.Length;
                        int Details_txtFontend = Details_txtLine.IndexOf(Positionlast, Details_txtFontstart);
                        string Details_txtaFont = Details_txtLine.ToUpper().Substring(Details_txtFontstart, Details_txtFontend - Details_txtFontstart);
                        Font Details_txtFont;
                        if (Details_txtaFont == "1")
                        {
                            Details_txtFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (Details_txtaFont == "2")
                        {
                            Details_txtFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (Details_txtaFont == "3")
                        {
                            Details_txtFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            Details_txtFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //Details
                        //find Details
                        string DetailsLine = spliter[43].ToString();
                        int Detailsstart1 = DetailsLine.IndexOf(Search1) + Search1.Length;
                        int Detailsend1 = DetailsLine.IndexOf(Search2, Detailsstart1);
                        string Details = DetailsLine.ToUpper().Substring(Detailsstart1, Detailsend1 - Detailsstart1);
                        string FinalDetails = String.Join(" ", Details.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries));
                        //find Details locX
                        int DetailsXstart = DetailsLine.IndexOf(Positionx) + Positionx.Length;
                        int DetailsXend = DetailsLine.IndexOf(Positionlast, DetailsXstart);
                        string DetailsaX = DetailsLine.ToUpper().Substring(DetailsXstart, DetailsXend - DetailsXstart);
                        int DetailsX = Int32.Parse(DetailsaX);
                        //find Details locXY
                        int DetailsYstart = DetailsLine.IndexOf(Positiony) + Positiony.Length;
                        int DetailsYend = DetailsLine.IndexOf(Positionlast, DetailsYstart);
                        string DetailsaY = DetailsLine.ToUpper().Substring(DetailsYstart, DetailsYend - DetailsYstart);
                        int DetailsY = Int32.Parse(DetailsaY);
                        //Find Details width
                        int DetailsWstart = DetailsLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int DetailsWend = DetailsLine.IndexOf(Positionlast, DetailsWstart);
                        string DetailsaW = DetailsLine.ToUpper().Substring(DetailsWstart, DetailsWend - DetailsWstart);
                        int DetailsW = Int32.Parse(DetailsaW);
                        float DetailsFW = float.Parse(DetailsaW, CultureInfo.InvariantCulture.NumberFormat);
                        //find Details height
                        int DetailsHstart = DetailsLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int DetailsHend = DetailsLine.IndexOf(Positionlast, DetailsHstart);
                        string DetailsaH = DetailsLine.ToUpper().Substring(DetailsHstart, DetailsHend - DetailsHstart);
                        int DetailsH = Int32.Parse(DetailsaH);
                        float DetailsFH = float.Parse(DetailsaH, CultureInfo.InvariantCulture.NumberFormat);
                        //find Details Color
                        int DetailsCstart1 = DetailsLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int DetailsCend1 = DetailsLine.IndexOf(Positionlast, DetailsCstart1);
                        string DetailsCH = DetailsLine.ToUpper().Substring(DetailsCstart1, DetailsCend1 - DetailsCstart1);
                        string DetailsCF = DetailsCH.Replace("0x", "#");
                        Color DetailsC = ColorTranslator.FromHtml(DetailsCF);
                        if (DetailsCF == "000000")
                        {
                            DetailsC = Color.FromArgb(255, 255, 255, 255);
                        }
                        var Detailsbrush = new SolidBrush(Color.FromArgb(DetailsC.A, DetailsC.R, DetailsC.G, DetailsC.B));
                        //find Details Font
                        int DetailsFontstart = DetailsLine.IndexOf(CFont) + CFont.Length;
                        int DetailsFontend = DetailsLine.IndexOf(Positionlast, DetailsFontstart);
                        string DetailsaFont = DetailsLine.ToUpper().Substring(DetailsFontstart, DetailsFontend - DetailsFontstart);
                        Font DetailsFont;
                        if (DetailsaFont == "1")
                        {
                            DetailsFont = new Font(IDF1, IDS1, FontStyle.Bold);
                        }
                        else if (DetailsaFont == "2")
                        {
                            DetailsFont = new Font(IDF2, IDS2, FontStyle.Bold);
                        }
                        else if (DetailsaFont == "3")
                        {
                            DetailsFont = new Font(IDF3, IDS3, FontStyle.Regular);
                        }
                        else
                        {
                            DetailsFont = new Font(IDF4, IDS4, FontStyle.Regular);
                        }

                        //for Horizontal Line
                        string LineLine = spliter[45].ToString();
                        //Find Line x1
                        int LineX1start = LineLine.IndexOf(x1) + x1.Length;
                        int LineX1end = LineLine.IndexOf(Positionlast, LineX1start);
                        string LineaX1 = LineLine.ToUpper().Substring(LineX1start, LineX1end - LineX1start);
                        int LineX1 = Int32.Parse(LineaX1);
                        //find Line x2
                        int LineX2start = LineLine.IndexOf(x2) + x2.Length;
                        int LineX2end = LineLine.IndexOf(Positionlast, LineX2start);
                        string LineaX2 = LineLine.ToUpper().Substring(LineX2start, LineX2end - LineX2start);
                        int LineX2 = Int32.Parse(LineaX2);
                        //find Line y1
                        int LineY1start = LineLine.IndexOf(y1) + y1.Length;
                        int LineY1end = LineLine.IndexOf(Positionlast, LineY1start);
                        string LineaY1 = LineLine.ToUpper().Substring(LineY1start, LineY1end - LineY1start);
                        int LineY1 = Int32.Parse(LineaY1);
                        //find Line y2
                        int LineY2start = LineLine.IndexOf(y2) + y2.Length;
                        int LineY2end = LineLine.IndexOf(Positionlast, LineY2start);
                        string LineaY2 = LineLine.ToUpper().Substring(LineY2start, LineY2end - LineY2start);
                        int LineY2 = Int32.Parse(LineaY2);

                        string QRLine = spliter[44].ToString();
                        int QRstart1 = QRLine.IndexOf(Search1) + Search1.Length;
                        int QRend1 = QRLine.IndexOf(Search2, QRstart1);
                        string QR = QRLine.ToUpper().Substring(QRstart1, QRend1 - QRstart1);
                        //find QR locX
                        int QRXstart = QRLine.IndexOf(Positionx) + Positionx.Length;
                        int QRXend = QRLine.IndexOf(Positionlast, QRXstart);
                        string QRaX = QRLine.ToUpper().Substring(QRXstart, QRXend - QRXstart);
                        int QRX = Int32.Parse(QRaX);
                        //find QR locXY
                        int QRYstart = QRLine.IndexOf(Positiony) + Positiony.Length;
                        int QRYend = QRLine.IndexOf(Positionlast, QRYstart);
                        string QRaY = QRLine.ToUpper().Substring(QRYstart, QRYend - QRYstart);
                        int QRY = Int32.Parse(QRaY);
                        //find QR Color
                        int QRAddCstart1 = QRLine.IndexOf(Colorsearch) + Colorsearch.Length;
                        int QRAddCend1 = QRLine.IndexOf(Positionlast, QRAddCstart1);
                        string QRAddC = QRLine.ToUpper().Substring(QRAddCstart1, QRAddCend1 - QRAddCstart1);
                        //Find QR width
                        int QRWstart = QRLine.IndexOf(Widthsearch) + Widthsearch.Length;
                        int QRWend = QRLine.IndexOf(Positionlast, QRWstart);
                        string QRaW = QRLine.ToUpper().Substring(QRWstart, QRWend - QRWstart);
                        int QRFW = Int32.Parse(QRaW);
                        //find QR height
                        int QRHstart = QRLine.IndexOf(Heightsearch) + Heightsearch.Length;
                        int QRHend = QRLine.IndexOf(Positionlast, QRHstart);
                        string QRaH = QRLine.ToUpper().Substring(QRHstart, QRHend - QRHstart);
                        int QRFH = Int32.Parse(QRaH);
                        //find QR multiplier
                        int QRMstart = QRLine.IndexOf(Multipliersearch) + Multipliersearch.Length;
                        int QRMend = QRLine.IndexOf(Positionlast, QRMstart);
                        string QRaM = QRLine.ToUpper().Substring(QRMstart, QRMend - QRMstart);
                        int QRM = Int32.Parse(QRaM);

                        //universal
                        List<System.Drawing.Bitmap> images = new List<System.Drawing.Bitmap>();
                        System.Drawing.Bitmap Warrantyimg = null;

                        int BackgroundSizeW = 1063;
                        int BackgroundSizeH = 827;
                        int PapersizeW = 900;
                        int PapersizeH = 700;
                        var Colortransparent = new SolidBrush(Color.Transparent);
                        var Colorwhite = new SolidBrush(Color.White);

                        //for QRcode
                        SfBarcode barcode = new SfBarcode();
                        barcode.Text = QR;
                        barcode.Symbology = BarcodeSymbolType.QRBarcode;

                        //Set dimension for each block
                        QRBarcodeSetting setting = new QRBarcodeSetting();
                        setting.XDimension = 3;
                        barcode.SymbologySettings = setting;
                        //Export to image
                        Image QRimg = barcode.ToImage(new SizeF(QRFH, QRFW));

                        //for Rectangle/Container Header
                        System.Drawing.Pen blackPen = new System.Drawing.Pen(Color.Black, 4);

                        //for Center Text
                        StringFormat stringFormat = new StringFormat();
                        stringFormat.Alignment = StringAlignment.Center;
                        stringFormat.LineAlignment = StringAlignment.Center;

                        //for letft Text
                        StringFormat stringFormat1 = new StringFormat();
                        stringFormat1.Alignment = StringAlignment.Near;
                        stringFormat1.LineAlignment = StringAlignment.Near;

                        Rectangle HeaderRec = new Rectangle(HeaderX, HeaderY-10, HeaderW, HeaderH+10);

                        Rectangle WarehouseLocRec = new Rectangle(WarehouseLocation_txtX, WarehouseLocation_txtY, WarehouseLocation_txtW, WarehouseLocation_txtH);
                        Rectangle FixedAssetRec = new Rectangle(FixedAssetLocation_txtX, FixedAssetLocation_txtY, FixedAssetLocation_txtW, FixedAssetLocation_txtH);
                        Rectangle WarehouseLocValRec = new Rectangle(WarehouseLocationX, WarehouseLocationY, WarehouseLocationW, WarehouseLocationH);
                        Rectangle FixedAssetValRec = new Rectangle(FixedAssetLocationX, FixedAssetLocationY, FixedAssetLocationW, FixedAssetLocationH);

                        Rectangle MFGDateRec = new Rectangle(MFGDate_txtX, MFGDate_txtY, MFGDate_txtW, MFGDate_txtH);
                        Rectangle LifeSpanRec = new Rectangle(Lifespan_txtX, Lifespan_txtY, Lifespan_txtW, Lifespan_txtH);
                        Rectangle AdmissionDateRec = new Rectangle(AdmissionDate_txtX, AdmissionDate_txtY, AdmissionDate_txtW, AdmissionDate_txtH);
                        Rectangle MFGDateValRec = new Rectangle(MFGDateX, MFGDateY, MFGDateW, MFGDateH);
                        Rectangle LifeSpanValRec = new Rectangle(LifespanX, LifespanY, LifespanW, LifespanH);
                        Rectangle AdmissionDateValRec = new Rectangle(AdmissionDateX, AdmissionDateY, AdmissionDateW, AdmissionDateH);

                        Rectangle EXPDateRec = new Rectangle(EXPDate_txtX, EXPDate_txtY, EXPDate_txtW, EXPDate_txtH);
                        Rectangle StartWarrantyRec = new Rectangle(StartWarranty_txtX, StartWarranty_txtY, StartWarranty_txtW, StartWarranty_txtH);
                        Rectangle EndWarrantyRec = new Rectangle(EndWarranty_txtX, EndWarranty_txtY, EndWarranty_txtW, EndWarranty_txtH);
                        Rectangle EXPDateValRec = new Rectangle(EXPDateX, EXPDateY, EXPDateW, EXPDateH);
                        Rectangle StartWarrantyValRec = new Rectangle(StartWarrantyX, StartWarrantyY, StartWarrantyW, StartWarrantyH);
                        Rectangle EndWarrantyValRec = new Rectangle(EndWarrantyX, EndWarrantyY, EndWarrantyW, EndWarrantyH);

                        Rectangle WarrantyspanRec = new Rectangle(WarrantySpanX, WarrantySpanY, WarrantySpanW, WarrantySpanH);

                        Rectangle StatusRec = new Rectangle(Status_txtX, Status_txtY, Status_txtW, Status_txtH);
                        Rectangle StatusValRec = new Rectangle(StatusX, StatusY, StatusW, StatusH);

                        Rectangle FinalExpRec = new Rectangle(FinalExpiration_txtX, FinalExpiration_txtY, FinalExpiration_txtW, FinalExpiration_txtH);
                        Rectangle FinalEXPValRec = new Rectangle(FinalExpirationX, FinalExpirationY, FinalExpirationW, FinalExpirationH);

                        Rectangle NoteRec = new Rectangle(DetailsX, DetailsY, DetailsW, DetailsH);
                        Rectangle QRfillRec = new Rectangle(QRX, QRY, QRFW, QRFH);


                        Warrantyimg = new System.Drawing.Bitmap(1063, 827);
                        using (System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(Warrantyimg))
                        {
                            graphics.FillRectangle(Colorwhite, 0, 0, BackgroundSizeW, BackgroundSizeH);
                            graphics.DrawRectangle(blackPen, HeaderRec);
                            graphics.DrawString(Header, HeaderFont, Headerbrush, HeaderRec, stringFormat);
                            graphics.DrawString(Lot_txt, Lot_txtFont, Lot_txtbrush, Lot_txtX, Lot_txtY);
                            graphics.DrawString(Lot, LotFont, Lotbrush, LotX, LotY);

                            graphics.DrawString(Description_txt, Description_txtFont, Description_txtbrush, Description_txtX, Description_txtY);
                            graphics.DrawString(Description, DescriptionFont, Descriptionbrush, DescriptionX, DescriptionY);
                            graphics.DrawString(MFGSN_txt, MFGSN_txtFont, MFGSN_txtbrush, MFGSN_txtX, MFGSN_txtY);
                            graphics.DrawString(MFGSN, MFGSNFont, MFGSNbrush, MFGSNX, MFGSNY);

                            graphics.DrawString(WarehouseLocation_txt, WarehouseLocation_txtFont, WarehouseLocation_txtbrush, WarehouseLocRec, stringFormat);
                            graphics.DrawString(FixedAssetLocation_txt, FixedAssetLocation_txtFont, FixedAssetLocation_txtbrush, FixedAssetRec, stringFormat);
                            graphics.DrawString(WarehouseLocation, WarehouseLocationFont, WarehouseLocationbrush, WarehouseLocValRec, stringFormat);
                            graphics.DrawString(FixedAssetLocation, FixedAssetLocationFont, FixedAssetLocationbrush, FixedAssetValRec, stringFormat);

                            graphics.DrawString(MFGDate_txt, MFGDate_txtFont, MFGDate_txtbrush, MFGDateRec, stringFormat);
                            graphics.DrawString(AdmissionDate_txt, AdmissionDate_txtFont, AdmissionDate_txtbrush, AdmissionDateRec, stringFormat);
                            graphics.DrawString(Lifespan_txt, Lifespan_txtFont, Lifespan_txtbrush, LifeSpanRec, stringFormat);
                            graphics.DrawString(MFGDate, MFGDateFont, MFGDatebrush, MFGDateValRec, stringFormat);
                            graphics.DrawString(Lifespan, LifespanFont, Lifespanbrush, LifeSpanValRec, stringFormat);
                            graphics.DrawString(AdmissionDate, AdmissionDateFont, AdmissionDatebrush, AdmissionDateValRec, stringFormat);

                            graphics.DrawString(EXPDate_txt, EXPDate_txtFont, EXPDate_txtbrush, EXPDateRec, stringFormat);
                            graphics.DrawString(EXPDate, EXPDateFont, EXPDatebrush, EXPDateValRec, stringFormat);
                            graphics.DrawString(StartWarranty_txt, StartWarranty_txtFont, StartWarranty_txtbrush, StartWarrantyRec, stringFormat);
                            graphics.DrawString(EndWarranty_txt, EndWarranty_txtFont, EndWarranty_txtbrush, EndWarrantyRec, stringFormat);
                            graphics.DrawString(StartWarranty, StartWarrantyFont, StartWarrantybrush, StartWarrantyValRec, stringFormat);
                            graphics.DrawString(EndWarranty, EndWarrantyFont, EndWarrantybrush, EndWarrantyValRec, stringFormat);

                            graphics.DrawString(WarrantySpan + " YEAR/S", WarrantySpanFont, WarrantySpanbrush, WarrantyspanRec, stringFormat);
                            graphics.DrawRectangle(blackPen, WarrantyspanRec);

                            graphics.DrawString(Status, StatusFont, Statusbrush, StatusValRec, stringFormat);
                            graphics.DrawString(Status_txt, Status_txtFont, Status_txtbrush, StatusRec, stringFormat);
                            graphics.DrawString(FinalExpiration_txt, FinalExpiration_txtFont, FinalExpiration_txtbrush, FinalExpRec, stringFormat);
                            graphics.DrawRectangle(blackPen, FinalEXPValRec);
                            graphics.DrawString(FinalExpiration, FinalExpirationFont, FinalExpirationbrush, FinalEXPValRec, stringFormat);

                            graphics.DrawImage(QRimg, new PointF(QRX, QRY));
                            graphics.DrawLine(blackPen, LineX1, LineY1, LineX2, LineY2);
                            graphics.DrawString(Details_txt, Details_txtFont, Details_txtbrush, Details_txtX, Details_txtY);
                            graphics.DrawString(Details, DetailsFont, Detailsbrush, NoteRec, stringFormat1);
                        }

                        //for resizing
                        var destRect = new Rectangle(0, 0, PapersizeW, PapersizeH);
                        var destImage = new Bitmap(PapersizeW, PapersizeH);

                        destImage.SetResolution(Warrantyimg.HorizontalResolution, Warrantyimg.VerticalResolution);
                        using (var graphics = Graphics.FromImage(destImage))
                        {
                            graphics.CompositingMode = CompositingMode.SourceCopy;
                            graphics.CompositingQuality = CompositingQuality.HighQuality;
                            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            graphics.SmoothingMode = SmoothingMode.HighQuality;
                            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                            using (var wrapMode = new ImageAttributes())
                            {
                                wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                                graphics.DrawImage(Warrantyimg, destRect, 0, 0, Warrantyimg.Width, Warrantyimg.Height, GraphicsUnit.Pixel, wrapMode);
                            }
                        }

                        //for printing text file

                        string GeneratedImagePath = ImagePath;
                        string Date = DateTime.Now.ToString("MM-dd-yyyy");
                        string saveimgpath = GeneratedImagePath + @"\" + Date + "-" + QR;

                        destImage.Save(saveimgpath + ".bmp", ImageFormat.Bmp);
                        SetLabelText(string.Format(QR + " Successfully Generated Image and saved to folder"));
                        SystemSetup.SuccessAppend(string.Format(QR + " Successfully Generated Image"));

                        using (PrintDocument pd = new PrintDocument())
                        {                                                                       
                            //for paper size
                            PaperSize paperSize = new PaperSize("WarrantySticker", 340, 264);
                            paperSize.RawKind = (int)PaperKind.Custom;

                            pd.DefaultPageSettings.PaperSize = paperSize;

                            string filePath = GeneratedImagePath + @"\" + Date + "-" + QR + ".Bmp";
                            pd.OriginAtMargins = true;
                            //choosing of printer
                            pd.PrinterSettings.PrinterName = "Honeywell PM42 (203 dpi)";
                            pd.DefaultPageSettings.Margins.Top = 0;
                            pd.DefaultPageSettings.Margins.Bottom = 0;
                            pd.DefaultPageSettings.Margins.Left = 0;
                            pd.DefaultPageSettings.Margins.Right = 0;
                            pd.DefaultPageSettings.Landscape = false;
                            pd.DocumentName = filePath;
                            pd.PrintPage += printDocument1_PrintPage;
                            pd.Print();
                            pd.PrintPage -= printDocument1_PrintPage;

                            string time = DateTime.Now.ToString("HH");
                            string Success = successPath;

                            destImage.Save(Success + @"\" + Date + "_" + time + QR + ".bmp", ImageFormat.Bmp);
                            SetLabelText(string.Format(QR + " successfully Printed!"));
                            SystemSetup.SuccessAppend(string.Format(QR + " successfully Printed!"));
                        }
                    }
                    //moving of text file to success
                    string Success1 = successPath;
                    string Date1 = DateTime.Now.ToString("MM-dd-yyyy");
                    string XMLName = Path.GetFileName(XML);
                    string MoveTextPath = Success1 + @"\" + Date1 + "_" + XMLName;
                    if (File.Exists(MoveTextPath))
                    {
                     
                        File.Delete(MoveTextPath);
                    }
                    File.Move(XML, MoveTextPath);
                }

                return true;
            }

            catch (Exception ex)
            {
                SystemSetup.ErrorAppend(string.Format("Error Code: {0} " + "Error Desc: {1} " + "-999 " + ex.Message));
                GC.Collect();
                return false;
            }
        }
        //2480 x 3508
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (WarrantySticker() == true)
            {
                SetLabelText(string.Format("Finished generating warranty sticker.. Closing Program"));
            }
            Task.Delay(10000).Wait();
            this.Close();
        }
        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            string labelPath = ((PrintDocument)sender).DocumentName;
            System.Drawing.Image img = System.Drawing.Image.FromFile(labelPath);
            e.Graphics.DrawImage(img, e.MarginBounds);

            //Bitmap img = new Bitmap(labelPath);
            //e.Graphics.DrawImage(img, PrintRec);
        }
    }
}