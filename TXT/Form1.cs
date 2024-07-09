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
using System.Threading;
using System.Data.SqlClient;
using System.Data;

namespace WarrantySticker
{
    public partial class Form1 : Form
    {
        int FiveM;
        int loop = 1;
        string strINIPath = Application.StartupPath + "\\config.ini";
        string successPath = Application.StartupPath + @"\SuccessFile";
        string ImagePath = Application.StartupPath + @"\Generated Image";
		string SQLSettings = Application.StartupPath + "\\SAPconnsettings.ini";

		public Form1()
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
            Startbtn.Enabled = false;
            backgroundWorker1.RunWorkerAsync();
            CheckForIllegalCrossThreadCalls = false;
            backgroundWorker1.WorkerSupportsCancellation = true;
            backgroundWorker1.WorkerReportsProgress = true;
            DateTime now = DateTime.Now;
            DateTime when = new DateTime(now.Year, now.Month, now.Day, 01, 00, 00);
            if (now > when)
            {
                when = when.AddDays(1); // Assign the updated DateTime back to 'when'
            }

            timerend.Interval = (int)((when - now).TotalMilliseconds);
            timerend.Start();
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

		private bool QRsticker(string location)
		{
			List<System.Drawing.Bitmap> images = new List<System.Drawing.Bitmap>();
			System.Drawing.Bitmap Barcodeimg = null;
			var IniGet = new INIFile(strINIPath);
			var Printernamesmol = IniGet.Read("QRPrinter");
			string txtFiles = "*.txt"; // .txt
			string[] GetBarcodetxtFile = Directory.GetFiles(location, txtFiles);

			try
			{
				foreach (var TXT in GetBarcodetxtFile)
				{
					List<string> QRD = new List<string>();
					using (var reader = new StreamReader(TXT))
					{
						while (!reader.EndOfStream)
						{
							string value = reader.ReadLine();
							string[] spliter = value.Split('\n');
							string MFGSN = spliter[0].ToString();
							QRD.Add(MFGSN);

						}
						int groupSize = 3;
						int dataIndex = 0;

						while (dataIndex < QRD.Count)
						{
							List<string> subList = new List<string>();
							for (int i = 0; i < groupSize && dataIndex < QRD.Count; i++, dataIndex++)
							{
								subList.Add(QRD[dataIndex]);
							}
							string groupAsString = string.Join("|\n", subList);
							//string[] spliter = groupAsString.Split('|');
							string[] grp = groupAsString.Split(new string[] { "|\n" }, StringSplitOptions.None);

							string QRone = "";
							string QRtwo = "";
							string QRthree = "";
							Barcodeimg = new System.Drawing.Bitmap(940, 240);

							int barcodestickW = 940;
							int barcodestickH = 240;

							var Colorwhite = new SolidBrush(Color.White);

							using (System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(Barcodeimg))
							{
								graphics.FillRectangle(Colorwhite, 0, 0, barcodestickW, barcodestickH);
							}

							if (grp.Length > 0)
							{
								QRone = grp[0].ToString();
								SfBarcode barcode = new SfBarcode();

								barcode.Text = QRone;
								barcode.Symbology = BarcodeSymbolType.QRBarcode;

								QRBarcodeSetting setting = new QRBarcodeSetting();
								setting.XDimension = 3;
								barcode.SymbologySettings = setting;

								Image QRoneimg = barcode.ToImage(new SizeF(340, 240F));

								using (System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(Barcodeimg))
								{
									graphics.DrawImage(QRoneimg, new PointF(0, 0));
								}
							}


							if (grp.Length > 1)
							{
								QRtwo = grp[1].ToString();
								if (!string.IsNullOrEmpty(QRtwo))
								{
									SfBarcode barcode2 = new SfBarcode();
									barcode2.Text = QRtwo;
									barcode2.Symbology = BarcodeSymbolType.QRBarcode;

									QRBarcodeSetting setting = new QRBarcodeSetting();
									setting.XDimension = 3;
									barcode2.SymbologySettings = setting;

									Image QRtwoimg = barcode2.ToImage(new SizeF(340, 240F));

									using (System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(Barcodeimg))
									{
										graphics.DrawImage(QRtwoimg, new PointF(310, 0));
									}
								}
							}

							if (grp.Length > 2)
							{
								QRthree = grp[2].ToString();
								if (!string.IsNullOrEmpty(QRthree))
								{
									QRBarcodeSetting setting = new QRBarcodeSetting();
									setting.XDimension = 3;
									//for QRcode
									SfBarcode barcode3 = new SfBarcode();
									barcode3.Text = QRthree;
									barcode3.Symbology = BarcodeSymbolType.QRBarcode;
									barcode3.SymbologySettings = setting;
									Image QRthreeimg = barcode3.ToImage(new SizeF(340, 240F));

									using (System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(Barcodeimg))
									{
										graphics.DrawImage(QRthreeimg, new PointF(620, 0));
									}
								}
							}

							string GeneratedImagePath = ImagePath;
							string Date = DateTime.Now.ToString("MM-dd-yyyy");


							string saveimgpath = GeneratedImagePath + @"\" + $"{Date}" + ".bmp";

							if (File.Exists(saveimgpath))
							{
								SetLabelText(" there is a duplicated code.");
								SystemSetup.ErrorAppend(string.Format(" there is a duplicated code."));

								File.Delete(saveimgpath);
							}

							Barcodeimg.Save(saveimgpath, ImageFormat.Bmp);

							SetLabelText(string.Format(QRone + "-" + QRtwo + "-" + QRthree + " Successfully Generated Barcode Image and saved to folder"));
							SystemSetup.SuccessAppend(string.Format(QRone + "-" + QRtwo + "-" + QRthree + " Successfully Generated Barcode Image and saved to folder"));

							using (PrintDocument pd = new PrintDocument())
							{

								//for paper size
								PaperSize paperSize = new PaperSize("Shelf label", 355, 95);
								paperSize.RawKind = (int)PaperKind.Custom;

								pd.DefaultPageSettings.PaperSize = paperSize;

								string filePath = GeneratedImagePath + @"\" + $"{Date}" + ".Bmp";

								pd.OriginAtMargins = true;
								//choosing of printer
								pd.PrinterSettings.PrinterName = Printernamesmol;
								pd.PrinterSettings.DefaultPageSettings.Margins.Top = 0;
								pd.PrinterSettings.DefaultPageSettings.Margins.Bottom = 0;
								pd.PrinterSettings.DefaultPageSettings.Margins.Left = 0;
								pd.PrinterSettings.DefaultPageSettings.Margins.Right = 0;
								pd.DefaultPageSettings.Margins.Top = 0;
								pd.DefaultPageSettings.Margins.Bottom = 0;
								pd.DefaultPageSettings.Margins.Left = 0;
								pd.DefaultPageSettings.Margins.Right = 0;
								pd.DefaultPageSettings.Landscape = false;
								pd.DocumentName = filePath;
								pd.PrintPage += printDocument2_PrintPage;
								pd.Print();
								pd.PrintPage -= printDocument2_PrintPage;
								SetLabelText(string.Format(QRone + "," + QRtwo + "," + QRthree + " successfully Printed!"));
								SystemSetup.SuccessAppend(string.Format(QRone + "," + QRtwo + "," + QRthree + " successfully Printed!"));
								Thread.Sleep(2000);
							}
						}
					}
				}
				return true;
			}
			catch (Exception err)
			{
				SetLabelText(" Printing of QR Sticker Failed " + err.Message);
				SystemSetup.ErrorAppend(string.Format(" Printing of QR Sticker Failed " + err.Message));
				return false;
			}
		}

		private bool WarrantySticker(string location)
        {
            //universal
            List<System.Drawing.Bitmap> images = new List<System.Drawing.Bitmap>();
            System.Drawing.Bitmap Warrantyimg = null;
            string txtFiles = "*.txt";
            var IniGet = new INIFile(strINIPath);
            var Printernamebig = IniGet.Read("WarrantyPrinter");

            string[] GetTXTFiles = Directory.GetFiles(location, txtFiles);
            try
            {
                foreach (var TXT in GetTXTFiles)
                {
                    using (var reader = new StreamReader(TXT))
                    {
                        while (!reader.EndOfStream)
                        {
                            string value = reader.ReadLine();
                            string[] spliter = value.Split('|');
							string ItemCode = spliter.Length > 0 ? spliter[0].ToString() : " ";
							string Barcode = spliter.Length > 1 ? spliter[1].ToString() : " ";
							string Description = spliter.Length > 2 ? spliter[2].ToString() : " ";
							string MFGSN = spliter.Length > 3 ? spliter[3].ToString() : " ";
							string LOTNO = spliter.Length > 4 ? spliter[4].ToString() : " ";
							string LOCATION = spliter.Length > 5 ? spliter[5].ToString() : " ";
							string StartWarranty = spliter.Length > 6 ? spliter[6].ToString() : " ";
							string EndWarranty = spliter.Length > 7 ? spliter[7].ToString() : " ";
							string WarrantySpan = spliter.Length > 8 ? spliter[8].ToString() : " ";
							string ADMDate = spliter.Length > 9 ? spliter[9].ToString() : " ";
							string GRNTStart = spliter.Length > 10 ? spliter[10].ToString() : " ";
							string GRNTEXP = spliter.Length > 11 ? spliter[11].ToString() : " ";
							string Status = spliter.Length > 12 ? spliter[12].ToString() : " ";
							string MFGDate = spliter.Length > 13 ? spliter[13].ToString() : " ";
							string EXPDate = spliter.Length > 14 ? spliter[14].ToString() : " ";
							string LifeSpan = spliter.Length > 15 ? spliter[15].ToString() : " ";
							string Notes = spliter.Length > 16 ? spliter[16].ToString() : " ";
							string WarehouseLoc = spliter.Length > 17 ? spliter[17].ToString() : " ";
							string FinalEXP = spliter.Length > 18 ? spliter[18].ToString() : " ";

							int BackgroundSizeW = 1063;
                            int BackgroundSizeH = 827;
                            int PapersizeW = 900;
                            int PapersizeH = 700;

                            var Colortransparent = new SolidBrush(Color.Transparent);
                            var Colorwhite = new SolidBrush(Color.White);

                            //for QRcode
                            SfBarcode barcode = new SfBarcode();
                            barcode.Text = MFGSN;
                            barcode.Symbology = BarcodeSymbolType.QRBarcode;

                            //Set dimension for each block
                            QRBarcodeSetting setting = new QRBarcodeSetting();
                            setting.XDimension = 3;
                            barcode.SymbologySettings = setting;
                            //Export to image
                            Image QRimg = barcode.ToImage(new SizeF(230F, 230F));

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

                            Rectangle HeaderRec = new Rectangle(30, 20, 993, 25);

                            Rectangle WarehouseLocRec = new Rectangle(30, 170, 501, 30);
                            Rectangle FixedAssetRec = new Rectangle(531, 170, 501, 30);
                            Rectangle WarehouseLocValRec = new Rectangle(30, 200, 501, 30);
                            Rectangle FixedAssetValRec = new Rectangle(531, 200, 501, 30);

                            Rectangle MFGDateRec = new Rectangle(30, 240, 334, 25);
                            Rectangle LifeSpanRec = new Rectangle(698, 240, 334, 25);
                            Rectangle AdmissionDateRec = new Rectangle(364, 240, 335, 25);
                            Rectangle MFGDateValRec = new Rectangle(30, 270, 334, 25);
                            Rectangle LifeSpanValRec = new Rectangle(698, 270, 334, 25);
                            Rectangle AdmissionDateValRec = new Rectangle(364, 270, 335, 25);

                            Rectangle EXPDateRec = new Rectangle(30, 310, 334, 25);
                            Rectangle StartWarrantyRec = new Rectangle(364, 310, 335, 25);
                            Rectangle EndWarrantyRec = new Rectangle(698, 310, 335, 25);
                            Rectangle EXPDateValRec = new Rectangle(30, 340, 334, 25);
                            Rectangle StartWarrantyValRec = new Rectangle(364, 340, 335, 25);
                            Rectangle EndWarrantyValRec = new Rectangle(698, 340, 335, 25);

                            Rectangle WarrantyspanRec = new Rectangle(364, 390, 670, 30);

                            Rectangle StatusRec = new Rectangle(30, 380, 334, 25);
                            Rectangle StatusValRec = new Rectangle(30, 410, 334, 25);

                            Rectangle FinalExpRec = new Rectangle(30, 450, 335, 25);
                            Rectangle FinalEXPValRec = new Rectangle(364, 450, 670, 30);

                            Rectangle NoteRec = new Rectangle(160, 510, 690, 300);
                            Rectangle QRfillRec = new Rectangle(850, 510, 183, 210);

                            var BlackBrush = new SolidBrush(Color.Black);
                            var InviBrush = new SolidBrush(Color.Transparent);
                            var WhiteBrush = new SolidBrush(Color.White);
                            var ABFontS = new Font("Arial", 20, FontStyle.Bold);
                            var ABFont1 = new Font("Arial", 23, FontStyle.Bold);
                            var ARFont = new Font("Arial", 22, FontStyle.Regular);
                            //fonts for QR
                            var QRHeadFont = new Font("Arial", 9, FontStyle.Bold);
                            Warrantyimg = new System.Drawing.Bitmap(1063, 827);
                            SetLabelText("Generating Image");
                            using (System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(Warrantyimg))
                            {
                                graphics.FillRectangle(Colorwhite, 0, 0, BackgroundSizeW, BackgroundSizeH);
                                graphics.DrawLine(blackPen, 30, 10, 1033, 10);
                                graphics.DrawLine(blackPen, 1033, 10, 1033, 50);
                                graphics.DrawLine(blackPen, 1033, 50, 30, 50);
                                graphics.DrawLine(blackPen, 30, 50, 30, 10);
                                graphics.DrawString("REIC WARRANTY STICKER", ABFont1, BlackBrush, HeaderRec, stringFormat);
                                graphics.DrawString("LOT", ARFont, BlackBrush, 500, 50);
                                graphics.DrawString(LOTNO, ABFontS, BlackBrush, 560, 51);

                                graphics.DrawString("DESCRIPTION", ARFont, BlackBrush, 30, 90);
                                graphics.DrawString(Description, ABFontS, BlackBrush, 240, 89);
                                graphics.DrawString("MFG.S/N", ARFont, BlackBrush, 30, 129);
                                graphics.DrawString(MFGSN, ABFontS, BlackBrush, 240, 129);

                                graphics.DrawString("WAREHOUSE LOCATION", ARFont, BlackBrush, WarehouseLocRec, stringFormat);
                                graphics.DrawString("FIXED ASSET LOCATION", ARFont, BlackBrush, FixedAssetRec, stringFormat);
                                graphics.DrawString(WarehouseLoc, ABFontS, BlackBrush, WarehouseLocValRec, stringFormat);
                                graphics.DrawString(LOCATION, ABFontS, BlackBrush, FixedAssetValRec, stringFormat);

                                graphics.DrawString("MFG.DATE", ARFont, BlackBrush, MFGDateRec, stringFormat);
                                graphics.DrawString("ADMISSION DATE", ARFont, BlackBrush, AdmissionDateRec, stringFormat);
                                graphics.DrawString("LIFESPAN", ARFont, BlackBrush, LifeSpanRec, stringFormat);
                                graphics.DrawString(MFGDate, ABFont1, BlackBrush, MFGDateValRec, stringFormat);
                                graphics.DrawString(LifeSpan, ABFont1, BlackBrush, LifeSpanValRec, stringFormat);
                                graphics.DrawString(ADMDate, ABFont1, BlackBrush, AdmissionDateValRec, stringFormat);

                                graphics.DrawString("EXP.DATE", ARFont, BlackBrush, EXPDateRec, stringFormat);
                                graphics.DrawString(EXPDate, ABFont1, BlackBrush, EXPDateValRec, stringFormat);
                                graphics.DrawString("START WARRANTY", ARFont, BlackBrush, StartWarrantyRec, stringFormat);
                                graphics.DrawString("END WARRANTY", ARFont, BlackBrush, EndWarrantyRec, stringFormat);
                                graphics.DrawString(StartWarranty, ABFont1, BlackBrush, StartWarrantyValRec, stringFormat);
                                graphics.DrawString(EndWarranty, ABFont1, BlackBrush, EndWarrantyValRec, stringFormat);

                                graphics.DrawString(WarrantySpan + "Year/s", ABFont1, BlackBrush, WarrantyspanRec, stringFormat);
                                graphics.DrawRectangle(blackPen, WarrantyspanRec);

                                graphics.DrawString(Status, ABFont1, BlackBrush, StatusValRec, stringFormat);
                                graphics.DrawString("STATUS", ARFont, BlackBrush, StatusRec, stringFormat);
                                graphics.DrawString("FINAL EXPIRATION", ARFont, BlackBrush, FinalExpRec, stringFormat);
                                graphics.DrawRectangle(blackPen, FinalEXPValRec);
                                graphics.DrawString(FinalEXP, ABFont1, BlackBrush, FinalEXPValRec, stringFormat);

                                graphics.DrawImage(QRimg, new PointF(820, 530));
                                graphics.DrawLine(blackPen, 30, 500, 1033, 500);
                                graphics.DrawString("DETAILS", ARFont, BlackBrush, 30, 510);
                                graphics.DrawString(Notes, ABFont1, BlackBrush, NoteRec, stringFormat1);
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

                            string GeneratedImagePath = ImagePath;
                            string Date = DateTime.Now.ToString("MM-dd-yyyy");
                            string saveimgpath = GeneratedImagePath + @"\" + Date + "-" + MFGSN + ".bmp";

                            if (File.Exists(saveimgpath))
                            {
                                SetLabelText("there is a duplicated code.");
                                SystemSetup.ErrorAppend(string.Format("there is a duplicated code."));
                                File.Delete(saveimgpath);
                            }

                            destImage.Save(saveimgpath, ImageFormat.Bmp);
                            SetLabelText(string.Format(ItemCode + " Successfully Generated Image and saved to folder"));
                            SystemSetup.SuccessAppend(string.Format(ItemCode + " Successfully Generated Image"));

                            using (PrintDocument pd = new PrintDocument())
                            {
                                //for paper size
                                PaperSize paperSize = new PaperSize("WarrantySticker", 340, 264);
                                paperSize.RawKind = (int)PaperKind.Custom;

                                pd.DefaultPageSettings.PaperSize = paperSize;

                                string filePath = GeneratedImagePath + @"\" + Date + "-" + MFGSN + ".Bmp";
                                pd.OriginAtMargins = true;
                                //choosing of printer
                                pd.PrinterSettings.PrinterName = Printernamebig;
                                pd.DefaultPageSettings.Margins.Top = 0;
                                pd.DefaultPageSettings.Margins.Bottom = 0;
                                pd.DefaultPageSettings.Margins.Left = 0;
                                pd.DefaultPageSettings.Margins.Right = 0;
                                pd.DefaultPageSettings.Landscape = false;
                                pd.DocumentName = filePath;
                                pd.PrintPage += printDocument1_PrintPage;
                                pd.Print();
                                pd.PrintPage -= printDocument1_PrintPage;

                                SetLabelText(string.Format(MFGSN + " successfully Printed!"));
                                SystemSetup.SuccessAppend(string.Format(MFGSN + " successfully Printed!"));
                                //Thread.Sleep(2000);
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception err)
            {
                SetLabelText("Printing of Warranty Sticker Failed " + err.Message);
                SystemSetup.ErrorAppend(string.Format("Printing of Warranty Sticker Failed " + err.Message));
                return false;
            }
        }

            private bool MoveFiles(string location)
        {
            try
            {
                string txtFiles = "*.txt";
                string[] GettxtFile = Directory.GetFiles(location, txtFiles);
                //moving of text file to success
                string Success1 = successPath;
                string Date1 = DateTime.Now.ToString("MM-dd-yyyy");
                string time1 = DateTime.Now.ToString("HH");
                foreach (var txt in GettxtFile)
                {
                    string TxtName = Path.GetFileName(txt);
                    Directory.CreateDirectory(Success1 + @"\" + Date1 + "-" + time1 + @"\");
                    string MoveTextPath = Success1 + @"\" + Date1 + "-" + time1 + @"\" + TxtName;

                    if (File.Exists(MoveTextPath))
                    {
                        File.Delete(MoveTextPath);
                    }
                    File.Move(txt, MoveTextPath);
                    File.Delete(txt);
                }
                return true;
            }
            catch (Exception err)                             
            {
                SetLabelText("Moving file to success failed " + err.Message);
                SystemSetup.ErrorAppend(string.Format("Moving file to success failed " + err.Message));
                return false;
            }
        }

        private bool Process1()
        {
            try
            {
                while (loop == 1)
                {
                    var IniGet = new INIFile(strINIPath);
                    var WarrantyPath = IniGet.Read("WarrantyPath");
                    var QRPath = IniGet.Read("QRPath");
                    //var BothPath = IniGet.Read("BothPath");
                    string txtFiles = "*.txt";
                    string[] GetWarrantyFile = Directory.GetFiles(WarrantyPath, txtFiles);
                    string[] GetQRFile = Directory.GetFiles(QRPath, txtFiles);
                    //string[] GetBothFile = Directory.GetFiles(BothPath, txtFiles);

                    if (GetWarrantyFile.Length > 0)
                    {
                        SetLabelText(string.Format("TXT file found for Warranty Sticker. Starting Process.."));
                        SystemSetup.SuccessAppend(string.Format("TXT file found for Warranty Sticker. Starting Process.."));
                        WarrantySticker(WarrantyPath);
                        MoveFiles(WarrantyPath);
                        SetLabelText(string.Format("Finish printing all warranty stickers"));
                        SystemSetup.SuccessAppend(string.Format("Finish printing all warranty stickers"));
                    }
                        
                    if(GetQRFile.Length > 0)
                    {
                        SetLabelText(string.Format("TXT file found for QR Sticker. Starting Process.."));
                        SystemSetup.SuccessAppend(string.Format("TXT file found for QR Sticker. Starting Process.."));
                        QRsticker(QRPath);
                        MoveFiles(QRPath);
                        SetLabelText(string.Format("finish printing all QR Stickers"));
                        SystemSetup.SuccessAppend(string.Format("finish printing all QR stickers."));
                    }
                   
                    //if(GetBothFile.Length > 0)
                    //{
                    //    SetLabelText(string.Format("TXT file found for Warranty and QR Sticker. Starting Process.."));
                    //    SystemSetup.SuccessAppend(string.Format("TXT file found for Warranty and QR Sticker. Starting Process.."));
                    //    WarrantySticker(BothPath);
                    //    QRsticker(BothPath);
                    //    MoveFiles(BothPath);
                    //    SetLabelText(string.Format("Finish printing QR and warranty stickers"));
                    //    SystemSetup.SuccessAppend(string.Format("Finish printing QR and warranty stickers"));
                    //}

                    string Bmpfiles = "*.bmp";
                    string[] Getbmpfile = Directory.GetFiles(ImagePath, Bmpfiles);
                    foreach (var bmp in Getbmpfile)
                    {
                        File.Delete(bmp);
                    }
                    loop = 0;
                }
                return true;
            }
            catch (Exception err)
            {
                SetLabelText("Printing Failed " + err.Message);
                SystemSetup.ErrorAppend(string.Format(err.Message));
                return false;
            }
        }
        //2480 x 3508
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            ConnectToSQL();
            Process1();

        }

        private void ConnectToSQL()
        {
			StreamReader sr = null;
			string strConn;

			try
			{
				sr = new StreamReader(SQLSettings);

				string Server = sr.ReadLine().Trim();
				string DbUserName = sr.ReadLine().Trim();
				string DbPassword = sr.ReadLine().Trim();
				string CompanyDB = sr.ReadLine().Trim();

				string server = Server.Substring(Server.IndexOf("=") + 1);
				string dbUserName = DbUserName.Substring(DbUserName.IndexOf("=") + 1);
				string dbPassword = DbPassword.Substring(DbPassword.IndexOf("=") + 1);
				string companyDB = CompanyDB.Substring(CompanyDB.IndexOf("=") + 1);

				sr.Close();

				strConn = "Data Source=" + server + ";" +
						  "Initial Catalog=" + companyDB + ";" +
						  "User ID=" + dbUserName + ";" +
						  "Password=" + dbPassword;

				using (SqlConnection SQLConnDB = new SqlConnection())
				{
					try
					{
						SQLConnDB.ConnectionString = strConn;
						SQLConnDB.Open();
						Thread.Sleep(4000);
						SetLabelText(string.Format("Successfully connected to the SQL. Reading files, please wait..."));
						SystemSetup.SuccessAppend(string.Format("Successfully connected to the SQL. Reading files, please wait..."));
						Thread.Sleep(3000);
						ReadAndMove(SQLConnDB);
					}
					catch (Exception err)
					{
						SetLabelText(string.Format(" Check your SQL Connection Settings. " + err.Message));
						SystemSetup.SuccessAppend(string.Format(" Check your SQL Connection Settings. " + err.Message));
					}
				}
			}
			catch (Exception err)
			{
				sr = null;
				//MessageBox.Show("Please setup the SQL Server Connection Settings.", "SQL Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				SetLabelText(string.Format(" Please setup the SQL Server Connection Settings. ", "SQL Connection Error" + err.Message));
				SystemSetup.SuccessAppend(string.Format(" Please setup the SQL Server Connection Settings.", "SQL Connection Error" + err.Message));
				return;
			}
		}

		private void ReadAndMove(SqlConnection SQLConnDB)
		{
			var IniGet = new INIFile(strINIPath);
			var TxtPath = IniGet.Read("Warranty1");
			var WarrantyPath = IniGet.Read("WarrantyPath");
			var QRPath = IniGet.Read("QRPath");
			var txt = "*.txt";

			StringBuilder qr = new StringBuilder();
			StringBuilder warr = new StringBuilder();

			DataTable dt = new DataTable();

			try
			{
				string[] GetFiles = Directory.GetFiles(TxtPath, txt);

				foreach (string GetFile in GetFiles)
				{
					string[] lines = File.ReadAllLines(GetFile);

					foreach (string line in lines)
					{
						//string[] val = line.Split(new string[] { " | " }, StringSplitOptions.None);
						string[] val = line.Split('|');
						string ItemCode = val[0];
					    //string slt = "SELECT OITM.U_ITMUDF9 FROM OITM WHERE OITM.ItemCode='" + ItemCode + "'";
						string slt = "SELECT qr_warranty_sticker FROM product_item WHERE item_code='" + ItemCode + "'";

						try
						{
							SqlDataAdapter adp = new SqlDataAdapter(slt, SQLConnDB);
							adp.Fill(dt);

							if (dt.Rows.Count > 0)
							{
								foreach (DataRow row in dt.Rows)
								{
									//string udf = row["U_ITMUDF9"].ToString();
									string udf = row["qr_warranty_sticker"].ToString();

									if (udf == "0")
									{
										SetLabelText(string.Format("No print."));
										SystemSetup.SuccessAppend(string.Format("No print."));
									}
									else if (udf == "1")
									{
										qr.AppendLine(line);
										SetLabelText(string.Format("QR Code Stickers are ready to print."));
										SystemSetup.SuccessAppend(string.Format("QR Code Stickers are ready to print"));
									}
									else if (udf == "2")
									{
										warr.AppendLine(line);
										SetLabelText(string.Format("Detailed Stickers are ready to print."));
										SystemSetup.SuccessAppend(string.Format("Detailed Stickers are ready to print"));
									}
									else if (udf == "3")
									{
										qr.AppendLine(line);
										warr.AppendLine(line);
										SetLabelText("Both QR Code and Detailed Stickers are ready to print.");
										SystemSetup.SuccessAppend("Both QR Code and Detailed Stickers are ready to print.");
									}
									else
									{
										SetLabelText(string.Format("Invalid value: " + udf));
										SystemSetup.ErrorAppend(string.Format("Invalid value: " + udf));
									}
								}
							}
						}
						catch (SqlException err)
						{
							SetLabelText(string.Format("Error executing SQL query: " + err.Message));
							SystemSetup.SuccessAppend(string.Format("Error executing SQL query: " + err.Message));
						}
					}
				}

				if (qr.Length > 0)
				{
					string qrFileName = "QRsticker.txt";
					string destQRPath = Path.Combine(QRPath, qrFileName);
					File.WriteAllText(destQRPath, qr.ToString());
					SetLabelText(string.Format("QR Sticker txt file successfully created."));
					SystemSetup.SuccessAppend(string.Format("QR Sticker txt file successfully created."));
				}

				if (warr.Length > 0)
				{
					string warrFileName = "Detailedsticker.txt";
					string destWarrPath = Path.Combine(WarrantyPath, warrFileName);
					File.WriteAllText(destWarrPath, warr.ToString());
					SetLabelText(string.Format("Detailed Sticker txt file successfully created."));
					SystemSetup.SuccessAppend(string.Format("Detailed Sticker txt file successfully created."));
				}
			}
			catch (Exception err)
			{
				SetLabelText(string.Format("Failed to read or process files: " + err.Message));
				SystemSetup.ErrorAppend(string.Format("Failed to read or process files: " + err.Message));
			}
		}
		private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            loop = 1;
            FiveM = 900;
            timer1.Interval = 1000;
            timer1.Start();
        }

        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            string labelPath = ((PrintDocument)sender).DocumentName;
            System.Drawing.Image img = System.Drawing.Image.FromFile(labelPath);
            e.Graphics.DrawImage(img, e.MarginBounds);
            img.Dispose();
            //Bitmap img = new Bitmap(labelPath);
            //e.Graphics.DrawImage(img, PrintRec);
        }

        private void printDocument2_PrintPage(object sender, PrintPageEventArgs e)
        {
            string labelPath = ((PrintDocument)sender).DocumentName;
            System.Drawing.Image img = System.Drawing.Image.FromFile(labelPath);
            e.Graphics.DrawImage(img, e.MarginBounds);
            img.Dispose();
        }

        //private void timerend_Tick(object sender, EventArgs e)
        //{
        //    this.Close();
        //}

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
            NotifyIcon.Visible = false;
        }


        private void Form1_Shown_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            NotifyIcon.Visible = true;
            this.ShowInTaskbar = false;
            NotifyIcon.Text = "WarrantySticker-Running";
        }

        private void Form1_Resize_1(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                NotifyIcon.Visible = true;
                this.ShowInTaskbar = false;
                NotifyIcon.Text = "WarrantySticker-Running";
            }
        }

        private void Stopbtn_Click(object sender, EventArgs e)
        {
            this.Text = "Warranty Sticker Stopped";
            Stopbtn.Enabled = false;
            Startbtn.Enabled = true;
            backgroundWorker1.CancelAsync();
            timer1.Stop();
            timer1.Dispose();
            loop = 0;
            SetLabelText(string.Format("Application Stopped"));
            SystemSetup.SuccessAppend(string.Format("Application Stopped"));
        }

        private void Startbtn_Click(object sender, EventArgs e)
        {
            Stopbtn.Enabled = true;
            Startbtn.Enabled = false;
            loop = 1;
            SetLabelText(string.Format("Application Started"));
            SystemSetup.SuccessAppend(string.Format("Application Started"));
            backgroundWorker1.RunWorkerAsync();

        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {
            if (FiveM > 0)
            {
                string countDown = TimeSpan.FromSeconds(FiveM--).ToString(@"mm\:ss");
                this.Text = "Warranty Sticker       " + countDown;
            }
            else
            {
                timer1.Dispose();
                backgroundWorker1.RunWorkerAsync();
            }
        }

        private void timerend_Tick(object sender, EventArgs e)
        {
            Application.Restart();          
        }
    }
}

//for barcode

//foreach (var TXT in GetBarcodetxtFile)
//{
//    using (var reader = new StreamReader(TXT))
//    {
//        while (!reader.EndOfStream)
//        {
//            string value = reader.ReadLine();
//            string[] spliter = value.Split('|');
//            string ItemCode = spliter[0].ToString();
//            string MFGSN = spliter[3].ToString();
//            string Description = spliter[2].ToString();

//            //for smol size
//            int barcodestickW = 900;
//            int barcodestickH = 240;


//            Rectangle MFGSNrec = new Rectangle(10, 180, 900, 40);

//            //Generate Barcode
//            SfBarcode createBarcode = new SfBarcode();
//            createBarcode.LightBarColor = Color.Transparent;
//            createBarcode.TextGapHeight = 0;
//            createBarcode.DisplayText = false;
//            createBarcode.Text = MFGSN;
//            createBarcode.Symbology = BarcodeSymbolType.Code128A;
//            QRBarcodeSetting setting1 = new QRBarcodeSetting();
//            setting1.XDimension = 860;
//            createBarcode.SymbologySettings = setting1;
//            Image Codeimg = createBarcode.ToImage();

//            var Colorwhite = new SolidBrush(Color.White);
//            var BlackBrush = new SolidBrush(Color.Black);

//            var ARFont = new Font("Arial", 22, FontStyle.Regular);
//            var ABFontS = new Font("Arial", 20, FontStyle.Bold);
//            var MFGFont = new Font("Arial", 30, FontStyle.Bold);

//            //for Center Text
//            StringFormat stringFormat = new StringFormat();
//            stringFormat.Alignment = StringAlignment.Center;
//            stringFormat.LineAlignment = StringAlignment.Center;

//            Barcodeimg = new System.Drawing.Bitmap(900, 240);
//            using (System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(Barcodeimg))
//            {
//                graphics.FillRectangle(Colorwhite, 0, 0, barcodestickW, barcodestickH);
//                graphics.DrawString("DESCRIPTION:", ARFont, BlackBrush, 10, 10);
//                graphics.DrawString(Description, ABFontS, BlackBrush, 218, 12);
//                graphics.DrawImage(Codeimg, 100, 50, 700, 120);
//                graphics.DrawString(MFGSN, MFGFont, BlackBrush, MFGSNrec, stringFormat);
//            }
//            string GeneratedImagePath = GeneratedBarPath;
//            string Date = DateTime.Now.ToString("MM-dd-yyyy");

//            string saveimgpath = GeneratedImagePath + @"\" + Date + "-" + MFGSN + ".bmp";

//            if (File.Exists(saveimgpath))
//            {
//                File.Delete(saveimgpath);
//            }

//            Barcodeimg.Save(saveimgpath, ImageFormat.Bmp);
//            SetLabelText(string.Format(ItemCode + " Successfully Generated Barcode Image and saved to folder"));
//            SystemSetup.SuccessAppend(string.Format(ItemCode + " Successfully Generated Barcode Image and saved to folder"));

//        }
//    }
//}