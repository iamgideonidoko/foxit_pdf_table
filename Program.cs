using foxit;
using foxit.common;
using foxit.common.fxcrt;
using foxit.addon;
using foxit.pdf;
using foxit.pdf.annots;

using DotNetEnv;

namespace TablePDF
{
    internal class Program 
    {
        public static string output_path = "./output/pdf/";
        public static string DateTimeToString(foxit.common.DateTime datetime)
        {
            string s_datetime = string.Format("{0}/{1}/{2}-{3}:{4}:{5} {6}{7}:{8}", datetime.year, datetime.month, datetime.day,
                datetime.hour, datetime.minute, datetime.second, datetime.utc_hour_offset > 0 ? "+" : "-",
                datetime.utc_hour_offset,
                datetime.utc_minute_offset);
            return s_datetime;
        }

        public static foxit.common.DateTime GetLocalDateTime()
        {
            DateTimeOffset rime = DateTimeOffset.Now;
            foxit.common.DateTime datetime = new foxit.common.DateTime();
            datetime.year = (UInt16)rime.Year;
            datetime.month = (UInt16)rime.Month;
            datetime.day = (ushort)rime.Day;
            datetime.hour = (UInt16)rime.Hour;
            datetime.minute = (UInt16)rime.Minute;
            datetime.second = (UInt16)rime.Second;

            datetime.utc_hour_offset = (short)rime.Offset.Hours;
            datetime.utc_minute_offset = (ushort)rime.Offset.Minutes;

            return datetime;
        }

        public static void SetTableTextStyle(int index, RichTextStyle style)
        {
            using (style.font = new Font(Font.StandardID.e_StdIDHelvetica)) { }
            style.text_size = 10;
            style.text_alignment = Alignment.e_AlignmentLeft;
            style.text_color = 0x000000;
            style.is_bold = false;
            style.is_italic = false;
            style.is_underline = false;
            style.is_strikethrough = false;
            style.mark_style = RichTextStyle.CornerMarkStyle.e_CornerMarkNone;

            switch (index)
            {
                case 1:
                    style.text_alignment = Alignment.e_AlignmentCenter;
                    break;
                case 2:
                    {
                        style.text_alignment = Alignment.e_AlignmentRight;
                        style.text_color = 0x00FF00;
                        break;
                    }
                case 3:
                    style.text_size = 15;
                    break;
                case 4:
                    {
                        if (Environment.OSVersion.Platform == PlatformID.Unix)
                            using (style.font = new Font("Times New Roman", 0, Font.Charset.e_CharsetANSI, 0)) { }
                        else
                            using (style.font = new Font("Times New Roman", 0, Font.Charset.e_CharsetANSI, 0)) { }
                        style.text_color = 0xFF0000;
                        style.text_alignment = Alignment.e_AlignmentRight;
                        break;
                    }
                case 5:
                    {
                        if (Environment.OSVersion.Platform == PlatformID.Unix)
                            using (style.font = new Font("FreeSerif", 0, Font.Charset.e_CharsetANSI, 0)) { }
                        else
                            using (style.font = new Font("Times New Roman", 0, Font.Charset.e_CharsetANSI, 0))
                                style.is_bold = true;
                        style.text_alignment = Alignment.e_AlignmentRight;
                        break;
                    }
                case 6:
                    {
                        style.is_bold = true;
                        style.is_italic = true;
                        break;
                    }
                case 7:
                    {
                        style.is_bold = true;
                        style.is_italic = true;
                        style.text_alignment = Alignment.e_AlignmentCenter;
                        break;
                    }
                case 8:
                    {
                        style.is_underline = true;
                        style.text_alignment = Alignment.e_AlignmentRight;
                        break;
                    }
                case 9:
                    style.is_strikethrough = true;
                    break;
                case 10:
                    style.mark_style = RichTextStyle.CornerMarkStyle.e_CornerMarkSubscript;
                    break;
                case 11:
                    style.mark_style = RichTextStyle.CornerMarkStyle.e_CornerMarkSuperscript;
                    break;
                default:
                    break;
            }
        }

        public static string GetTableCellText(int index)
        {
            string cell_text = " ";
            switch (index)
            {
                case 0:
                    cell_text = "Reference style Ifex";
                    break;
                case 1:
                    cell_text = "Alignment center";
                    break;
                case 2:
                    cell_text = "Green text color and alignment right";
                    break;
                case 3:
                    cell_text = "Text font size 15";
                    break;
                case 4:
                    {
                        if (Environment.OSVersion.Platform == PlatformID.Unix)
                            cell_text = "Red text color, FreeSerif font and alignment right";
                        else
                            cell_text = "Red text color, Times New Roman font and alignment right";
                        break;
                    }
                case 5:
                    {
                        if (Environment.OSVersion.Platform == PlatformID.Unix)
                            cell_text = "Bold, FreeSerif font and alignment right";
                        else
                            cell_text = "Bold, Times New Roman font and alignment right";
                        break;
                    }
                case 6:
                    cell_text = "Bold and italic";
                    break;
                case 7:
                    cell_text = "Bold, italic and alignment center";
                    break;
                case 8:
                    cell_text = "Underline and alignment right";
                    break;
                case 9:
                    cell_text = "Strikethrough";
                    break;
                case 10:
                    cell_text = "CornerMarkSubscript";
                    break;
                case 11:
                    cell_text = "CornerMarkSuperscript";
                    break;
                default:
                    cell_text = " ";
                    break;
            }
            return cell_text;
        }

        public static void AddElectronicTable(PDFPage page)
        {
            // Add a spreadsheet with 4 rows and 3 columns
            {
                int index = 0;
                using (TableCellDataArray cell_array = new TableCellDataArray())
                {
                    for (int row = 0; row < 4; row++)
                    {
                        // nested using statements
                        using (RichTextStyle style = new RichTextStyle())
                        using (TableCellDataColArray col_array = new TableCellDataColArray())
                        {
                            for (int col = 0; col < 3; col++)
                            {
                                string cell_text = GetTableCellText(index);
                                SetTableTextStyle(index++, style);
                                Image image = new Image();
                                using (TableCellData cell_data = new TableCellData(style, cell_text, image, new RectF()))
                                {
                                    col_array.Add(cell_data);
                                }
                            }
                            cell_array.Add(col_array);
                        }
                    }
                    float page_width = page.GetWidth();
                    float page_height = page.GetHeight();
                    TableBorderInfo outside_border_left = new TableBorderInfo();
                    outside_border_left.line_width = 1;
                    outside_border_left.table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid;
                    TableBorderInfo outside_border_right = new TableBorderInfo();
                    outside_border_right.line_width = 1;
                    outside_border_right.table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid;
                    TableBorderInfo outside_border_top = new TableBorderInfo();
                    outside_border_top.line_width = 1;
                    outside_border_top.table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid;
                    TableBorderInfo outside_border_bottom = new TableBorderInfo();
                    outside_border_bottom.line_width = 1;
                    outside_border_bottom.table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid;
                    TableBorderInfo inside_border_row_info = new TableBorderInfo();
                    inside_border_row_info.line_width = 1;
                    inside_border_row_info.table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid;
                    TableBorderInfo inside_border_col_info = new TableBorderInfo();
                    inside_border_col_info.line_width = 1;
                    inside_border_col_info.table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid;
                    using (RectF rect = new RectF(100, 550, page_width - 100, page_height - 100))
                    using (TableData data = new TableData(rect, 4, 3, outside_border_left, outside_border_right, outside_border_top, outside_border_bottom, inside_border_row_info, inside_border_col_info, new TableCellIndexArray(), new FloatArray(), new FloatArray()))
                    {
                        TableGenerator.AddTableToPage(page, data, cell_array);
                    }
                }
            }

            //Add a spreadsheet with 5 rows and 6 columns
            {
                string cell_text = " ";
                string[] show_text = { "Foxit Software Incorporated", "Foxit Reader", "Foxit MobilePDF", "Foxit PhantomPDF", "Foxit PDF SDKs" };
                using (TableCellDataArray cell_array = new TableCellDataArray())
                {
                    for (int row = 0; row < 5; row++)
                    {
                        using (RichTextStyle style = new RichTextStyle())
                        using (TableCellDataColArray col_array = new TableCellDataColArray())
                        {
                            for (int col = 0; col < 6; col++)
                            {
                                if (col == 5)
                                    cell_text = DateTimeToString(GetLocalDateTime());
                                else
                                    cell_text = show_text[col];
                                SetTableTextStyle(row, style);
                                using (TableCellData cell_data = new TableCellData(style, cell_text, new Image(), new RectF()))
                                {
                                    col_array.Add(cell_data);
                                }
                            }
                            cell_array.Add(col_array);
                        }
                    }

                    float page_width = page.GetWidth();
                    float page_height = page.GetHeight();

                    TableBorderInfo outside_border_left = new TableBorderInfo();
                    outside_border_left.line_width = 2;
                    outside_border_left.table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid;
                    TableBorderInfo outside_border_right = new TableBorderInfo();
                    outside_border_right.line_width = 2;
                    outside_border_right.table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid;
                    TableBorderInfo outside_border_top = new TableBorderInfo();
                    outside_border_top.line_width = 2;
                    outside_border_top.table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid;
                    TableBorderInfo outside_border_bottom = new TableBorderInfo();
                    outside_border_bottom.line_width = 2;
                    outside_border_bottom.table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid;
                    TableBorderInfo inside_border_row_info = new TableBorderInfo();
                    inside_border_row_info.line_width = 2;
                    inside_border_row_info.table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid;
                    TableBorderInfo inside_border_col_info = new TableBorderInfo();
                    inside_border_col_info.line_width = 2;
                    inside_border_col_info.table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid;
                    using (RectF rect = new RectF(10, 200, page_width - 10, page_height - 350))
                    using (TableData data = new TableData(rect, 5, 6, outside_border_left, outside_border_right, outside_border_top, outside_border_bottom, inside_border_row_info, inside_border_col_info, new TableCellIndexArray(), new FloatArray(), new FloatArray()))
                    {
                        TableGenerator.AddTableToPage(page, data, cell_array);
                    }
                }
            }
        }
        static void Main(string[] args) {
            Env.Load();
            Console.WriteLine("TablePDF Main");

            Directory.CreateDirectory(output_path);

            string sn = Environment.GetEnvironmentVariable("FOXIT_SDK_SN") ?? "";
            string key = Environment.GetEnvironmentVariable("FOXIT_SDK_KEY") ?? "";

            Console.WriteLine($"sn => {sn}");

            // Initialize library
            ErrorCode error_code = Library.Initialize(sn, key);
            if (error_code != ErrorCode.e_ErrSuccess)
            {
                Console.WriteLine("Library Initialize Error: {0}\n", error_code);
                return;
            }

            try
            {
                using (PDFDoc doc = new PDFDoc())
                {
                    // Get first page with index 0
                    using (PDFPage page = doc.InsertPage(0, PDFPage.Size.e_SizeLetter))
                    {
                        AddElectronicTable(page);
                        // Save PDF file
                        string output_file = output_path + "TablePDF.pdf";
                        doc.SaveAs(output_file, (int)PDFDoc.SaveFlags.e_SaveFlagNoOriginal);
                        Console.WriteLine("electronictable demo.");
                    }
                }
            }
            catch (PDFException e)
            {
                Console.WriteLine(e.Message);
            }

            Library.Release();
        }
    }
}