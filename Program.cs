using System.Text;

using foxit;
using foxit.common;
using foxit.common.fxcrt;
using foxit.addon;
using foxit.pdf;
using foxit.pdf.annots;

using DotNetEnv;
using ExcelDataReader;
using System.Data;
using foxit.pdf.interform;

namespace TablePDF
{
    internal class Program 
    {
        public static readonly string output_path = "./output/pdf/";
        public static readonly string data_path = "./data/";

        private static readonly bool is_unix_platform = Environment.OSVersion.Platform == PlatformID.Unix;

        private static int row_count;
        private static int col_count;
        private static int num_of_pages;
        private static readonly int row_per_page = 10;
        private static readonly int col_per_page = 8;
        private static DataSet? dataSet;

        public static void GetTables(DataSet dataSet)
        {
            // Get Each DataTable in the DataTableCollection and
            // print each row value.
            foreach (DataTable table in dataSet.Tables)
                foreach (DataRow row in table.Rows)
                    foreach (DataColumn column in table.Columns)
                        if (row[column] != null)
                            Console.WriteLine(row[column]);
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
                        if (is_unix_platform)
                            using (style.font = new Font("Times New Roman", 0, Font.Charset.e_CharsetANSI, 0)) { }
                        else
                            using (style.font = new Font("Times New Roman", 0, Font.Charset.e_CharsetANSI, 0)) { }
                        style.text_color = 0xFF0000;
                        style.text_alignment = Alignment.e_AlignmentRight;
                        break;
                    }
                case 5:
                    {
                        if (is_unix_platform)
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

        public static void AddElectronicTable(PDFPage page, int page_index)
        {
           
            //Add a spreadsheet with 5 rows and 6 columns
            {
                string[] show_text = { "Foxit Software Incorporated", "Foxit Reader", "Foxit MobilePDF", "Foxit PhantomPDF", "Foxit PDF SDKs", "Col 6" };
                Random rand = new();
                using TableCellDataArray cell_array = new();

                DataTable? table = dataSet?.Tables[0];
                if (dataSet != null)
                {
                    GetTables(dataSet);
                }

                int row_start = 0 * (page_index + 1);
                int row_end = (page_index + 1) * 10;
                int actual_row_end = row_end > row_count ? row_count : row_end;

                for (int row = row_start; row < actual_row_end; row++)
                {
                    using RichTextStyle style = new();
                    using TableCellDataColArray col_array = new();
                    for (int col = 0; col < col_count; col++)
                    {
                        DataRow? actual_row = table?.Rows[row];
                        DataColumn? actual_column = table?.Columns[col];
                        if (actual_row != null && actual_column != null && actual_row[actual_column] != null)
                        {
                            string cell_text = $"{actual_row[actual_column]}";
                            Console.WriteLine($"cell_text => {actual_row}");
                            SetTableTextStyle(row, style);
                            using TableCellData cell_data = new(style, cell_text, new Image(), new RectF());
                            col_array.Add(cell_data);
                        }
                    }
                    cell_array.Add(col_array);
                }

                /*for (int row = 0; row < row_count; row++)
                {
                    using RichTextStyle style = new();
                    using TableCellDataColArray col_array = new();
                    for (int col = 0; col < col_count; col++)
                    {
                        string cell_text = show_text[rand.Next(0, show_text.Length)];
                        SetTableTextStyle(row, style);
                        using TableCellData cell_data = new(style, cell_text, new Image(), new RectF());
                        col_array.Add(cell_data);
                    }
                    cell_array.Add(col_array);
                }*/

                float page_width = page.GetWidth();
                float page_height = page.GetHeight();

                Console.WriteLine($"Page width => {page_width}");
                Console.WriteLine($"Page height => {page_height}");

                TableBorderInfo outside_border_left = new()
                {
                    line_width = 2,
                    table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid
                };
                TableBorderInfo outside_border_right = new()
                {
                    line_width = 2,
                    table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid
                };
                TableBorderInfo outside_border_top = new()
                {
                    line_width = 2,
                    table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid
                };
                TableBorderInfo outside_border_bottom = new()
                {
                    line_width = 2,
                    table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid
                };
                TableBorderInfo inside_border_row_info = new()
                {
                    line_width = 2,
                    table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid
                };
                TableBorderInfo inside_border_col_info = new()
                {
                    line_width = 2,
                    table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid
                };
                // using RectF rect = new(10, 200, page_width - 10, page_height - 350);
                using RectF rect = new(10, 200, page_width - 10, page_height - 40);
                using TableData data = new(rect, row_count, col_count, outside_border_left, outside_border_right, outside_border_top, outside_border_bottom, inside_border_row_info, inside_border_col_info, new TableCellIndexArray(), new FloatArray(), new FloatArray());
                // TableGenerator.AddTableToPage(page, data, cell_array);
            }
        }
        static void Main(string[] args) {
            try
            {
                // Add encoding for ExcelReader support
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                string input_file = data_path + "serious-injury-outcome-indicators-2000-2020.xlsx";

                using var stream = File.Open(input_file, FileMode.Open, FileAccess.Read);

                using var reader = ExcelReaderFactory.CreateReader(stream);
                row_count = reader.RowCount;
                col_count = reader.FieldCount > col_per_page ? col_per_page : reader.FieldCount;
                decimal pages = reader.RowCount / row_per_page;
                num_of_pages = (int) Math.Ceiling(pages);

                dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    UseColumnDataType = false,
                });

                // GetTables(result);

            } catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.GetType()} : {ex.Message}");
            }

            Env.Load();

            Directory.CreateDirectory(output_path);

            string sn = Environment.GetEnvironmentVariable("FOXIT_SDK_SN") ?? "";
            string key = Environment.GetEnvironmentVariable("FOXIT_SDK_KEY") ?? "";

            Console.WriteLine($"sn => {sn}");

            // Initialize Foxit library
            ErrorCode error_code = Library.Initialize(sn, key);
            if (error_code != ErrorCode.e_ErrSuccess)
            {
                Console.WriteLine("Library Initialize Error: {0}\n", error_code);
                return;
            }

            try
            {
                using PDFDoc doc = new();

                for (int i = 0; i < num_of_pages; i++)
                {
                    using PDFPage page = doc.InsertPage(i, PDFPage.Size.e_SizeLetter);
                    AddElectronicTable(page, i);

                }

                // page2.AddText("TESTING SOME THINGS OUT", new RectF(10, 200, page.GetWidth() - 10, page.GetHeight() - 40), new RichTextStyle());


                // Save PDF file
                string output_file = output_path + "TablePDF.pdf";
                doc.SaveAs(output_file, (int)PDFDoc.SaveFlags.e_SaveFlagNoOriginal);
                Console.WriteLine("Done.");
            }
            catch (PDFException e)
            {
                Console.WriteLine(e.Message);
            }

            Library.Release();
        }
    }
}