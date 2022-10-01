using System.Text;
using System.Data;

using foxit;
using foxit.common;
using foxit.common.fxcrt;
using foxit.addon;
using foxit.pdf;
using foxit.pdf.annots;

using DotNetEnv;
using ExcelDataReader;

namespace TablePDF
{
    internal class Program 
    {
        public static readonly string output_path = "./output/pdf/";
        public static readonly string data_path = "./data/";

        private static int row_count;
        private static int col_count;
        private static int num_of_pages;
        private static readonly int row_per_page = 10;
        private static readonly int col_per_page = 8;
        private static DataSet? dataSet;

        public static void SetTableTextStyle(int index, RichTextStyle style)
        {
            using (style.font = new Font(Font.StandardID.e_StdIDHelvetica)) { }
            style.text_size = 10;
            style.text_alignment = Alignment.e_AlignmentLeft;
            style.text_color = 0x000000;
            style.is_bold = index == 0;
            style.is_italic = false;
            style.is_underline = false;
            style.is_strikethrough = false;
            style.mark_style = RichTextStyle.CornerMarkStyle.e_CornerMarkNone;
        }

        public static void AddElectronicTable(PDFPage page, int page_index)
        {
           
            {
                using TableCellDataArray cell_array = new();

                DataTable? table = dataSet?.Tables[0];

                // Loop bounds
                int row_start = row_per_page * page_index;
                int row_end = row_start + row_per_page;
                int actual_row_end = row_end > row_count ? row_count : row_end;

                for (int row = row_start; row < actual_row_end; row++)
                {
                    using RichTextStyle style = new();
                    using TableCellDataColArray col_array = new();
                    for (int col = 0; col < col_count; col++)
                    {
                        DataRow? actual_row = table?.Rows[row];
                        DataColumn? actual_column = table?.Columns[col];
                        string cell_text = $"{actual_row?[col]}";
                        SetTableTextStyle(row, style);
                        using TableCellData cell_data = new(style, cell_text, new Image(), new RectF());
                        col_array.Add(cell_data);
                    }
                    cell_array.Add(col_array);
                }

                float page_width = page.GetWidth();
                float page_height = page.GetHeight();

                TableBorderInfo outside_border_left = new()
                {
                    line_width = 1,
                    table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid
                };
                TableBorderInfo outside_border_right = new()
                {
                    line_width = 1,
                    table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid
                };
                TableBorderInfo outside_border_top = new()
                {
                    line_width = 1,
                    table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid
                };
                TableBorderInfo outside_border_bottom = new()
                {
                    line_width = 1,
                    table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid
                };
                TableBorderInfo inside_border_row_info = new()
                {
                    line_width = 1,
                    table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid
                };
                TableBorderInfo inside_border_col_info = new()
                {
                    line_width = 1,
                    table_border_style = TableBorderInfo.TableBorderStyle.e_TableBorderStyleSolid
                };
                using RectF rect = new(10, 200, page_width - 10, page_height - 40);
                using TableData data = new(rect, row_per_page, col_count, outside_border_left, outside_border_right, outside_border_top, outside_border_bottom, inside_border_row_info, inside_border_col_info, new TableCellIndexArray(), new FloatArray(), new FloatArray());
                TableGenerator.AddTableToPage(page, data, cell_array);
            }
        }
        static void Main(string[] args) {
            Console.WriteLine($"Generating Table PDF...");
            try
            {
                // Add encoding required to parsed stringin excel docs
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                string input_file = data_path + "serious-injury-outcome-indicators-2000-2020.xlsx";

                using var stream = File.Open(input_file, FileMode.Open, FileAccess.Read);

                using var reader = ExcelReaderFactory.CreateReader(stream);
                row_count = reader.RowCount;
                col_count = reader.FieldCount > col_per_page ? col_per_page : reader.FieldCount;
                decimal pages = reader.RowCount / row_per_page;
                num_of_pages = (int) Math.Ceiling(pages);

                dataSet = reader.AsDataSet();

            } catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.GetType()} : {ex.Message}");
            }

            // Load variables from .env file to Environment
            Env.Load();

            Directory.CreateDirectory(output_path);

            string sn = Environment.GetEnvironmentVariable("FOXIT_SDK_SN") ?? "";
            string key = Environment.GetEnvironmentVariable("FOXIT_SDK_KEY") ?? "";

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