using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;


namespace sanpdfconvert_sdk_demo
{
    class Program
    {
        //pdf tool functions
        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_tool_init", CallingConvention = CallingConvention.Cdecl)]
        public static extern void pdf_tool_init(string server_path);
        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_tool_set_authorization_code", CallingConvention = CallingConvention.Cdecl)]
        public static extern void pdf_tool_set_authorization_code(string code);
        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_tool_get_authorization_request_code", CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr pdf_tool_get_authorization_request_code();

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_tool_check_pdf_password", CallingConvention = CallingConvention.Cdecl)]
        public static extern bool pdf_tool_check_pdf_password(string file_name, string password);

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_tool_get_pdf_page_count", CallingConvention = CallingConvention.Cdecl)]
        public static extern int pdf_tool_get_pdf_page_count(string file_name, string password);

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_tool_create_pdf_converter", CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr pdf_tool_create_pdf_converter();

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_tool_destroy_pdf_converter", CallingConvention = CallingConvention.Cdecl)]
        public static extern void pdf_tool_destroy_pdf_converter(IntPtr converter);


        //pdf converter functions
        [StructLayout(LayoutKind.Sequential)]
        public struct MERGE_INPUT_FILE
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string filename;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string password;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string convert_ranges;
        }

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_converter_merge_pdf", CallingConvention = CallingConvention.Cdecl)]
        public static extern void pdf_converter_merge_pdf(IntPtr converter, string output_file, [In]MERGE_INPUT_FILE[] merge_input_file, int input_file_count);

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_converter_split_pdf", CallingConvention = CallingConvention.Cdecl)]
        public static extern void pdf_converter_split_pdf(IntPtr converter, string input_file, string password, string ranges, string output_dir);

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_converter_pdf_to_office", CallingConvention = CallingConvention.Cdecl)]
        public static extern void pdf_converter_pdf_to_office(IntPtr converter, string output_file, string input_file, string ranges = "");

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_converter_office_to_pdf", CallingConvention = CallingConvention.Cdecl)]
        public static extern void pdf_converter_office_to_pdf(IntPtr converter, string output_file, string input_file, string ranges = "");

        public enum IMAGE_TYPE
        {
            JPG = 0,
            PNG = 1,
            BMP = 2,
        }

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_converter_pdf_to_image", CallingConvention = CallingConvention.Cdecl)]
        public static extern void pdf_converter_pdf_to_image(IntPtr converter, string output_path, IMAGE_TYPE image_type, string input_file, string ranges = "1-N", string password = "");


        [StructLayout(LayoutKind.Sequential)]
        public struct MERGE_PIC_TO_PDF_INPUT_FILE
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string filename;


        }

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_converter_image_to_pdf", CallingConvention = CallingConvention.Cdecl)]
        public static extern void pdf_converter_image_to_pdf(IntPtr converter, string output_file, [In] MERGE_PIC_TO_PDF_INPUT_FILE[] input_files, int input_file_count);

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_converter_clean_password", CallingConvention = CallingConvention.Cdecl)]
        public static extern bool pdf_converter_clean_password(IntPtr converter, string output_file, string input_file, string password);

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_converter_set_password", CallingConvention = CallingConvention.Cdecl)]
        public static extern bool pdf_converter_set_password(IntPtr converter, string output_file, string input_file, string password);

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_converter_get_progress", CallingConvention = CallingConvention.Cdecl)]
        public static extern int pdf_converter_get_progress(IntPtr converter);

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_converter_get_page_count", CallingConvention = CallingConvention.Cdecl)]
        public static extern int pdf_converter_get_page_count(IntPtr converter);

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_converter_get_convert_count", CallingConvention = CallingConvention.Cdecl)]
        public static extern int pdf_converter_get_convert_count(IntPtr converter);

        [DllImport("PDFTool.dll", CharSet = CharSet.Ansi, EntryPoint = "pdf_converter_get_error", CallingConvention = CallingConvention.Cdecl)]
        public static extern int pdf_converter_get_error(IntPtr converter);


        private static bool pdf_to_office(string in_file, string out_file, string convert_ranges)
        {
            IntPtr pdftooffice = pdf_tool_create_pdf_converter();

            if (pdftooffice == IntPtr.Zero)
                return false;

            pdf_converter_pdf_to_office(pdftooffice, out_file, in_file, convert_ranges);

            int convert_progress = 0;

            do
            {

                convert_progress = pdf_converter_get_progress(pdftooffice);

                Console.WriteLine("convert progress : {0:g}", convert_progress);

                Thread.Sleep(1000);

            } while (convert_progress >= 0 && convert_progress < 100);


            if (convert_progress >= 100)
            {
                //convert complete.
                Console.WriteLine("convert complete.");
            }
            else
            {
                Console.WriteLine("convert failed.errcode={0:g}", pdf_converter_get_error(pdftooffice));
            }

            pdf_tool_destroy_pdf_converter(pdftooffice);

            return (convert_progress >= 100);
        }

        private static bool office_to_pdf(string in_file, string out_file, string convert_ranges)
        {
            IntPtr pdftooffice = pdf_tool_create_pdf_converter();

            if (pdftooffice == IntPtr.Zero)
                return false;

            pdf_converter_office_to_pdf(pdftooffice, out_file, in_file, convert_ranges);

            int convert_progress = 0;

            do
            {

                convert_progress = pdf_converter_get_progress(pdftooffice);

                Console.WriteLine("convert progress : {0:g}", convert_progress);

                Thread.Sleep(1000);

            } while (convert_progress >= 0 && convert_progress < 100);


            if (convert_progress >= 100)
            {
                //convert complete.
                Console.WriteLine("convert complete.");
            }
            else
            {
                Console.WriteLine("convert failed.errcode={0:g}", pdf_converter_get_error(pdftooffice));
            }

            pdf_tool_destroy_pdf_converter(pdftooffice);

            return (convert_progress >= 100);
        }

        private static bool pic_to_pdf()
        {
            IntPtr pdftooffice = pdf_tool_create_pdf_converter();

            if (pdftooffice == IntPtr.Zero)
                return false;

            string current_dir = Directory.GetCurrentDirectory();
            string out_file = current_dir + "\\pic";

            MERGE_PIC_TO_PDF_INPUT_FILE[] input_files = new MERGE_PIC_TO_PDF_INPUT_FILE[2];

            string in_file1 = out_file + "\\1.jpg";
            input_files[0].filename = in_file1;
            string in_file2 = out_file + "\\2.jpg";
            input_files[1].filename = in_file2;

            out_file = current_dir + "\\pictopdf.pdf";

            pdf_converter_image_to_pdf(pdftooffice, out_file, input_files, 2);

            int convert_progress = 0;

            do
            {

                convert_progress = pdf_converter_get_progress(pdftooffice);

                Console.WriteLine("convert progress : {0:g}", convert_progress);

                Thread.Sleep(1000);

            } while (convert_progress >= 0 && convert_progress < 100);


            if (convert_progress >= 100)
            {
                //convert complete.
                Console.WriteLine("convert complete.");
            }
            else
            {
                Console.WriteLine("convert failed.errcode={0:g}", pdf_converter_get_error(pdftooffice));
            }

            pdf_tool_destroy_pdf_converter(pdftooffice);

            return (convert_progress >= 100);
        }
        static void Main(string[] args)
        {


            //get request code
            IntPtr pReqCode = pdf_tool_get_authorization_request_code();
            string request_code = Marshal.PtrToStringAnsi(pReqCode);
            Console.WriteLine(request_code);
            //TODO::

            //sanpdf convert sdk api 
            //1、init sanpdf sdk model
            string auth_code = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJzdWIiOiIyMDE5MTIwMzAwMDEiLCJhZGRDbGFpbXMiOnsicGsiOiIyIiwic2FuIjoiWlNRUUdtcEVYNWV4ZmlSZ2dUMFgxdk1SZ1lZbkxHZFRRMDhlemdGWWo5ZTZhQ0RwVWpSTkF6WGIwU3BPY2xKbGthRHN4dkpRYUJJdyt2SzlsSnp5MHdLSER6eDR6SGw0WTVUVTUyWVdvb3NZdGZwMXA2eGVKc0VNMGRSRnlqWFU4NjFhQThXdlhZNlwvXC9LUDNTM0tGUXJONUxXcGw5TkRmUnQ0T3dsR0FiVHZpTHRZdDJLMmF4RUVcL1VnRHROZWVyaXRvNWhidTlicVNGbVIzTlgxcHdHQ1Q1bnppT1RGXC9WYjc1TjBPXC9PUTlWVE9VcmIyTlZTd1c1cW9sbGJcL1JkMU9CM3lDK2hxa0xNVGh1NjdiRXBqc1wvVTlpOXJBSmxrNDNiZ0x4OHU5WXE3ZU1GUXBuMVg3OVRCXC90ZTdKZHZvRjJTc2hGNHNRbzIwM1RWU0tqdkoyZGc9PSJ9LCJpc3MiOiJodHRwczpcL1wvc2FuYXV0aC5wZGZrei5jb21cL2FwaVwvR2V0QXV0aG9yaXphdGlvbkNvZGUiLCJpYXQiOjE1NzU1MTMzOTQsImV4cCI6NDcyOTExMzM5NCwibmJmIjoxNTc1NTEzMzk0LCJqdGkiOiI3UWM2ZmZCU1dKZGZkTlkwIn0.VfYEDrSg6k4HUYNZIW35XW8ZCHqRDMu3HXcuUk-Mzhc3GspGzppQrsk5bzSZ3tBX8D5zZm_6PR2jtuEGxJ3uDTUkUgkDEq0SJvsB4kOcrwsHuIoVXTDG-4aoMAMhnrDTDTnewlCrj7cbIL0jmaVZq--5hqfYoSoc5PYmmdIoo9oI6vPCWMj9_wXZI30KY1TlM6CDvbWWjxE77m15NpQSJ4G9De-O4mqkg5yjWKNBfmZUBVjd2-wASZgxqGoH4JMk5vz-uQJB6uVPMYbCI9HlSIXh0xFTryvdCwFjdhg1pGsAh29TE0CJ3f6PHu30qJh4fFO0IniSkPr_79F2tYNabA";
            pdf_tool_init("https://convert.sanpdf.com");
            pdf_tool_set_authorization_code(auth_code);

            //2、PDF TO WORD demo
            string current_dir = Directory.GetCurrentDirectory();
            string in_file = current_dir + "\\test.pdf";
            string out_file = current_dir + "\\test.xlsx";



            //check pdf file is exists password or not
            //if pdf_tool_check_pdf_password api return true the pdf file not exists password
            //else return false exists password.
            bool check = pdf_tool_check_pdf_password(in_file, "");

            if (!check)
            {
                Console.WriteLine("input pdf file exists password.\n");
                return;
            }

            //pdf to word or pdf to excel or pdf to ppt 

            if (pdf_to_office(in_file, out_file, "1-3"))
            {
                Console.WriteLine("convert ok.");
            }

            string pdf_out_file = current_dir + "\\test.pdf";
            string doc_in_file = current_dir + "\\test.xlsx";

            //word to pdf or excel to pdf or ppt to pdf
            office_to_pdf(doc_in_file, pdf_out_file, "1-N");

            //images(jpg,png,bmp,eg.) to pdf 
            pic_to_pdf();

        }
    }
}

