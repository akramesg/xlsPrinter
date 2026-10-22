using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;

namespace XlsPrinter.Core
{
    /// <summary>
    /// Generates authentic GDI Enhanced Metafile (.EMF) and NT EMF Spool (.SPL) test streams
    /// containing Arabic & English tabular invoice/accounting data, simulating a classic Win32
    /// accounting/ERP system printing a report.
    /// </summary>
    public static class SamplePrintGenerator
    {
        public static byte[] CreateSampleArabicAccountingEmf()
        {
            // Create an in-memory EMF metafile using standard Win32 GDI Graphics
            using (Bitmap refBmp = new Bitmap(1, 1))
            using (Graphics refG = Graphics.FromImage(refBmp))
            {
                IntPtr hdc = refG.GetHdc();
                try
                {
                    MemoryStream emfStream = new MemoryStream();
                    // A4 size in 0.01mm units: 21000 x 29700
                    RectangleF frameRect = new RectangleF(0, 0, 21000, 29700);
                    using (Metafile mf = new Metafile(emfStream, hdc, frameRect, MetafileFrameUnit.GdiCompatible, EmfType.EmfOnly))
                    using (Graphics g = Graphics.FromImage(mf))
                    {
                        g.PageUnit = GraphicsUnit.Millimeter;

                        using (Font titleFont = new Font("Tahoma", 13f, FontStyle.Bold))
                        using (Font headerFont = new Font("Tahoma", 10f, FontStyle.Bold))
                        using (Font bodyFont = new Font("Tahoma", 9.5f, FontStyle.Regular))
                        using (Pen linePen = new Pen(Color.DarkSlateGray, 0.3f))
                        using (SolidBrush blackBrush = new SolidBrush(Color.Black))
                        using (SolidBrush headerBrush = new SolidBrush(Color.Navy))
                        {
                            // Report Header (Bilingual Arabic / English Classic ERP Style)
                            g.DrawString("Tripack", titleFont, headerBrush, 15f, 14f);
                            g.DrawString("كشف حساب المبيعات والفواتير - لشهر أكتوبر 2026", titleFont, headerBrush, 105f, 14f);

                            g.DrawString("Report No: INV-2026-108", bodyFont, blackBrush, 15f, 24f);
                            g.DrawString("Date: 2026-10-20", bodyFont, blackBrush, 75f, 24f);
                            g.DrawString("الفرع الرئيسي - الرياض", bodyFont, blackBrush, 135f, 24f);

                            g.DrawLine(linePen, 12f, 31f, 196f, 31f);

                            // Table Column Positions (in millimeters across A4 width)
                            float colCodeX = 15f;
                            float colDateX = 38f;
                            float colCustomerX = 68f;
                            float colItemX = 118f;
                            float colQtyX = 158f;
                            float colTotalX = 176f;

                            // Table Header Row (Y = 35mm)
                            float y = 35f;
                            g.DrawString("رقم الفاتورة", headerFont, headerBrush, colCodeX, y);
                            g.DrawString("التاريخ", headerFont, headerBrush, colDateX, y);
                            g.DrawString("اسم العميل", headerFont, headerBrush, colCustomerX, y);
                            g.DrawString("البيان / الصنف", headerFont, headerBrush, colItemX, y);
                            g.DrawString("الكمية", headerFont, headerBrush, colQtyX, y);
                            g.DrawString("الإجمالي", headerFont, headerBrush, colTotalX, y);

                            g.DrawLine(linePen, 12f, 42f, 196f, 42f);

                            // Sample Invoice Rows (Arabic + Numbers + English codes)
                            string[,] rows = new string[,]
                            {
                                { "10041", "2026-10-01", "شركة الأمل للتجارة", "أجهزة حاسب مكتبي", "12", "28500.00" },
                                { "10042", "2026-10-03", "مؤسسة النهضة الحديثة", "طابعات ليزر ملونة", "5", "7250.50" },
                                { "10043", "2026-10-05", "مستشفى الشفاء التخصصي", "خوادم شبكات رئيسية", "2", "44000.00" },
                                { "10044", "2026-10-08", "شركة الخليج للمقاولات", "شاشات عرض تفاعلية", "8", "15600.75" },
                                { "10045", "2026-10-11", "مكتبة المعرفة الكبرى", "برمجيات محاسبة ومخازن", "15", "22500.00" },
                                { "10046", "2026-10-14", "مجموعة الفيصل الصناعية", "وحدات تخزين احتياطي", "10", "13400.00" },
                                { "10047", "2026-10-17", "شركة الشرق الأوسط", "عقود صيانة سنوية", "1", "18000.00" },
                                { "10048", "2026-10-19", "مؤسسة البناء المتطور", "كابلات ومعدات ربط", "40", "6800.25" },
                                { "10049", "2026-10-20", "مركز الإبداع التقني", "أجهزة قارئ باركود", "25", "9375.00" }
                            };

                            y = 46f;
                            for (int r = 0; r < rows.GetLength(0); r++)
                            {
                                g.DrawString(rows[r, 0], bodyFont, blackBrush, colCodeX, y);
                                g.DrawString(rows[r, 1], bodyFont, blackBrush, colDateX, y);
                                g.DrawString(rows[r, 2], bodyFont, blackBrush, colCustomerX, y);
                                g.DrawString(rows[r, 3], bodyFont, blackBrush, colItemX, y);
                                g.DrawString(rows[r, 4], bodyFont, blackBrush, colQtyX, y);
                                g.DrawString(rows[r, 5], bodyFont, blackBrush, colTotalX, y);
                                y += 8.0f;
                            }

                            g.DrawLine(linePen, 12f, y + 2f, 196f, y + 2f);
                            y += 5f;

                            // Summary Row
                            g.DrawString("TOTAL / المجموع الكلي", headerFont, headerBrush, colCustomerX, y);
                            g.DrawString("118", headerFont, headerBrush, colQtyX, y);
                            g.DrawString("165426.50", headerFont, headerBrush, colTotalX, y);

                            // Footer
                            y += 18f;
                            g.DrawString("اعتمد من قبل الإدارة المالية - توقيع المحاسب", bodyFont, blackBrush, 15f, y);
                            g.DrawString("Page 1 of 1", bodyFont, blackBrush, 165f, y);
                        }
                    }
                    return emfStream.ToArray();
                }
                finally
                {
                    refG.ReleaseHdc(hdc);
                }
            }
        }
    }
}
