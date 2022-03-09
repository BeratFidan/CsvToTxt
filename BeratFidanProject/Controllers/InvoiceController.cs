using BeratFidanProject.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace BeratFidanProject.Controllers
{
    public class InvoiceController : Controller
    {
        // GET: Invoice
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelfile)
        {

            if (excelfile == null || excelfile.ContentLength == 0)
            {
                ViewBag.Error1 = "Lutfen bir dosya secin.";
                return View("Index");
            }
            else
            {
                if (excelfile.FileName.EndsWith("csv"))
                {

                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open("C:/Users/Berat/Desktop/Yeni/" + excelfile.FileName);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;

                    List<Invoice> invoices = new List<Invoice>();

                    for (int i = 1; i <= range.Rows.Count; i++)
                    {
                        Invoice inv = new Invoice();

                        inv.FaturaNo = ((Excel.Range)range.Cells[i, 1]).Text;
                        invoices.Add(inv);

                    }

                    if (invoices != null)
                    {
                        ViewBag.InvList = invoices;

                        return View("View");

                    }
                    else
                    {
                        ViewBag.Error = "Bir seyler yanlis gitti.";
                        return View("Index");
                    }

                }
                else
                {
                    ViewBag.Error = "Bir seyler yanlis gitti.";
                    return View("Index");
                }
            }

        }


        [HttpPost]
        public ActionResult TxtOlustur(HttpPostedFileBase excelfile)
        {
            if (excelfile == null || excelfile.ContentLength == 0)
            {
                ViewBag.Error2 = "Lutfen bir dosya secin.";
                return View("Index");
            }

            else
            {
                if (excelfile.FileName.EndsWith("csv"))
                {

                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open("C:/Users/Berat/Desktop/Yeni/" + excelfile.FileName);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;

                    List<Invoice> invoices = new List<Invoice>();

                    for (int i = 1; i <= range.Rows.Count; i++)
                    {
                        Invoice inv = new Invoice();

                        inv.FaturaNo = ((Excel.Range)range.Cells[i, 1]).Text;
                        string[] value = inv.FaturaNo.Split(';');
                        inv.FaturaNo = value[0];
                        inv.Ad = value[1];
                        inv.Soyad = value[2];
                        inv.Tutar = value[3];

                        invoices.Add(inv);
                       

                        return View(invoices);
                    }



                    if (invoices != null)
                    {

                        return View(invoices);

                    }
                    else
                    {
                        ViewBag.Error3 = "Bir seyler yanlis gitti.";
                        return View("Index");
                    }

                }
                else
                {
                    ViewBag.Error4 = "Bir seyler yanlis gitti.";
                    return View("Index");
                }
            }
        }


    }
}