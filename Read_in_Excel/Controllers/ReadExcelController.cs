using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Windows.Forms;
using Read_in_Excel.Models;
using System.IO;

namespace Read_in_Excel.Controllers
{
    public class ReadExcelController : Controller
    {
        ReadExcelModel ob = new ReadExcelModel();
        //
        // GET: /ReadExcel/
        //
        public ActionResult Index()
        {
            // OpenFileDialog opendialog = new OpenFileDialog();
            //opendialog.ShowDialog();        
            //if (opendialog.ShowDialog() == DialogResult.OK)
            //{
            //    //ob.Get_Values_in_Excel(opendialog.FileName);
            // //   ob.Get_Values_in_Excel(@"C:\Users\$ARVARBEK\Desktop\Shablon.xlsx");
            //    MessageBox.Show("satr=" + ob.satr);
            //    MessageBox.Show("ustun=" + ob.ustun);
            //}
            return View();
        }
        [HttpPost]
        public ActionResult Index (HttpPostedFileBase file)
        {
            //var files = Request.Files[0];
            //MessageBox.Show("Ura1");
            //MessageBox.Show(files.ContentLength.ToString());
            //if(file!=null && file.ContentLength>0)
            //{
            //    MessageBox.Show("Ura2");
            //    MessageBox.Show(Path.GetFileName(file.FileName));
            //    var filename = Path.GetFileName(file.FileName);
            //    var path = Path.Combine(Server.MapPath("~/Images/"), filename);
            //    file.SaveAs(path);
            //}
            return RedirectToAction("Index");
        }
        [HttpPost]
        public ActionResult Index1()
        {
            var files = Request.Files[0];
            
            MessageBox.Show("Ura1");            
            if (files != null && files.ContentLength > 0)
            {
                MessageBox.Show("Ura2");
                ob.Read();
                var filename = Path.GetFileName(files.FileName);
                var path = Path.Combine(Server.MapPath("~/Images/"), filename);
                //files.SaveAs(path);
            }
            return RedirectToAction("Index");
        }

        
    }
}
