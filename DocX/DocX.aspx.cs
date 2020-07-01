using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
namespace DocX
{
    public partial class DocX : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            var Name = "Thapakorn";
            var Age = "40";
            string savePath = Server.MapPath("~/PersonInfo.doc");
            string templatePath = Server.MapPath("~/wordTemplate2.doc");
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();
            doc = app.Documents.Open(templatePath);
            doc.Activate();
            if (doc.Bookmarks.Exists("Name"))
            {
                doc.Bookmarks["Name"].Range.Text = Name;
            }
            if (doc.Bookmarks.Exists("Age"))
            {
                doc.Bookmarks["Age"].Range.Text = Age;
            }
            if (doc.Bookmarks.Exists("Time"))
            {
                doc.Bookmarks["Time"].Range.Text = DateTime.Now.ToString("yyyy-MM-dd");
            }
            

            doc.SaveAs2(savePath);
            doc.Close();
            app.Application.Quit();

            DownloadFile(savePath);

            //Process.Start("WINWORD.EXE", "\"" + savePath + "\"");
            
            //Response.Write("Success");
        }
        private void DownloadFile(string file)
        {
            var fi = new FileInfo(file);
            Response.Clear();
            Response.ContentType = "application/octet-stream";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + fi.Name);
            Response.WriteFile(file);
            Response.End();
        }
    }
}