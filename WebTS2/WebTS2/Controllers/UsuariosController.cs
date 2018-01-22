using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using WebTS2.Models;
using WebTS2.Helper;
using OfficeOpenXml;
using iTextSharp.text;
using System.IO;
using iTextSharp.text.pdf;

namespace WebTS2.Controllers
{

	public class UsuarioIndexViewModel
    {
		public List<Usuario> Items { get; set; }
        public Pager Pager { get; set; }
    }

    public class UsuariosController : Controller
    {
        private EntitiesTierraSanta db = new EntitiesTierraSanta();

		
        // GET: Usuarios
        public ActionResult Index(int? page, String Search)
        {
		//
			var viewModel = new UsuarioIndexViewModel();            

            if (Search == null || Search.Equals(""))
            {
				var pager = new Pager(db.Usuario.Count(), page);
                viewModel.Items = db.Usuario
                        .OrderBy(c => c.Nombre)
                        .Skip((pager.CurrentPage - 1) * pager.PageSize)
                        .Take(pager.PageSize).ToList();
                viewModel.Pager = pager;
            }
            else
            {
				var pager = new Pager(db.Usuario.Where(c => c.Nombre.Contains(Search)).Count(), page);
                viewModel.Items = db.Usuario.Where(c => c.Nombre.Contains(Search))
                        .OrderBy(c => c.Nombre)
                        .Skip((pager.CurrentPage - 1) * pager.PageSize)
                        .Take(pager.PageSize).ToList();
				viewModel.Pager = pager;
				@ViewBag.Search = Search;
            }
            return View(viewModel);
        }



        // GET: Usuarios/Details/5
        public ActionResult Details(Guid? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Usuario usuario = db.Usuario.Find(id);
            if (usuario == null)
            {
                return HttpNotFound();
            }
            return View(usuario);
        }

        // GET: Usuarios/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Usuarios/Create
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "UsuarioId,Nombre,Login,Clave,Email,Estado,EsAdministrador,Creado,Modificado")] Usuario usuario)
        {
            if (ModelState.IsValid)
            {
                usuario.UsuarioId = Guid.NewGuid();
                usuario.Creado = DateTime.Now;
                usuario.Modificado = DateTime.Now;
                db.Usuario.Add(usuario);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(usuario);
        }

        // GET: Usuarios/Edit/5
        public ActionResult Edit(Guid? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Usuario usuario = db.Usuario.Find(id);
            if (usuario == null)
            {
                return HttpNotFound();
            }
            return View(usuario);
        }

        // POST: Usuarios/Edit/5
        // Para protegerse de ataques de publicación excesiva, habilite las propiedades específicas a las que desea enlazarse. Para obtener 
        // más información vea http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "UsuarioId,Nombre,Login,Clave,Email,Estado,EsAdministrador,Creado,Modificado")] Usuario usuario)
        {
            if (ModelState.IsValid)
            {
                db.Entry(usuario).State = EntityState.Modified;
				usuario.Modificado = DateTime.Now;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(usuario);
        }

        // GET: Usuarios/Delete/5
        public ActionResult Delete(Guid? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Usuario usuario = db.Usuario.Find(id);
            if (usuario == null)
            {
                return HttpNotFound();
            }
            return View(usuario);
        }

        // POST: Usuarios/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(Guid id)
        {
            Usuario usuario = db.Usuario.Find(id);
            db.Usuario.Remove(usuario);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

		public ActionResult ReportExcel()
        {
            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("Report");
                ExcelWorksheet ws = package.Workbook.Worksheets[1];
                ws.Name = "Report";
                ws.Cells.Style.Font.Size = 11;
                ws.Cells.Style.Font.Name = "Calibri";
				//
                List<Usuario> list = db.Usuario.ToList();
                int pos = 4;
								ws.Cells[pos, 3].Value = "Nombre";
									ws.Cells[pos, 4].Value = "Login";
									ws.Cells[pos, 5].Value = "Clave";
									ws.Cells[pos, 6].Value = "Email";
									ws.Cells[pos, 7].Value = "Estado";
									ws.Cells[pos, 8].Value = "EsAdministrador";
									ws.Cells[pos, 9].Value = "Creado";
									ws.Cells[pos, 10].Value = "Modificado";
					
                foreach (var item in list)
                {
                    pos++;
								ws.Cells[pos, 3].Value = item.Nombre == null ? "" : item.Nombre.ToString();				
									ws.Cells[pos, 4].Value = item.Login == null ? "" : item.Login.ToString();				
									ws.Cells[pos, 5].Value = item.Clave == null ? "" : item.Clave.ToString();				
									ws.Cells[pos, 6].Value = item.Email == null ? "" : item.Email.ToString();				
									ws.Cells[pos, 7].Value = item.Estado == null ? "" : item.Estado.ToString();				
									ws.Cells[pos, 8].Value = item.EsAdministrador == null ? "" : item.EsAdministrador.ToString();				
									ws.Cells[pos, 9].Value = item.Creado == null ? "" : item.Creado.ToString();				
									ws.Cells[pos, 10].Value = item.Modificado == null ? "" : item.Modificado.ToString();				
					                }
				ws.Cells["B3:F" + pos].AutoFitColumns();


                Response.Clear();
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", "attachment; filename=Report.xlsx");
                Response.BinaryWrite(package.GetAsByteArray());
                Response.Flush();
                Response.Close();
            }

            return null;
        }

		public ActionResult ReportPDF()
        {
            var document = new Document(PageSize.A4, 50, 50, 25, 25);
            var output = new MemoryStream();
            var writer = PdfWriter.GetInstance(document, output);
            document.Open();

			
            var table = new PdfPTable(8);

            var boldTableFont = FontFactory.GetFont("Arial", 10, Font.BOLD);
            var bodyFont = FontFactory.GetFont("Arial", 10, Font.NORMAL);

							table.AddCell(new Phrase("Nombre", boldTableFont));
									table.AddCell(new Phrase("Login", boldTableFont));
									table.AddCell(new Phrase("Clave", boldTableFont));
									table.AddCell(new Phrase("Email", boldTableFont));
									table.AddCell(new Phrase("Estado", boldTableFont));
									table.AddCell(new Phrase("EsAdministrador", boldTableFont));
									table.AddCell(new Phrase("Creado", boldTableFont));
									table.AddCell(new Phrase("Modificado", boldTableFont));
									              
//
            List<Usuario> list = db.Usuario.ToList();

			foreach (var item in list)
                {
                    
								table.AddCell(new Phrase(item.Nombre == null ? "" : item.Nombre.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.Login == null ? "" : item.Login.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.Clave == null ? "" : item.Clave.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.Email == null ? "" : item.Email.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.Estado == null ? "" : item.Estado.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.EsAdministrador == null ? "" : item.EsAdministrador.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.Creado == null ? "" : item.Creado.ToString(), bodyFont));			
									table.AddCell(new Phrase(item.Modificado == null ? "" : item.Modificado.ToString(), bodyFont));			
									}

            

            document.Add(table);
            document.Close();

            Response.Clear();
            Response.ContentType = "application/pdf";
            Response.AddHeader("Content-Disposition", "attachment; filename=Report.pdf");
            Response.BinaryWrite(output.ToArray());
            Response.Flush();
            Response.Close();

            return null;
        }
    }
}
