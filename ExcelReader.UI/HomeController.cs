using Microsoft.AspNetCore.Mvc;

// For more information on enabling MVC for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace ExcelReader.UI {
    public class HomeController : Controller {
        // GET: /<controller>/
        public IActionResult Index() {

            var reader = new Reader();

            // ViewBag.Aniversariantes, será usada na tela "Index.cshtml
            ViewBag.Aniversariantes = reader.ObterAniversariantes();

            return View();
        }
    }
}
