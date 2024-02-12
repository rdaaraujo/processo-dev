using Microsoft.AspNetCore.Mvc;

namespace Application.Controllers
{
    public class CepsController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}
