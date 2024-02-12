using Microsoft.AspNetCore.Mvc;
using _1___Application.Models;
using _4___Domain;
using Services;
using System.Diagnostics;
using _1___Application.Controllers;

namespace Application.Controllers
{
    public class DadosController : Controller
    {
        private readonly ILogger<DadosController> _logger;

        public GetData getData = new GetData();

        public ReturnApiCEPModel Return { get; private set; }

        public DadosController(ILogger<DadosController> logger)
        {
            _logger = logger;
        }

        public IActionResult DadosBusca()
            {
                string cepDigitado = Request.Form["cepValor"];

                if (cepDigitado != null)
                {
                    var result = getData.ApiCEPGet(cepDigitado);
                    ViewBag.cep = result.cep;
                    ViewBag.logradouro = result.logradouro;
                    ViewBag.complemento = result.complemento;
                    ViewBag.bairro = result.bairro;
                    ViewBag.localidade = result.localidade;
                    ViewBag.uf = result.uf;
                    ViewBag.ibge = result.ibge;
                    ViewBag.gia = result.gia;
                    ViewBag.ddd = result.ddd;
                    ViewBag.siafi = result.siafi;
                    return View();
                }
                return View();
            }
    }
}