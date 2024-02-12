using _2___Services;
using _4___Domain;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace Application.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        Crud crud = new Crud();

        // GET api/<ValuesController>/5
        [HttpGet]
        public List<DadosCli> Get()
        {
            var value = crud.getData();
            return value;
        }

        // POST api/<ValuesController>
        [HttpPost]
        public DadosCli Post([FromBody] DadosCli value)
        {
            crud.insertData(value);
            return value;
        }

        // PUT api/<ValuesController>/5
        [HttpPut]
        public string Put(int Id)
        {
            crud.updateData(Id);
            return "Situacação atualizada com sucesso!";
        }

        // DELETE api/<ValuesController>/5
        [HttpDelete]
        public string Delete(string Id)
        {
            crud.deleteData(Id);
            return "Clientes deletados com sucesso!";
        }

        //// GET api/<ValuesController>/5
        //[HttpGet]
        //public List<DadosCli> GetAll()
        //{
        //    var value = crud.getAllData();
        //    return value;
        //}
    }
}
