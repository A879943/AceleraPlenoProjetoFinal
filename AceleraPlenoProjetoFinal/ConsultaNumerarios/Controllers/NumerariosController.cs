using ConsultaNumerarios.Interfaces;
using ConsultaNumerarios.Models;
using ConsultaNumerarios.Services;
using Microsoft.AspNetCore.Mvc;

namespace ConsultaNumerarios.Controllers
{
    [ApiController]
    [Route("api/v0.1/[controller]")]
    public class NumerariosController : ControllerBase
    {

        private readonly IUnidadeInstituicao _pa;
        private readonly ITerminalService _terminalService;
        private readonly ILogger<NumerariosController> _logger;

        public NumerariosController(ILogger<NumerariosController> logger, IUnidadeInstituicao pa, ITerminalService terminalService)
        {
            _logger = logger;
            _pa = pa;
            _terminalService = terminalService;
        }
        
        //TODO: Testar
        [HttpGet("ListaDePas")]
        public IActionResult GetPasAtivos()
        {
            List<UnidadeInstituicaoModel> unidades = _pa.GetUnidadeInstituicao();
            if (unidades.Count == 0)
                return NotFound("Erro na busca de PAs");
            
            return Ok(unidades);
        }

        [HttpGet("SaldoDeTerminais{idPa}")]
        public IActionResult GetSaldosDosTerminais([FromRoute] string idPa)
        {
            if (int.TryParse(idPa, out int idPaParsed))
                return BadRequest(idPa + " não é um id válido");

            return Ok(_terminalService.CalculaSaldosDeTerminais(idPaParsed));

        }
    }
}
