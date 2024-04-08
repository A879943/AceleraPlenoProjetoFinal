using AceleraPlenoProjetoFinal.Api.Data;
using AceleraPlenoProjetoFinal.Api.Models;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;

namespace AceleraPlenoProjetoFinal.Api.Controllers;

[ApiController]
[Route("api/v1/cargas")]
public class CargasController : ControllerBase
{
    private readonly DataContext _dataContext;

    public CargasController(DataContext dataContext)
    {
        _dataContext = dataContext;
    }

    [HttpPost]
    [Route("import/transportadoravalores")]
    public IActionResult InserirTransportadoraValores()
    {
        List<TransportadoraValoresModel> transportadoraList = new List<TransportadoraValoresModel>();

        string excelFilePath = System.IO.Directory.GetCurrentDirectory() + "\\Resources\\Transportadoras de Valores.xlsx";

        var workbook = new XLWorkbook(excelFilePath);

        var planilha = workbook.Worksheets.First(w => w.Name == "Planilha1");

        var totalLinhas = planilha.Rows().Count();

        // primeira linha é o cabecalho
        for (int l = 3; l <= totalLinhas; l++)
        {
            if (!string.IsNullOrEmpty(planilha.Cell($"B{l}").Value.ToString()))
            {
                TransportadoraValoresModel transportadoraObj = new TransportadoraValoresModel()
                {
                    NumeroCnpj = planilha.Cell($"B{l}").Value.ToString(),
                    DescricaoTransportadora = planilha.Cell($"C{l}").Value.ToString(),
                    PA = planilha.Cell($"D{l}").Value.ToString(),
                    DataHoraCarga = DateTime.Now
                };

                transportadoraList.Add(transportadoraObj);
            }
        }

        _dataContext.AddRange(transportadoraList);
        _dataContext.SaveChanges();

        return Ok(transportadoraList);
    }

    [HttpPost]
    [Route("import/tipoterminal")]
    public IActionResult InserirTipoTerminal()
    {
        List<TipoTerminalModel> tipoTerminalList = new List<TipoTerminalModel>();

        string excelFilePath = System.IO.Directory.GetCurrentDirectory() + "\\Resources\\TIPOS_TERMINAL.xlsx";

        var workbook = new XLWorkbook(excelFilePath);

        var planilha = workbook.Worksheets.First(w => w.Name == "Planilha1");

        var totalLinhas = planilha.Rows().Count();

        // primeira linha é o cabecalho
        for (int l = 2; l <= totalLinhas; l++)
        {
            if (!string.IsNullOrEmpty(planilha.Cell($"A{l}").Value.ToString()))
            {
                TipoTerminalModel tipoTerminalObj = new TipoTerminalModel()
                {
                    CodigoTipoTerminal = int.Parse(planilha.Cell($"B{l}").Value.ToString()),
                    DescricaoTipoTerminal = planilha.Cell($"C{l}").Value.ToString(),
                    AcessoLiberado = true,
                    NumCheckAlteracao = 0,
                    PA = int.Parse(planilha.Cell($"A{l}").Value.ToString()),
                    LimiteSuperior = int.Parse(planilha.Cell($"D{l}").Value.ToString()),
                    LimiteInferior = int.Parse(planilha.Cell($"E{l}").Value.ToString()),
                    DataHoraCarga = DateTime.Now
                };

                tipoTerminalList.Add(tipoTerminalObj);
            }
        }

        _dataContext.AddRange(tipoTerminalList);
        _dataContext.SaveChanges();

        return Ok(tipoTerminalList);
    }
}