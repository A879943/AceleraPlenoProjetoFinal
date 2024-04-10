using AceleraPlenoProjetoFinal.Api.Data;
using AceleraPlenoProjetoFinal.Api.Models;
using Microsoft.AspNetCore.Mvc;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace AceleraPlenoProjetoFinal.Api.Controllers;

[ApiController]
[Route("api/v1/cargas")]
public class CargasController : ControllerBase
{
    private readonly DataContext _dataContext;
    private Excel.Application excelApp;

    public CargasController(DataContext dataContext)
    {
        _dataContext = dataContext;
        excelApp = new Excel.Application();
    }

    [HttpPost]
    [Route("import/transportadoravalores")]
    public IActionResult InserirTransportadoraValores()
    {
        var transportadoraList = new List<TransportadoraValoresModel>();

        Excel.Workbook excelWB = excelApp.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + @"\Resources\Transportadoras de Valores.xlsx");
        Excel._Worksheet excelWS = excelWB.Sheets[1];
        Excel.Range excelRange = excelWS.UsedRange;

        int rowCount = excelRange.Rows.Count;
        int columnCount = excelRange.Columns.Count;

        for (int i = 3; i <= rowCount; i++)
        {
            if (!string.IsNullOrEmpty(excelRange.Cells[i, 1].Value))
            {
                var transportadoraObj = new TransportadoraValoresModel()
                {
                    NumeroCnpj = excelRange.Cells[i, 1].Value2.ToString(),
                    DescricaoTransportadora = excelRange.Cells[i, 2].Value2.ToString(),
                    PA = excelRange.Cells[i, 3].Value2.ToString(),
                    DataHoraCarga = DateTime.Now
                };

                transportadoraList.Add(transportadoraObj);
            }
        }

        Marshal.ReleaseComObject(excelWS);
        Marshal.ReleaseComObject(excelRange);
        excelWB.Close();
        Marshal.ReleaseComObject(excelWB);
        excelApp.Quit();
        Marshal.ReleaseComObject(excelApp);

        _dataContext.AddRange(transportadoraList);
        _dataContext.SaveChanges();

        return Ok(transportadoraList);
    }

    [HttpPost]
    [Route("import/tipoterminal")]
    public IActionResult InserirTipoTerminal()
    {
        var tipoTerminalList = new List<TipoTerminalModel>();

        Excel.Workbook excelWB = excelApp.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + @"\Resources\TIPOS_TERMINAL.xlsx");
        Excel._Worksheet excelWS = excelWB.Sheets[1];
        Excel.Range excelRange = excelWS.UsedRange;

        int rowCount = excelRange.Rows.Count;
        int columnCount = excelRange.Columns.Count;

        for (int i = 2; i <= rowCount; i++)
        {
            if (!string.IsNullOrEmpty(excelRange.Cells[i, 3].Value))
            {
                var tipoTerminalObj = new TipoTerminalModel()
                {
                    CodigoTipoTerminal = int.Parse(excelRange.Cells[i, 2].Value2.ToString()),
                    DescricaoTipoTerminal = excelRange.Cells[i, 3].Value2.ToString(),
                    AcessoLiberado = true,
                    NumCheckAlteracao = 0,
                    PA = int.Parse(excelRange.Cells[i, 1].Value2.ToString()),
                    LimiteSuperior = int.Parse(excelRange.Cells[i, 4].Value2.ToString()),
                    LimiteInferior = int.Parse(excelRange.Cells[i, 5].Value2.ToString()),
                    DataHoraCarga = DateTime.Now
                };

                tipoTerminalList.Add(tipoTerminalObj);
            }
        }


        Marshal.ReleaseComObject(excelWS);
        Marshal.ReleaseComObject(excelRange);
        excelWB.Close();
        Marshal.ReleaseComObject(excelWB);
        excelApp.Quit();
        Marshal.ReleaseComObject(excelApp);

        _dataContext.AddRange(tipoTerminalList);
        _dataContext.SaveChanges();

        return Ok(tipoTerminalList);
    }

    [HttpPost]
    [Route("import/saldosiniciais")]
    public IActionResult InserirSaldosIniciais()
    {
        var operacaoList = new List<OperacaoModel>();

        Excel.Workbook excelWB = excelApp.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + @"\Resources\SALDOS INICIAIS 01.09.2022.xlsx");
        Excel._Worksheet excelWS = excelWB.Sheets[3];
        Console.WriteLine(excelWS.Name);
        Console.WriteLine(excelWB.Sheets.Count);
        Excel.Range excelRange = excelWS.UsedRange;

        int rowCount = excelRange.Rows.Count;
        int columnCount = excelRange.Columns.Count;

        for (int i = 3; i <= rowCount; i++)
        {
            if (!string.IsNullOrEmpty(excelRange.Cells[i, 1].Value))
            {
                var operacaoObj = new OperacaoModel()
                {
                    CodigoTipoTerminal = 2,
                    CodigoOperacao = excelRange.Cells[i, 1].Value2.ToString(),
                    DescricaoOperacao = excelRange.Cells[i, 2].Value2.ToString(),
                    CodigoHistorico = int.Parse(excelRange.Cells[i, 3].Value2.ToString()),
                    DescricaoHistorico = excelRange.Cells[i, 4].Value2.ToString(),
                    DataOperacao = DateTime.Parse(excelRange.Cells[i, 5].Value2.ToString()),
                    CodigoTerminal = excelRange.Cells[i, 6].Value2.ToString(),
                    Pa = int.Parse(excelRange.Cells[i, 7].Value2.ToString()),
                    CodigoAut = excelRange.Cells[i, 8].Value.ToString(),
                    Valor = decimal.Parse(excelRange.Cells[i, 9].Value2.ToString()),
                    DataHoraCarga = DateTime.Now
                };

                operacaoList.Add(operacaoObj);
            }
        }

        Marshal.ReleaseComObject(excelWS);
        Marshal.ReleaseComObject(excelRange);
        excelWB.Close();
        Marshal.ReleaseComObject(excelWB);
        excelApp.Quit();
        Marshal.ReleaseComObject(excelApp);

        _dataContext.AddRange(operacaoList);
        _dataContext.SaveChanges();

        return Ok(operacaoList);
    }
}