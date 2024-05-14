using AceleraPlenoProjetoFinal.Api.Data;
using AceleraPlenoProjetoFinal.Api.Models;
using AceleraPlenoProjetoFinal.Api.Validations;
using Microsoft.AspNetCore.Mvc;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace AceleraPlenoProjetoFinal.Api.Controllers;

[ApiController]
[Route("api/v1/CargasIniciais")]
public class CargasIniciaisController : ControllerBase
{
    private readonly DataContext _dataContext;
    private Excel.Application excelApp;
    private readonly ValidarDadosCarga _validar;

    public CargasIniciaisController(DataContext dataContext)
    {
        _dataContext = dataContext;
        excelApp = new Excel.Application();
        _validar = new ValidarDadosCarga();
    }

    [HttpPost]
    [Route("UnidadeInstituicao")]
    public IActionResult InserirUnidadeInstituicao()
    {
        var unidadeInstList = new List<UnidadeInstituicaoModel>();

        Excel.Workbook excelWB = excelApp.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + @"\Resources\UNIDADEINSTITUICAO.xlsx");
        Excel._Worksheet excelWS = excelWB.Sheets[1];
        Excel.Range excelRange = excelWS.UsedRange;

        int rowCount = excelRange.Rows.Count;
        int columnCount = excelRange.Columns.Count;

        for (int i = 3; i <= rowCount; i++)
        {
            if (!string.IsNullOrEmpty(excelRange.Cells[i, 4].Value))
            {
                var unidadeInstObj = new UnidadeInstituicaoModel();

                unidadeInstObj.IdUnidadeInstituicao = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 2].Value2));
                unidadeInstObj.IdInstituicao = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 1].Value2));
                unidadeInstObj.Cnpj = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 3].Value2));
                unidadeInstObj.NomeUnidade = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 4].Value2));
                unidadeInstObj.SiglaUnidade = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 5].Value2));
                unidadeInstObj.DataCadastramento = _validar.VerificarDadosData(Convert.ToString(excelRange.Cells[i, 6].Value2));
                unidadeInstObj.CodTipoInstituicao = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 7].Value2));
                unidadeInstObj.CodTipoUnidade = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 8].Value2));
                unidadeInstObj.DataInicioSicoob = _validar.VerificarDadosData(Convert.ToString(excelRange.Cells[i, 9].Value2));
                unidadeInstObj.DataFimSicoob = _validar.VerificarDadosData(Convert.ToString(excelRange.Cells[i, 10].Value2));
                unidadeInstObj.NumCheckAlteracao = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 12].Value2));
                unidadeInstObj.IdUnidadeInstResp = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 13].Value2));
                unidadeInstObj.CodSituacaoUnid = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 14].Value2));
                unidadeInstObj.NumSirc = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 15].Value2));
                unidadeInstObj.DescricaoEndInternet = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 16].Value2));
                unidadeInstObj.DataInicioUtilizacaoMarcaSicoob = _validar.VerificarDadosData(Convert.ToString(excelRange.Cells[i, 17].Value2));
                unidadeInstObj.BolAtentimentoPublicoExterno = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 18].Value2));
                unidadeInstObj.NumInscricaoMunicipal = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 19].Value2));
                unidadeInstObj.NumNire = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 20].Value2));
                unidadeInstObj.IdInstituicaoIncorporadora = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 21].Value2));
                unidadeInstObj.DataIncorporacao = _validar.VerificarDadosData(Convert.ToString(excelRange.Cells[i, 22].Value2));
                unidadeInstObj.BolUtilizaCompartilhamento = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 23].Value2));
                unidadeInstObj.DataInicioFuncionamento = _validar.VerificarDadosData(Convert.ToString(excelRange.Cells[i, 24].Value2));
                unidadeInstObj.DataFimFuncionamento = _validar.VerificarDadosData(Convert.ToString(excelRange.Cells[i, 25].Value2));
                unidadeInstObj.BolUtilizaSisbr = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 26].Value2));
                unidadeInstObj.DataInicioUtilizaSisbr = _validar.VerificarDadosData(Convert.ToString(excelRange.Cells[i, 27].Value2));
                unidadeInstObj.DataFimUtilizaSisbr = _validar.VerificarDadosData(Convert.ToString(excelRange.Cells[i, 28].Value2));
                unidadeInstObj.BolIsentoInscricaoMunicipal = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 29].Value2));
                unidadeInstObj.BolIsentoNire = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 30].Value2));
                unidadeInstObj.BolSinalizadoSicoob = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 31].Value2));
                unidadeInstObj.BolPaIncorporado = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 32].Value2));
                unidadeInstObj.DataHoraCarga = _validar.VerificarDadosData(Convert.ToString(excelRange.Cells[i, 11].Value2));
                unidadeInstObj.CodCriadoPor = 1;
                unidadeInstObj.DataHoraCriacao = DateTime.Now;

                unidadeInstList.Add(unidadeInstObj);
            }
        }

        Marshal.ReleaseComObject(excelWS);
        Marshal.ReleaseComObject(excelRange);
        excelWB.Close();
        Marshal.ReleaseComObject(excelWB);
        excelApp.Quit();
        Marshal.ReleaseComObject(excelApp);

        _dataContext.AddRange(unidadeInstList);
        _dataContext.SaveChanges();

        return Ok(unidadeInstList);
    }

    [HttpPost]
    [Route("TransportadoraValores")]
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
                var transportadoraObj = new TransportadoraValoresModel();

                transportadoraObj.Cnpj = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 1].Value2));
                transportadoraObj.DescricaoTransportadora = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 2].Value2));
                transportadoraObj.CodCriadoPor = 1;
                transportadoraObj.DataHoraCriacao = DateTime.Now;

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
    [Route("UnidadeInstituicaoTransportadoraValores")]
    public IActionResult InserirUnidadeInstTransportadoraValores()
    {
        var unidadeTransportadoraList = new List<UnidadeInstituicaoTransportadoraValoresModel>();

        Excel.Workbook excelWB = excelApp.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + @"\Resources\Transportadoras de Valores.xlsx");
        Excel._Worksheet excelWS = excelWB.Sheets[1];
        Excel.Range excelRange = excelWS.UsedRange;

        int rowCount = excelRange.Rows.Count;
        int columnCount = excelRange.Columns.Count;

        for (int i = 3; i <= rowCount; i++)
        {
            if (!string.IsNullOrEmpty(excelRange.Cells[i, 1].Value))
            {
                string cnpjTransportadora = excelRange.Cells[i, 1].Value2.ToString();

                var transportadoraObj = _dataContext.TransportadoraValores.Where(t => t.Cnpj.Contains(cnpjTransportadora)).FirstOrDefault();

                if (transportadoraObj != null)
                {
                    string[] unidadeInstList = excelRange.Cells[i, 3].Value2.ToString().Split(',');

                    foreach (string pa in unidadeInstList)
                    {
                        var unidadeTransportadoraObj = new UnidadeInstituicaoTransportadoraValoresModel();

                        unidadeTransportadoraObj.IdTransportadoraValores = transportadoraObj.IdTransportadoraValores;
                        unidadeTransportadoraObj.IdUnidadeInst = pa;
                        unidadeTransportadoraObj.CodCriadoPor = 1;
                        unidadeTransportadoraObj.DataHoraCriacao = DateTime.Now;

                        unidadeTransportadoraList.Add(unidadeTransportadoraObj);
                    }
                }
            }
        }

        Marshal.ReleaseComObject(excelWS);
        Marshal.ReleaseComObject(excelRange);
        excelWB.Close();
        Marshal.ReleaseComObject(excelWB);
        excelApp.Quit();
        Marshal.ReleaseComObject(excelApp);

        _dataContext.AddRange(unidadeTransportadoraList);
        _dataContext.SaveChanges();

        return Ok(unidadeTransportadoraList);
    }

    [HttpPost]
    [Route("TipoTerminal")]
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
                var tipoTerminalObj = new TipoTerminalModel();

                tipoTerminalObj.IdTipoTerminal = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 2].Value2));
                tipoTerminalObj.IdUnidadeInst = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 1].Value2));
                tipoTerminalObj.DescricaoTipoTerminal = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 3].Value2));
                tipoTerminalObj.BolAcessoLiberado = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 6].Value2));
                tipoTerminalObj.NumCheckAlteracao = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 7].Value2));
                tipoTerminalObj.LimiteSuperior = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 4].Value2));
                tipoTerminalObj.LimiteInferior = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 5].Value2));
                tipoTerminalObj.CodCriadoPor = 1;
                tipoTerminalObj.DataHoraCriacao = DateTime.Now;

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
    [Route("TipoOperacao")]
    public IActionResult InserirTipoOperacao()
    {
        var tipoOperacaoList = new List<TipoOperacaoModel>();

        Excel.Workbook excelWB = excelApp.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + @"\Resources\Sensibilização de Saldos.xlsx");
        Excel._Worksheet excelWS = excelWB.Sheets[1];
        Excel.Range excelRange = excelWS.UsedRange;

        int rowCount = excelRange.Rows.Count;
        int columnCount = excelRange.Columns.Count;

        for (int i = 2; i <= rowCount; i++)
        {
            if (!string.IsNullOrEmpty(excelRange.Cells[i, 3].Value))
            {
                var tipoOperacaoObj = new TipoOperacaoModel();

                tipoOperacaoObj.IdGrupoCaixa = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 1].Value2));
                tipoOperacaoObj.IdOperacaoCaixa = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 2].Value2));
                tipoOperacaoObj.Operacao = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 3].Value2));
                tipoOperacaoObj.DescricaoOperacao = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 4].Value2));
                tipoOperacaoObj.CodHistorico = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 5].Value2));
                tipoOperacaoObj.DescricaoHistorico = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 6].Value2));
                tipoOperacaoObj.Sensibilizacao = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 7].Value2));
                tipoOperacaoObj.CodCriadoPor = 1;
                tipoOperacaoObj.DataHoraCriacao = DateTime.Now;

                tipoOperacaoList.Add(tipoOperacaoObj);
            }
        }


        Marshal.ReleaseComObject(excelWS);
        Marshal.ReleaseComObject(excelRange);
        excelWB.Close();
        Marshal.ReleaseComObject(excelWB);
        excelApp.Quit();
        Marshal.ReleaseComObject(excelApp);

        _dataContext.AddRange(tipoOperacaoList);
        _dataContext.SaveChanges();

        return Ok(tipoOperacaoList);
    }

    [HttpPost]
    [Route("SaldosIniciais")]
    public IActionResult InserirSaldosIniciais()
    {
        var operacaoList = new List<OperacaoModel>();

        Excel.Workbook excelWB = excelApp.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + @"\Resources\SALDOS INICIAIS 01.09.2022.xlsx");

        var tipoTerminalList = new List<TipoTerminalModel> {
            new TipoTerminalModel
            {
                IdTipoTerminal = 2,
                DescricaoTipoTerminal = "CAIXAS"
            },
            new TipoTerminalModel
            {
                IdTipoTerminal = 5,
                DescricaoTipoTerminal = "ATMS "
            },
            new TipoTerminalModel
            {
                IdTipoTerminal = 8,
                DescricaoTipoTerminal = "TESOUREIROS ELETETRÔNICOS"
            }
        };

        //Console.WriteLine(excelWS.Name);
        //Console.WriteLine(excelWB.Sheets.Count);

        for (int t = 1; t <= excelWB.Sheets.Count; t++)
        {
            Excel._Worksheet excelWS = excelWB.Sheets[t];

            var tipoTerminal = tipoTerminalList.Where(t => t.DescricaoTipoTerminal == excelWS.Name).FirstOrDefault();

            if (tipoTerminal != null)
            {
                Excel.Range excelRange = excelWS.UsedRange;

                int rowCount = excelRange.Rows.Count;
                int columnCount = excelRange.Columns.Count;

                for (int i = 3; i <= rowCount; i++)
                {
                    if (!string.IsNullOrEmpty(excelRange.Cells[i, 6].Value))
                    {
                        var operacaoObj = new OperacaoModel();

                        operacaoObj.IdTipoTerminal = tipoTerminal.IdTipoTerminal;
                        operacaoObj.IdUnidadeInst = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 7].Value2));
                        operacaoObj.Operacao = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 1].Value2));
                        operacaoObj.DescricaoOperacao = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 2].Value2));
                        operacaoObj.CodHistorico = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 3].Value2));
                        operacaoObj.DescricaoHistorico = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 4].Value2));
                        operacaoObj.DataOperacao = _validar.VerificarDadosData(Convert.ToString(excelRange.Cells[i, 5].Value2));
                        operacaoObj.Terminal = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 6].Value2));
                        operacaoObj.CodigoAut = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 8].Value2));
                        operacaoObj.Valor = _validar.VerificarDadosMonetario(Convert.ToString(excelRange.Cells[i, 9].Value2));

                        var tipoOperacaoObj = _dataContext.TipoOperacao.Where(t => t.Operacao == operacaoObj.Operacao && t.CodHistorico == operacaoObj.CodHistorico).FirstOrDefault();

                        operacaoObj.IdTipoOperacao = tipoOperacaoObj.IdTipoOperacao;
                        operacaoObj.Sensibilizacao = tipoOperacaoObj.Sensibilizacao;
                        operacaoObj.CodCriadoPor = 1;
                        operacaoObj.DataHoraCriacao = DateTime.Now;

                        operacaoList.Add(operacaoObj);
                    }
                }

                Marshal.ReleaseComObject(excelWS);
                Marshal.ReleaseComObject(excelRange);
            }
        }

        excelWB.Close();
        Marshal.ReleaseComObject(excelWB);
        excelApp.Quit();
        Marshal.ReleaseComObject(excelApp);

        _dataContext.AddRange(operacaoList);
        _dataContext.SaveChanges();

        return Ok(operacaoList);
    }

    [HttpPost]
    [Route("Usuario")]
    public IActionResult InserirUsuario()
    {
        var usuarioList = new List<UsuarioModel>();

        Excel.Workbook excelWB = excelApp.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + @"\Resources\Base de Usuários e Terminais Atuailzada.xlsx");
        Excel._Worksheet excelWS = excelWB.Sheets[1];
        Excel.Range excelRange = excelWS.UsedRange;

        int rowCount = excelRange.Rows.Count;
        int columnCount = excelRange.Columns.Count;

        for (int i = 2; i <= rowCount; i++)
        {
            if (!string.IsNullOrEmpty(excelRange.Cells[i, 4].Value))
            {
                var usuarioObj = new UsuarioModel();

                usuarioObj.IdUsuario = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 4].Value2));
                usuarioObj.IdUnidadeInst = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 6].Value2));
                usuarioObj.IdInstituicao = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 1].Value2));
                usuarioObj.NumCheckAlteracao = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 3].Value2));
                usuarioObj.IdInstituicaoUsuario = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 5].Value2));
                usuarioObj.DescNomeUsuario = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 7].Value2));
                usuarioObj.CpfUsuario = null;
                usuarioObj.DataNascimentoUsuario = null;
                usuarioObj.DescEmail = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 11].Value2));
                usuarioObj.CelularUsuario = null;
                usuarioObj.BolHabilitadoUsuario = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 8].Value2));
                usuarioObj.DescStatusUsuario = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 9].Value2));
                usuarioObj.BolVerificaNomeMaquina = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 10].Value2));
                usuarioObj.CodCriadoPor = 1;
                usuarioObj.DataHoraCriacao = DateTime.Now;

                usuarioList.Add(usuarioObj);
            }
        }


        Marshal.ReleaseComObject(excelWS);
        Marshal.ReleaseComObject(excelRange);
        excelWB.Close();
        Marshal.ReleaseComObject(excelWB);
        excelApp.Quit();
        Marshal.ReleaseComObject(excelApp);

        _dataContext.AddRange(usuarioList);
        _dataContext.SaveChanges();

        return Ok(usuarioList);
    }

    [HttpPost]
    [Route("Terminal")]
    public IActionResult InserirTerminal()
    {
        var terminalList = new List<TerminalModel>();

        Excel.Workbook excelWB = excelApp.Workbooks.Open(System.IO.Directory.GetCurrentDirectory() + @"\Resources\TERMINAL.xlsx");
        Excel._Worksheet excelWS = excelWB.Sheets[1];
        Excel.Range excelRange = excelWS.UsedRange;

        int rowCount = excelRange.Rows.Count;
        int columnCount = excelRange.Columns.Count;

        for (int i = 2; i <= rowCount; i++)
        {
            if (!string.IsNullOrEmpty(excelRange.Cells[i, 8].Value))
            {
                var terminalObj = new TerminalModel();

                terminalObj.IdUnidadeInst = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 3].Value2));
                terminalObj.IdTipoTerminal = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 6].Value2));
                terminalObj.IdUsuario = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 8].Value2));
                terminalObj.IdUsuarioLiberacao = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 9].Value2));
                terminalObj.IdInstituicao = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 1].Value2));
                terminalObj.IdProduto = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 2].Value2));
                terminalObj.DataProcessamento = _validar.VerificarDadosData(Convert.ToString(excelRange.Cells[i, 4].Value2));
                terminalObj.NumTerminal = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 5].Value2));
                terminalObj.IdSituacaoTerminal = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 7].Value2));
                terminalObj.DescEstTrabalho = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 10].Value2));
                terminalObj.NumUltAutenticacao = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 11].Value2));
                terminalObj.NumLoteCco = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 12].Value2));
                terminalObj.MenorValorNota = _validar.VerificarDadosMonetario(Convert.ToString(excelRange.Cells[i, 13].Value2));
                terminalObj.DataHoraLiberacao = _validar.VerificarDadosData(Convert.ToString(excelRange.Cells[i, 14].Value2));
                terminalObj.NumLoteCheque = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 15].Value2));
                terminalObj.NumUltSeqLancCco = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 16].Value2));
                terminalObj.NumUltRemessa = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 17].Value2));
                terminalObj.IdClienteCor = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 18].Value2));
                terminalObj.NumLoteDoc = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 19].Value2));
                terminalObj.NumUltSeqDoc = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 20].Value2));
                terminalObj.DescVersaoSo = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 21].Value2));
                terminalObj.DescMemoriaRam = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 22].Value2));
                terminalObj.DescEspacoDisco = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 23].Value2));
                terminalObj.DescPacoteServico = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 24].Value2));
                terminalObj.NumLoteDec = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 25].Value2));
                terminalObj.NumUltSeqDec = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 26].Value2));
                terminalObj.ValorLimiteSaque = _validar.VerificarDadosMonetario(Convert.ToString(excelRange.Cells[i, 27].Value2));
                terminalObj.ValorLimiteTerminal = _validar.VerificarDadosMonetario(Convert.ToString(excelRange.Cells[i, 28].Value2));
                terminalObj.NumTesoureiro = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 29].Value2));
                terminalObj.NumIpTesoureiro = _validar.VerificarDadosTexto(Convert.ToString(excelRange.Cells[i, 30].Value2));
                terminalObj.CodTipoBalanceamento = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 31].Value2));
                terminalObj.NumTimeOutDispensador = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 32].Value2));
                terminalObj.CodLadoDepositario = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 33].Value2));
                terminalObj.NumPortaTesoureiro = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 34].Value2));
                terminalObj.NumCheckAlteracao = _validar.VerificarDadosInteiros(Convert.ToString(excelRange.Cells[i, 35].Value2));
                terminalObj.CodCriadoPor = 1;
                terminalObj.DataHoraCriacao = DateTime.Now;

                terminalList.Add(terminalObj);

                _dataContext.Add(terminalObj);
                _dataContext.SaveChanges();
            }
        }


        Marshal.ReleaseComObject(excelWS);
        Marshal.ReleaseComObject(excelRange);
        excelWB.Close();
        Marshal.ReleaseComObject(excelWB);
        excelApp.Quit();
        Marshal.ReleaseComObject(excelApp);

        //_dataContext.AddRange(terminalList);
        //_dataContext.SaveChanges();

        return Ok(terminalList);
    }

    [HttpPost]
    [Route("UsuarioSistema")]
    public IActionResult InserirUsuarioSistema()
    {
        var usuarioObj = new UsuarioModel();

        usuarioObj.IdUsuario = "Administrador";
        usuarioObj.IdUnidadeInst = "0";
        usuarioObj.IdInstituicao = 691;
        usuarioObj.NumCheckAlteracao = 0;
        usuarioObj.IdInstituicaoUsuario = 691;
        usuarioObj.DescNomeUsuario = "Administrador";
        usuarioObj.CpfUsuario = null;
        usuarioObj.DataNascimentoUsuario = null;
        usuarioObj.DescEmail = null;
        usuarioObj.CelularUsuario = null;
        usuarioObj.BolHabilitadoUsuario = 1;
        usuarioObj.DescStatusUsuario = null;
        usuarioObj.BolVerificaNomeMaquina = 0;
        usuarioObj.CodCriadoPor = 1;
        usuarioObj.DataHoraCriacao = DateTime.Now;

        _dataContext.Add(usuarioObj);
        _dataContext.SaveChanges();

        var usuarioSistemaObj = new UsuarioSistemaModel();

        usuarioSistemaObj.IdUsuario = "Administrador";
        usuarioSistemaObj.Login = "admin";
        usuarioSistemaObj.Password = "123";
        usuarioSistemaObj.SecretKey = "ABCDEFGHIJKL";
        usuarioSistemaObj.BolPrimeiroLogin = 1;
        usuarioSistemaObj.CodCriadoPor = 1;
        usuarioSistemaObj.DataHoraCriacao = DateTime.Now;

        var grupoAcessoObj = new GrupoAcessoModel();

        grupoAcessoObj.DescGrupoAcesso = "Administrador";
        grupoAcessoObj.CodCriadoPor = 1;
        grupoAcessoObj.DataHoraCriacao = DateTime.Now;
        
        _dataContext.Add(usuarioSistemaObj);
        _dataContext.Add(grupoAcessoObj);
        _dataContext.SaveChanges();

        var usuarioSistemaGrupoAcessoObj = new UsuarioSistemaGrupoAcessoModel();

        usuarioSistemaGrupoAcessoObj.IdUsuarioSistema = _dataContext.UsuarioSistema.Where(u => u.IdUsuario == usuarioSistemaObj.IdUsuario).First().IdUsuarioSistema;
        usuarioSistemaGrupoAcessoObj.IdGrupoAcesso = _dataContext.GrupoAcesso.Where(u => u.DescGrupoAcesso == grupoAcessoObj.DescGrupoAcesso).First().IdGrupoAcesso;
        usuarioSistemaGrupoAcessoObj.CodCriadoPor = 1;
        usuarioSistemaGrupoAcessoObj.DataHoraCriacao = DateTime.Now;

        _dataContext.Add(usuarioSistemaGrupoAcessoObj);
        _dataContext.SaveChanges();

        return Ok("Usuário cadastrado com sucesso!");
    }
}