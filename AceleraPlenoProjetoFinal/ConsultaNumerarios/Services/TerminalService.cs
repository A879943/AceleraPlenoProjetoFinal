using ConsultaNumerarios.Dto;
using ConsultaNumerarios.Interfaces;

namespace ConsultaNumerarios.Services
{
    public class TerminalService : ITerminalService
    {
        private readonly IOperacao _operacao;
        public TerminalService(IOperacao operacao)
        {
            _operacao = operacao;
        }

        public SaldoDeTerminais CalculaSaldosDeTerminais(int idPaParsed)
        {
            var operacoes = _operacao.GetOperacoesPorPa(idPaParsed);
            var listaTerminais = operacoes.Select(op => op.Terminal).Distinct().ToList();

            var retorno = new SaldoDeTerminais();

            foreach (var t in listaTerminais)
            {
                var operacoesSelecionadas = operacoes.Select(op => op).Where(op => op.Terminal == t);
                var terminal = new Terminal
                {
                    Id = operacoesSelecionadas.FirstOrDefault().IdTipoTerminal,
                };
                foreach (var o in operacoesSelecionadas)
                {
                    retorno.Terminal.Add(new Terminal
                    {
                        Id = o.IdTipoTerminal,

                    });
                }
            }
            return retorno;
        }


    }
}
