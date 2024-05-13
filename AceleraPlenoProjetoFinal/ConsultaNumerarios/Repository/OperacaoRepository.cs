using ConsultaNumerarios.Data;
using ConsultaNumerarios.Interfaces;
using ConsultaNumerarios.Models;

namespace ConsultaNumerarios.Repository
{
    public class OperacaoRepository : IOperacao
    {
        private readonly DataContext _context;
        public OperacaoRepository(DataContext context) 
        {
            _context = context;
        }

        public List <OperacaoModel> GetOperacoesPorPa(int idPa)
        {
            return [.. _context.Operacao.Where(op => op.IdOperacao == idPa)];
        }
    }
}
