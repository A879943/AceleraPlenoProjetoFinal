using ConsultaNumerarios.Data;
using ConsultaNumerarios.Interfaces;
using ConsultaNumerarios.Models;

namespace ConsultaNumerarios.Repository
{
    public class UnidadeInstituicaoRepository : IUnidadeInstituicao
    {
        private readonly DataContext _context;
        public UnidadeInstituicaoRepository(DataContext context)
        {
            _context = context;
        }

        public List<UnidadeInstituicaoModel> GetUnidadeInstituicao()
        {
            //TODO validar campo para indicar que o PA está ativo.
            return [.. _context.UnidadeInstituicao.Where(pa => pa.DataFimFuncionamento != null)];
        }
    }
}
