using _4___Domain;
using _3___Infra;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _2___Services
{
    public class Crud
    {
        public PostgresRepository postgresRepository = new PostgresRepository();
        public void insertData(DadosCli dados)
        {
            var query = $"INSERT INTO psf_clientes (" +
                        $"nome_str, situacao_str, cpf_str, data_nascimento, endereco_str, telefone_str, email_str)" +
                        $"values (" +
                        $"'{dados.Nome}', '{dados.Situacao}','{dados.Cpf}', '{dados.DataNasc}', '{dados.Endereco}', '{dados.Telefone}', '{dados.Email}')";
            postgresRepository.GenericConnection(query);
        }

        public void updateData(int Id)
        {
            var query = $"UPDATE psf_clientes SET " +
                        $"situacao_str = 'BAIXADO'" +
                        $"where id = '{Id}'";
            postgresRepository.GenericConnection(query);
        }

        public void deleteData(string Id)
        {
            var query = $"DELETE FROM psf_clientes pv " +
                        $"WHERE pv.id = '{Id}'";
            postgresRepository.GenericConnection(query);
        }

        public void deleteAllData(string Id)
        {
            var query = $"DELETE FROM psf_clientes pv ";
            postgresRepository.GenericConnection(query);
        }

        public List<DadosCli> getData()
        {
            var query = $"SELECT * FROM psf_clientes";
                        /*$"WHERE id = '{Id}'";*/

            var result = postgresRepository.GetConnection(query);
            
            return result;

        }

    }
}
