using _4___Domain;
using Npgsql;
using System;
using System.Reflection.PortableExecutable;
using System.Security.AccessControl;
using System.Text;

namespace _3___Infra
{
    public class PostgresRepository
    {
        public string connString = "Host=Localhost;Username=admin;Password=admin;Database=postgres";

        public void GenericConnection(string query)
        {
            using (var conn = new NpgsqlConnection(connString))
            {
                conn.Open();

                using (var cmd = new NpgsqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = query;
                    cmd.Parameters.AddWithValue("id", 123);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Console.WriteLine("{0} {1}", reader.GetInt32(0), reader.GetString(1));
                        }
                    }
                }
            }
        }
        public List<DadosCli> GetConnection(string query)
        {
            using var conn = new NpgsqlConnection(connString);
            conn.Open();
            using var cmd = new NpgsqlCommand(query, conn);
            using NpgsqlDataReader reader = cmd.ExecuteReader();

            var lista = new List<DadosCli>();
            while (reader.Read())
            {
                lista.Add(new DadosCli()
                {
                    Id = reader.GetInt32(0),
                    Nome = reader.GetString(1),
                    Situacao = reader.GetString(2),
                    Cpf = reader.GetString(3),
                    DataNasc = reader.GetString(4),
                    Endereco = reader.GetString(5),
                    Telefone = reader.GetString(6),
                    Email = reader.GetString(7)
                });
            }
            conn.Close();

            return lista;
        }
    }
}


