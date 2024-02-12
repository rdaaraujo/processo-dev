--CRIAÇÃO DA TABELA NO SCHEMA PUBLIC
CREATE TABLE public.psf_clientes (
  id SERIAL,
  nome_str TEXT NOT NULL,
  situacao_str TEXT NOT NULL,
  cpf_str TEXT NOT NULL,
  data_nascimento TEXT NOT NULL,
  endereco_str TEXT NOT NULL,
  telefone_Str TEXT NOT NULL,
  email_str TEXT NOT NULL,
  CONSTRAINT psf_clientes_pkey PRIMARY KEY(id)
);