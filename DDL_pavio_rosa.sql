-- CRIANDO O BANCO DE DADOS

CREATE DATABASE db_pavio_rosa
GO

-- CRIANDO A TABELA DIMENSAO CLIENTE

CREATE TABLE dClientes(
	ID_cliente INT NOT NULL IDENTITY(1, 1),
	ID_localidade INT NOT NULL,
	Data_cadastro DATETIME NOT NULL,
	Data_atualizacao DATETIME DEFAULT NULL,
	Nome VARCHAR(50) NOT NULL,
	Sobrenome VARCHAR(100),
	Data_nascimento DATE,
	Genero VARCHAR(30),
	Email VARCHAR(100),
	Telefone VARCHAR(16),
	Rua VARCHAR(100),
	Numero VARCHAR,
	Bairro VARCHAR(50),
	Cidade VARCHAR(50),
	Estado VARCHAR(2),
	CEP VARCHAR(9),
	CONSTRAINT dClientes_id_cliente_pk PRIMARY KEY(ID_cliente)
)

GO

-- CRIANDO A TABELA DIMENSAO PRODUTO

CREATE TABLE dProdutos(
	Cod_produto VARCHAR(10) NOT NULL,
	Data_cadastro DATETIME NOT NULL,
	Data_atualizacao DATETIME DEFAULT NULL,
	Produto VARCHAR(100) NOT NULL,
	Preco DECIMAL(5, 2) NOT NULL,
	CONSTRAINT dProdutos_cod_produto_un UNIQUE(Cod_produto),
	CONSTRAINT dProdutos_id_produto_pk PRIMARY KEY(Cod_produto),
	CONSTRAINT dProdutos_preco_ck CHECK(Preco > 0)
)

GO

-- CRIANDO A TABELA FATO TRANSACOES

CREATE TABLE fTransacoes(
	ID_venda INT IDENTITY(1, 1),
	Data_venda DATE NOT NULL,
	Cod_produto VARCHAR(10) NOT NULL,
	ID_cliente INT NOT NULL,
	Quantidade INT NOT NULL,
	Preco DECIMAL(5, 2) NOT NULL,
	Valor_total DECIMAL (7, 2) NOT NULL,
	CONSTRAINT fTransacoes_id_venda_pk PRIMARY KEY(ID_venda),
	CONSTRAINT fTransacoes_cod_produto_fk FOREIGN KEY(Cod_produto) REFERENCES dProdutos(Cod_produto),
	CONSTRAINT fTransacoes_id_cliente_fk FOREIGN KEY(ID_cliente) REFERENCES dClientes(ID_cliente),
	CONSTRAINT fTransacoes_quantidade_ck CHECK(Quantidade > 0),
	CONSTRAINT fTransacoes_preco_ck CHECK(Preco > 0),
	CONSTRAINT fTransacoes_valor_total_ck CHECK(Valor_Total > 0)
)

GO