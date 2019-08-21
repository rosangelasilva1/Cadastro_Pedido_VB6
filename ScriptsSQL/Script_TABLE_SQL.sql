USE [Vendas]
GO

/****** Object:  Table [dbo].[Item]    Script Date: 20/08/2019 15:16:31 ******/
/******Rosangela Oliveira da Silva**************/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


DROP TABLE ITEMPEDIDO
GO 
DROP TABLE Pedido 
GO
DROP TABLE SituacaoPedido 
GO
DROP TABLE Item 
GO 
 


CREATE TABLE [dbo].[Item](
	[Codigo] [int] NOT NULL,
	[Descricao] [varchar](max) NOT NULL,
	[ValorUnitario] [numeric](18, 2) NOT NULL,
 CONSTRAINT [PK_Item] PRIMARY KEY CLUSTERED 
(
	[Codigo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO



CREATE TABLE [dbo].[SituacaoPedido](
	[Codigo] [int] NOT NULL,
	[Descricao] [varchar](50) NOT NULL,
 CONSTRAINT [PK_SituacaoPedido] PRIMARY KEY CLUSTERED 
(
	[Codigo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Pedido](
	[Codigo] [int] NOT NULL,
	[Data] [datetime] NOT NULL,
	[CPF] [varchar](15) NOT NULL,
	[Solicitante] [varchar](200) NOT NULL,
	[Situacao] [int] NOT NULL,
	[ValorTotal] [numeric](18, 2) NOT NULL,
 CONSTRAINT [PK_Pedido] PRIMARY KEY CLUSTERED 
(
	[Codigo] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO




CREATE TABLE [dbo].[ItemPedido](
	[CodigoItem] [int] NOT NULL,
	[CodigoPedido] [int] NOT NULL,
	[Quantidade] [int] NULL,
	[ValorUnitarioItem] [int] NULL,
	[ValorTotalItem] [numeric](18, 2) NOT NULL,
 CONSTRAINT [PK_ItemPedido] PRIMARY KEY CLUSTERED 
(
	[CodigoItem] ASC,
	[CodigoPedido] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE [ItemPedido] ADD CONSTRAINT fk_Item_ItemPedido FOREIGN KEY ( CodigoItem ) REFERENCES Item ( Codigo ) ;
go
ALTER TABLE [ItemPedido] ADD CONSTRAINT fk_Pedido_ItemPedido FOREIGN KEY ( CodigoPedido ) REFERENCES Pedido ( Codigo ) ;
 go
ALTER TABLE [Pedido] ADD CONSTRAINT fk_Pedido_SituacaoPedido FOREIGN KEY ( Situacao ) REFERENCES SituacaoPedido ( Codigo ) ;

GO
insert into  situacaoPedido (Codigo, Descricao ) values (1,'Pendente')
go
insert into  situacaoPedido (Codigo, Descricao ) values (2,'Concluido')
GO
use Vendas
