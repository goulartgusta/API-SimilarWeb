# API SimilarWeb
Essa é uma API que utiliza Pandas e Openpyxl, com integração à API do Google Sheets. Realizando um ETL, os dados são extraídos pela API em formato JSON diretamente da SimilarWeb (Site para análise de tráfego entre sites e apps), transformando-os em DataFrames. Em seguida, esses dados são tratados utilizando as bibliotecas, sendo carregados em planilhas do Excel e do Google Sheets. O sistema fornece informações detalhadas sobre o tráfego de varejo e o tráfego por canal de maneira rápida e eficiente.

## Objetivo
O principal objetivo da API é otimizar o tempo e o esforço necessários na extração de dados diretamente do site, reduzindo e muito o tempo e energia para obter informações no uso de análises e relatórios.

![image](https://github.com/goulartgusta/API-SimilarWeb/assets/101681743/00831e77-e0d0-4b52-983d-dc19ce89b37f)
(Fluxograma retratado de forma a atender conhecimento técnico do cliente)

## Segurança
Para garantir a segurança, as credenciais estão protegidas e ocultas. O funcionamento do código requer credenciais (Keys) para as APIs do SimilarWeb e Google Sheets, também é necessíaro as autorizações dessas chaves. Para o uso correto, basta atualizar as credenciais com suas prórpias.

## Planejamento Futuro
O projeto está atualmente disponível em sua forma inicial, mas tenho planos para algumas melhorias na organização e boas práticas de programação. Facilitando a manutenção e garantindo uma boa compreensão do código.

> Otimizações Planejadas:
- Implementação de boas práticas de programação para otimizar laços e estruturas de código.
- Melhor organização de arquivos e pastas para facilitar o gerenciamento do projeto.
- Automação da inserção de datas, eliminando a necessidade de realizar manualmente.
- Expansão da base de dados com mais informações de varejos/websites visando aumentar a abrangência da API.
- Criação de um dashboard dinâmico com dados extraídos do SimilarWeb via API.
