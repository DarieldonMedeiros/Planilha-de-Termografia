# Planilha para Automatizar Relatórios Fotográficos de Termografia

## Descrição

Este arquivo Excel (.xlsm) foi desenvolvido para automatizar a geração de relatórios fotográficos de termografia dos inversores, utilizado no Consórcio Marangatu (Motrice Soluções em Energia e BN Engenharia). Ele utiliza macros em VBA para facilitar e agilizar o processo de criação dos relatórios, integrando dados e imagens de forma automática.

## Como Funciona

O funcionamento do arquivo é baseado em três bases de dados principais:

- **BD TITULO**: Contém os títulos e informações principais de cada relatório ou item a ser documentado.
- **BD IMAGEM**: Armazena os caminhos ou referências das imagens capturadas pelas câmeras termográficas.
- **BD PH**: Possui dados complementares, como parâmetros de medição, observações ou outros campos necessários para o relatório.

O código VBA presente no arquivo realiza as seguintes etapas:

1. **Geração das Planilhas**: A partir das bases de dados, o sistema cria automaticamente as planilhas de relatório para cada item ou equipamento listado.
2. **Inserção Automática de Imagens**: Assim que as planilhas são geradas, as imagens termográficas correspondentes são automaticamente inseridas nos locais apropriados, conforme definido na base de dados de imagens.
3. **Preenchimento de Dados**: Os campos de cada relatório são preenchidos automaticamente com as informações das bases de dados.
4. **Pronto para Impressão**: Após a geração e preenchimento, os relatórios já estão prontos para serem impressos, otimizando o tempo e reduzindo erros manuais.

## Benefícios

- **Automação**: Elimina a necessidade de inserir manualmente dados e imagens, tornando o processo mais rápido e confiável.
- **Padronização**: Garante que todos os relatórios sigam o mesmo formato e padrão de qualidade.
- **Facilidade de Uso**: Basta alimentar as bases de dados e acionar a macro para gerar todos os relatórios automaticamente.

## Requisitos

- Microsoft Excel com suporte a macros (habilitar conteúdo VBA ao abrir o arquivo).
- As imagens termográficas devem estar organizadas conforme o esperado pela base de dados de imagens.

## Como Usar

1. Abra o arquivo `F 157_Relatório Fotográfico Termografia 12.1.xlsm` no Excel.
2. Habilite as macros (caso solicitado).
3. Preencha ou atualize as bases de dados: BD TITULO, BD IMAGEM e BD PH.
4. Execute a macro principal (normalmente há um botão ou atalho na planilha para isso).
5. Aguarde a geração automática dos relatórios.
6. Após a conclusão, os relatórios estarão prontos para revisão e impressão.

## Observações

- Certifique-se de que os caminhos das imagens estejam corretos e acessíveis.
- Caso precise adaptar o modelo para outros tipos de relatório, será necessário ajustar as bases de dados e, possivelmente, o código VBA.

## Contato

Quaisquer dúvidas, entrem em contato comigo:

[![Email](https://img.shields.io/badge/Email-darieldonbm99@outlook.com-blue?style=for-the-badge&logo=microsoft-outlook)](mailto:darieldonbm99@outlook.com)
[![WhatsApp](<https://img.shields.io/badge/WhatsApp-%2B55%20(86)%2099466--6480-25D366?style=for-the-badge&logo=whatsapp>)](https://wa.me/5586994666480)

---
