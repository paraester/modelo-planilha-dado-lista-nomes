# modelo-planilha-dado-lista-nomes

### README.md`

```md
# Projeto: Gerador de Planilhas com Formatação e Dados de Nomes

Este projeto tem como objetivo gerar planilhas de Excel com múltiplas abas, onde cada aba representa um colaborador da lista de nomes. A formatação de cada aba é clonada a partir de uma aba modelo. Além disso, dados como o **nome do colaborador** e o **nome da área responsável** são inseridos automaticamente nas células **B1** e **B2** de cada aba.

## Funcionalidades

- Geração de planilhas `.xlsx` com múltiplas abas.
- Clonagem da formatação de uma aba modelo para todas as novas abas.
- Inserção automática de dados nas células B1 e B2:
  - **B1**: Nome da área responsável (extraído do arquivo `area.txt`).
  - **B2**: Nome completo do colaborador (extraído do arquivo `lista_nomes.txt`).
  
## Pré-requisitos

- Python 3.x instalado
- Bibliotecas necessárias:
  - `openpyxl` (para manipulação de arquivos Excel)
  - `pandas` (para leitura da lista de nomes)
  - `ttkbootstrap` (para a interface gráfica)

### Instalação das dependências

Você pode instalar as dependências usando `pip`:

```bash
pip install openpyxl pandas ttkbootstrap
```

## Estrutura do Projeto

- **tk_modelo_planilha_dado_lista-de-nomes.py**: Arquivo principal do projeto. Contém o código para gerar a planilha, clonar formatação e preencher as células com os dados.
- **lista_nomes.txt**: Arquivo de texto contendo os nomes dos colaboradores. Cada linha representa um nome.
- **area.txt**: Arquivo de texto contendo o nome da área responsável.
- **modelo.xlsx**: Planilha modelo. A formatação da primeira aba será clonada para todas as novas abas.

### Exemplo de `lista_nomes.txt`

```txt
João da Silva
Maria Oliveira
Carlos Souza
Ana Costa
```

### Exemplo de `area.txt`

```txt
Departamento de Recursos Humanos
```

## Como Executar o Projeto

1. **Obtenha os arquivos de entrada**:
   - Um arquivo `modelo.xlsx`, que será utilizado como referência para a formatação de todas as abas.
   - Um arquivo `lista_nomes.txt`, contendo os nomes completos dos colaboradores (um nome por linha).
   - Um arquivo `area.txt`, contendo o nome da área responsável.

2. **Execute o script**:
   - Execute o arquivo `tk_modelo_planilha_dado_lista-de-nomes.py` utilizando Python.

```bash
python tk_modelo_planilha_dado_lista-de-nomes.py
```

3. **Interface Gráfica**:
   - A interface irá abrir. Nela, você deve selecionar os três arquivos (`modelo.xlsx`, `lista_nomes.txt`, e `area.txt`).
   - Escolha onde salvar o arquivo gerado e o nome do arquivo.
   - Clique em "Criar Planilha". O script irá gerar a planilha com múltiplas abas, uma para cada colaborador da lista, aplicando a formatação do modelo.

4. **Resultado**:
   - Um arquivo `.xlsx` será gerado contendo múltiplas abas. Cada aba representará um colaborador, com os dados inseridos nas células B1 (área responsável) e B2 (nome completo).

## Como Contribuir

- **Fork** este repositório.
- Crie uma nova branch para suas alterações.
- Envie um **Pull Request** com uma descrição detalhada das alterações.

## Licença

Este projeto é licenciado sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.
```

### Explicações:

1. **Pré-requisitos e dependências**: Detalhei as bibliotecas Python necessárias e o comando para instalá-las usando `pip`.
2. **Descrição do projeto**: Expliquei o propósito do projeto e como ele funciona.
3. **Estrutura do projeto**: Mostra o que cada arquivo faz.
4. **Exemplos de arquivos de entrada**: Exemplos de como devem ser os arquivos `lista_nomes.txt` e `area.txt`.
5. **Como executar o script**: Detalha o processo para executar o projeto usando a interface gráfica e como o resultado será gerado.
6. **Instruções de contribuição**: Para quem quiser contribuir com o código.

---
