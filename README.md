
# Automação do Controle de Helados

Script em Python que gera, de uma vez só, a planilha de controle de helados de cada loja a partir de um modelo padrão e de uma lista de lojas.

Eu trabalho com esse controle no dia a dia, em uma operação com dezenas de lojas. Antes, cada planilha era montada à mão (copiar o modelo, renomear, preencher o nome da loja, preencher o código em cada linha), o que levava muito tempo e era desgastante. Fiz este script para automatizar isso: hoje o processo todo leva cerca de 1 minuto.

> Todos os dados deste repositório são fictícios. A lista de lojas e os códigos são inventados, e as planilhas reais de cada mês ficam fora do repositório (veja a seção "Privacidade").

## O que ele faz

Para cada loja da lista, o script:

1. copia o modelo padrão e salva com o nome `Controle de Helados - <nome da loja>.xlsx`;
2. preenche o nome da loja na aba `01` (célula `D4`);
3. preenche o código da loja na aba `CONSOLIDADO` (células `A7` a `A22`), desprotegendo a aba só durante a escrita e protegendo de novo ao final, igual ao modelo original;
4. mostra no terminal o andamento e, no fim, quantas planilhas foram geradas e quais deram erro.

Se uma loja der problema, o script não para: registra o erro, segue para a próxima e mostra o resumo no final.

## Como usar

**Requisitos:** Python 3 e as bibliotecas `pandas` e `openpyxl`.

```
pip install pandas openpyxl
```

**Arquivos necessários**, na mesma pasta do `helados.py`:

- `Controle de Helados - Modelo Padrao.xlsx`: o modelo, com as abas `01` e `CONSOLIDADO`;
- `exemplo_lista_de_lojas.xlsx`: a lista de lojas, com as colunas `Nome da loja` e `Codigo Linx`. Também aceita `.csv`.

**Rodando:**

```
python helados.py
```

As planilhas são geradas na pasta definida em `PASTA_SAIDA`, no início do script. No dia a dia eu mudo esse nome para o do mês (por exemplo, `planilhas_dezembro`).

## Configuração

No topo do `helados.py` ficam as variáveis que mudam de um uso para outro:

| Variável                  | O que é                               |
| -------------------------- | -------------------------------------- |
| `ARQUIVO_MODELO`         | nome do modelo padrão                 |
| `ARQUIVO_LOJAS`          | lista de lojas (`.xlsx` ou `.csv`) |
| `PASTA_SAIDA`            | pasta onde as planilhas são geradas   |
| `COL_NOME` / `COL_COD` | nomes das colunas na lista de lojas    |

## Privacidade

O `.gitignore` deixa de fora as pastas `planilhas_*/` e a lista real de lojas, então os dados reais nunca chegam ao repositório. O que está aqui é só o código, um modelo vazio e uma lista de exemplo com lojas e códigos inventados.

## Tecnologias

Python, pandas e openpyxl.

## Possíveis melhorias

- Validar a lista de lojas antes de gerar (códigos duplicados, células vazias).
- Receber o modelo, a lista e a pasta de saída por argumentos de linha de comando, em vez de editar o código.
- Interface simples para quem não usa terminal.
