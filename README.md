# Automatização RSC

Programa em Python para reunir lançamentos RSC de arquivos ODS em uma única planilha para macro.

## Executar

Requer Python 3.9 ou superior com Tcl/Tk (tkinter). Não requer bibliotecas externas.

```sh
python extrair_rsc.py
```

Selecione a pasta das origens e clique em **Processar planilhas**. O programa lê a primeira aba de cada ODS diretamente nessa pasta.

## Lançamentos

| Campo | Origem |
| --- | --- |
| SIAPE/7 | B8 |
| Nome | B7 |
| MÊS/24 | B10 |
| SEQ | B11 |
| Justificativa | I11 |
| Doc legal | I10 |

I18 maior que zero gera uma linha com rubrica E15 e R/D `r`.
I24 maior que zero gera outra linha com rubrica E21 e R/D `d`.
Os valores são arredondados e exibidos com exatamente duas casas decimais.

## Resultados

Os resultados ficam na pasta `Lançamentos RSC - mês ano`:

- `Entrada para macro - X planilhas incluídas.ods`: X conta origens com ao menos um lançamento.
- `Relatório de Execução - RSC lançamento DD-MM-AAAA.txt`: resumo, nomes dos arquivos com erro e detalhes por origem.

Nomes repetidos recebem um sufixo numérico. Os arquivos originais são preservados.
Recalcule e salve as origens no LibreOffice antes da execução; o programa utiliza os resultados salvos das fórmulas.

## Interface

Identidade UFGD, seleção de pasta, cartões com gradiente, registro de execução, barra animada e confirmação de conclusão. O logo está incorporado ao Python.

