# Python Nativo no Excel M365

Códigos em Python nativo do Excel M365 usando apenas bibliotecas padrão e `pandas`. Exemplos práticos com a função `xl()` para leitura e escrita em planilhas. Ideal para automação, filtros e análises diretamente no Excel, sem dependências externas. Scripts simples, reutilizáveis e compatíveis com o ambiente Python do Excel.

## Objetivo

Reunir snippets de código prontos para uso direto em planilhas com suporte a Python nativo no Microsoft Excel 365.

## Requisitos

* Microsoft 365 com suporte ao Python nativo
* Nenhuma instalação adicional necessária

## Estrutura

* `filtro_excel.py`: exemplo de filtro dinâmico com múltiplos critérios e conversão de tipos

## Exemplo de uso

```python
# Em uma célula do Excel com Python:
=PY("""
import pandas as pd
df = xl("MinhaTabela[#Tudo]", headers=True)
df_filtrado = df[df['Valor'] > 1000]
df_filtrado
""")
```

## Licença

Este repositório está sob a licença MIT. Sinta-se à vontade para usar e contribuir!
