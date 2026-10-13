---
title: Documento de teste
author: Exemplo
---

# Relatório de Teste do Conversor

Este parágrafo tem **negrito**, *itálico*, ***negrito e itálico***, ~~tachado~~, `código inline`,
um [link para o site da Fatec](https://www.fatecsp.br) e texto com H<sub>2</sub>O e E = mc<sup>2</sup>.
Também há <u>sublinhado</u> e <mark>destaque</mark>. Acentuação: ação, coração, pão, você, à.

Quebra de linha forçada com dois espaços  
continua na linha seguinte.

## Listas

- Item com marcador
- Segundo item
  - Subitem aninhado
    - Terceiro nível
- Terceiro item

1. Primeiro passo
2. Segundo passo
   1. Sub-passo a
   2. Sub-passo b
3. Terceiro passo

Lista numerada que começa em 5:

5. Cinco
6. Seis

Lista de tarefas:

- [x] Escrever o conversor
- [ ] Testar no Word

Lista "solta" com parágrafos:

1. Primeiro item com um parágrafo longo.

   Segundo parágrafo do mesmo item, indentado.

2. Segundo item.

## Citação

> Esta é uma citação.
> Continua na mesma citação.
>
> > Citação aninhada.

## Código

```python
def fatorial(n: int) -> int:
    """Calcula n!"""
    if n <= 1:
        return 1
    return n * fatorial(n - 1)
```

## Tabela

| Amostra | Espessura (nm) | Composição | Observação |
|:--------|---------------:|:----------:|------------|
| A1      | 120,5          | TeO₂       | **ok**     |
| A2      | 98,2           | TeO₃       | refazer    |
| B1      | 210,0          | `mix`      | ver nota[^nota] |

## Equações

A energia é $E = h\nu$ e o desvio padrão é $\sigma = \sqrt{\frac{1}{N}\sum_{i=1}^{N}(x_i-\mu)^2}$.
Preços em reais não viram equação: R$ 5,00 e R$ 10,00.

$$
\chi^2_{red} = \frac{1}{N-p}\sum_{i=1}^{N}\frac{(y_i - f(x_i))^2}{\sigma_i^2}
$$

## Imagem

![Gráfico de exemplo](grafico_exemplo.png)

![Imagem remota](https://example.com/foto.png)

---

## Notas

Uma afirmação com nota de rodapé[^1] e outra[^longa].

[^1]: Esta é a primeira nota de rodapé, com **negrito**.
[^longa]: Nota com link: [Python](https://www.python.org).
[^nota]: Nota vinda de dentro da tabela.

### Título de nível 3

#### Título de nível 4

Fim do documento.
