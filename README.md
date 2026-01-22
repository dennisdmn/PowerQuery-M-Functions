# Power Query M Functions

Coleção de funções personalizadas do Power Query (M) para Excel e Power BI.

## Funções Disponíveis

### SubstituirTextoExato

**Descrição:** Substitui valores de uma coluna usando uma tabela de-para, apenas quando há correspondência exata do campo completo (não substitui frações de texto).

**Parâmetros:**
- `TabelaOriginal` (table): Tabela onde será feita a substituição
- `NomeColuna` (text): Nome da coluna onde os valores serão substituídos
- `TabelaSubstituicao` (table): Tabela de-para com colunas "TextoAntigo" e "TextoNovo"

**Retorno:** Tabela com os valores substituídos

**Estrutura da Tabela De-Para:**
```
| TextoAntigo    | TextoNovo |
|----------------|-----------||
| São Paulo      | SP        |
| Rio de Janeiro | RJ        |
| Minas Gerais   | MG        |
```

**Exemplo de Uso:**
```m
let
    Fonte = Excel.CurrentWorkbook(){[Name="MinhaTabela"]}[Content],
    TabelaDePara = Excel.CurrentWorkbook(){[Name="TabelaDePara"]}[Content],
    Resultado = SubstituirTextoExato(Fonte, "Estado", TabelaDePara)
in
    Resultado
```

**Comportamento:**
- "São Paulo" → "SP" (substitui campo completo)
- "São Paulo Centro" → "São Paulo Centro" (não substitui, pois não é exatamente "São Paulo")
- "Rio" → "Rio" (mantém original se não houver correspondência exata)

## Como Usar no Power Query

1. Baixe o arquivo `.pq` da função desejada
2. No Power Query Editor, vá em **Início** → **Nova Consulta** → **Outras Fontes** → **Consulta em Branco**
3. Abra o **Editor Avançado** e cole o código da função
4. Renomeie a consulta com o nome da função (ex: `SubstituirTextoExato`)
5. Use a função em suas queries conforme os exemplos

## Contribuições

Sugestões e melhorias são bem-vindas!

## Licença

MIT License - Uso livre para projetos pessoais e comerciais.