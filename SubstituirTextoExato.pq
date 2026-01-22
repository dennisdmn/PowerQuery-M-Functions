/*
Função: SubstituirTextoExato
Descrição: Substitui valores de uma coluna usando uma tabela de-para, apenas quando há correspondência exata (não substitui frações de texto)

Parâmetros:
- TabelaOriginal: Tabela onde será feita a substituição
- NomeColuna: Nome da coluna (texto) onde os valores serão substituídos
- TabelaSubstituicao: Tabela de-para com colunas "TextoAntigo" e "TextoNovo"

Retorno: Tabela com os valores substituídos

Exemplo de uso:
    let
        Fonte = Excel.CurrentWorkbook(){[Name="MinhaTabela"]}[Content],
        TabelaDePara = Excel.CurrentWorkbook(){[Name="TabelaDePara"]}[Content],
        Resultado = SubstituirTextoExato(Fonte, "Estado", TabelaDePara)
    in
        Resultado

Estrutura da TabelaDePara:
    | TextoAntigo    | TextoNovo |
    |----------------|-----------||
    | São Paulo      | SP        |
    | Rio de Janeiro | RJ        |
    | Minas Gerais   | MG        |
*/

(TabelaOriginal as table, NomeColuna as text, TabelaSubstituicao as table) =>
let
    // Converte a tabela de-para em uma lista de registros para lookup
    ListaPares = List.Zip({
        Table.Column(TabelaSubstituicao, "TextoAntigo"),
        Table.Column(TabelaSubstituicao, "TextoNovo")
    }),
    
    // Cria uma tabela de lookup para busca mais eficiente
    TabelaLookup = Table.FromColumns(
        {Table.Column(TabelaSubstituicao, "TextoAntigo"),
         Table.Column(TabelaSubstituicao, "TextoNovo")},
        {"Antigo", "Novo"}
    ),
    
    // Aplica substituição apenas quando o texto corresponde EXATAMENTE
    Resultado = Table.TransformColumns(
        TabelaOriginal,
        {{NomeColuna, each 
            try TabelaLookup{[Antigo=_]}[Novo] 
            otherwise _
        }}
    )
in
    Resultado