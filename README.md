# Power Query M Functions 📦

Repositório de funções personalizadas do Power Query (linguagem M) para Excel e Power BI, criadas com auxílio de Inteligência Artificial.

## 📌 Sobre este Repositório

Este repositório centraliza funções customizadas do Power Query organizadas por segmentos/categorias, facilitando a reutilização e manutenção de código em projetos de análise de dados.

## 📂 Estrutura por Segmentos

### 📝 Tratamento de Texto
Funções para manipulação, limpeza e transformação de strings.

- **[SubstituirTextoExato.pq](SubstituirTextoExato.pq)** - Substitui valores completos usando tabela de-para (match exato, não frações)

### 🎯 Tratamento de Cabeçalhos
*Em breve: Funções para normalização e padronização de cabeçalhos de tabelas*

### ⚙️ Transformação de Colunas
*Em breve: Funções para manipulação avançada de colunas*

### 🔀 Condicionais
*Em breve: Funções para lógica condicional complexa*

### 📅 Tratamento de Datas
*Em breve: Funções para manipulação de datas e períodos*

### 📊 Análise e Validação
*Em breve: Funções para validação de dados e qualidade*

---

## 📖 Documentação das Funções

### SubstituirTextoExato

**Categoria:** Tratamento de Texto  
**Arquivo:** [SubstituirTextoExato.pq](SubstituirTextoExato.pq)

**Descrição:**  
Substitui valores de uma coluna usando uma tabela de-para, realizando correspondência exata do campo completo. Não substitui frações ou partes de texto.

**Parâmetros:**
- `TabelaOriginal` (table): Tabela onde será feita a substituição
- `NomeColuna` (text): Nome da coluna onde os valores serão substituídos
- `TabelaSubstituicao` (table): Tabela de-para com colunas "TextoAntigo" e "TextoNovo"

**Retorno:**  
Tabela com os valores substituídos

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
- ✅ "São Paulo" → "SP" (substitui campo completo)
- ❌ "São Paulo Centro" → "São Paulo Centro" (não substitui - não é match exato)
- ❌ "Paulo" → "Paulo" (mantém original - sem correspondência)

---

## 🚀 Como Usar

### Método 1: Importar Função Individual

1. Acesse o arquivo `.pq` da função desejada neste repositório
2. Copie o código da função
3. No Power Query Editor:
   - Vá em **Início** → **Nova Consulta** → **Outras Fontes** → **Consulta em Branco**
   - Clique em **Editor Avançado**
   - Cole o código da função
   - Clique em **Concluído**
4. Renomeie a consulta com o nome da função (ex: `SubstituirTextoExato`)
5. A função estará disponível para uso em outras queries

### Método 2: Carregar do GitHub (Direto)

```m
let
    Fonte = Web.Contents("https://raw.githubusercontent.com/dennisdmn/PowerQuery-M-Functions/main/SubstituirTextoExato.pq"),
    Funcao = Expression.Evaluate(Text.FromBinary(Fonte), #shared)
in
    Funcao
```

## 🤖 Sobre a Criação com IA

Todas as funções deste repositório foram desenvolvidas com auxílio de Inteligência Artificial, combinando:
- Expertise em Power Query e linguagem M
- Boas práticas de programação funcional
- Documentação detalhada e exemplos práticos
- Testes e validações de casos de uso reais

## 💬 Contribuições

Sugestões, melhorias e novas funções são bem-vindas!  
Sinta-se à vontade para abrir issues ou pull requests.

## 📜 Licença

MIT License - Uso livre para projetos pessoais e comerciais.

---

**Desenvolvido por:** [@dennisdmn](https://github.com/dennisdmn)  
**Última atualização:** Janeiro 2026