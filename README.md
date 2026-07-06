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
Funções para validação de dados, qualidade, auditoria de campos obrigatórios e conferência de layouts.

- **[fxValidaCamposObrigatorios.pq](validacao/fxValidaCamposObrigatorios.pq)** - Valida múltiplas colunas obrigatórias informadas como texto separado por vírgula e retorna campos em branco por linha

### 🧪 Exemplos
Consultas de exemplo para testar e adaptar as funções do repositório.

- **[fxValidaCamposObrigatorios_exemplo.pq](exemplos/fxValidaCamposObrigatorios_exemplo.pq)** - Exemplo com base contábil simulada e validação de campos obrigatórios

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

**Exemplo de Uso:**
```m
let
    Fonte = Excel.CurrentWorkbook(){[Name="MinhaTabela"]}[Content],
    TabelaDePara = Excel.CurrentWorkbook(){[Name="TabelaDePara"]}[Content],
    Resultado = SubstituirTextoExato(Fonte, "Estado", TabelaDePara)
in
    Resultado
```

---

### fxValidaCamposObrigatorios

**Categoria:** Análise e Validação  
**Arquivo:** [validacao/fxValidaCamposObrigatorios.pq](validacao/fxValidaCamposObrigatorios.pq)  
**Exemplo:** [exemplos/fxValidaCamposObrigatorios_exemplo.pq](exemplos/fxValidaCamposObrigatorios_exemplo.pq)

**Descrição:**  
Valida, linha a linha, se colunas obrigatórias estão nulas, vazias ou preenchidas apenas com espaços. As colunas obrigatórias são informadas em um único texto separado por vírgula, facilitando o uso pela janela **Invocar Função Personalizada** do Power Query.

**Parâmetros:**
- `Tabela` (table): Tabela que será validada
- `ColunasObrigatoriasTexto` (text): Lista de colunas obrigatórias em texto, separadas por vírgula
- `NomeColunaErros` (nullable text, opcional): Nome da coluna que listará os campos em branco. Padrão: `Campos_Em_Branco`
- `NomeColunaStatus` (nullable text, opcional): Nome da coluna de status. Padrão: `Status_Validacao`

**Retorno:**  
Tabela original acrescida de duas colunas:
- `Campos_Em_Branco`: lista das colunas obrigatórias em branco naquela linha
- `Status_Validacao`: retorna `OK` ou `Campo obrigatório em branco`

**Exemplo de Uso:**
```m
let
    Fonte = Etapa_DefineTipos,

    Etapa_ValidaCamposObrigatorios = fxValidaCamposObrigatorios(
        Fonte,
        "empresa, filial, debito, credito, valor, hist, ccustod, ccustoc, usuario, clvlrdeb, clvlrcrd, ent05deb, ent05crd, tpsaldo, lote, versao, data"
    )
in
    Etapa_ValidaCamposObrigatorios
```

**Exemplo com nomes personalizados:**
```m
let
    Fonte = Etapa_DefineTipos,

    Etapa_ValidaCamposObrigatorios = fxValidaCamposObrigatorios(
        Fonte,
        "empresa, filial, debito, credito, valor, data",
        "Campos_Com_Problema",
        "Status"
    )
in
    Etapa_ValidaCamposObrigatorios
```

**Comportamento:**
- `null` é tratado como campo obrigatório em branco
- texto vazio é tratado como campo obrigatório em branco
- texto apenas com espaços é tratado como campo obrigatório em branco
- campos preenchidos retornam status `OK`
- coluna obrigatória informada, mas ausente na tabela, será tratada como campo em branco

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
4. Renomeie a consulta com o nome da função, por exemplo `fxValidaCamposObrigatorios`
5. A função estará disponível para uso em outras queries

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
**Última atualização:** Julho 2026
