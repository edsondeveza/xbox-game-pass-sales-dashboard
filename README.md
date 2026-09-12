# Dashboard Xbox Game Pass — Análise de Assinaturas

## Visão geral

Dashboard interativo em Excel para analisar assinaturas Xbox Game Pass, passes adicionais, renovação automática, planos e cupons.

### Recursos principais

- 4 cartões: Receita, Assinantes, Cupons e Renovação anual.
- 7 gráficos analíticos.
- Filtros por ano, semestre ou trimestre.
- Segmentações por Plano, Renovação, Periodicidade, EA Play e Minecraft.
- Comparação com o período anterior equivalente.
- Tratamento de combinações sem dados.
- Identidade visual baseada na marca Xbox.

---

## Perguntas de negócio

1. Qual é o faturamento dos planos anuais?
2. Como a receita anual se divide por renovação automática?
3. Qual é a quantidade de adesões e o faturamento do EA Play?
4. Qual é a quantidade de adesões e o faturamento do Minecraft Season Pass?
5. Quais planos geram mais assinantes, receita e ticket médio?
6. Como evoluíram a receita e as novas assinaturas durante 2024?
7. Qual é o impacto dos cupons no valor bruto?

---

## Estrutura da pasta de trabalho

| Planilha | Finalidade |
| --- | --- |
| `Instruções` | Página inicial com orientações de uso. |
| `D̳ashboard` | Cartões, filtros e gráficos. |
| `B̳ases` | Base transacional armazenada em `Tabela1`. |
| `C̳álculos` | Perguntas de negócio e análises auxiliares. |
| `Dashboard Engine` | Fontes calculadas dos cartões e gráficos. |
| `Pivot Engine` | Tabela Dinâmica mestre `PT_Master`. |
| `A̳ssets` | Cores, logos e elementos visuais. |

As planilhas técnicas podem permanecer ocultas.

---

## Dados utilizados

A fonte principal é `Tabela1`, localizada em `B̳ases`.

- 295 registros de assinantes.
- Período de janeiro a dezembro de 2024.
- Planos Core, Standard e Ultimate.
- Periodicidades Monthly, Quarterly e Annual.

### Dicionário de dados

| Campo | Descrição |
| --- | --- |
| `Subscriber ID` | Identificador do assinante. |
| `Name` | Nome do assinante. |
| `Plan` | Plano Core, Standard ou Ultimate. |
| `Start Date` | Data de início da assinatura. |
| `Auto Renewal` | Indica renovação automática. |
| `Subscription Price` | Preço base da assinatura. |
| `Subscription Type` | Periodicidade da assinatura. |
| `EA Play Season Pass` | Indica adesão ao EA Play. |
| `EA Play Season Pass Price` | Valor do EA Play. |
| `Minecraft Season Pass` | Indica adesão ao Minecraft. |
| `Minecraft Season Pass Price` | Valor do Minecraft. |
| `Coupon Value` | Desconto concedido. |
| `Total Value` | Receita líquida da transação. |
| `Month Start` | Primeiro dia do mês. |
| `Year` | Ano da assinatura. |

### Campos auxiliares

```excel
=DATE(YEAR([@[Start Date]]),MONTH([@[Start Date]]),1)
```

Cria o primeiro dia do mês.

```excel
=YEAR([@[Start Date]])
```

Extrai o ano da assinatura.

---

## Arquitetura técnica

### Tabela estruturada

A base utiliza uma tabela chamada `Tabela1`, que permite:

- Expansão automática.
- Referências estruturadas.
- Propagação de fórmulas.
- Atualização da Tabela Dinâmica.

### Pivô mestre

A Tabela Dinâmica `PT_Master`, em `Pivot Engine`, concentra as dimensões:

- Month Start.
- Plan.
- Auto Renewal.
- Subscription Type.
- EA Play Season Pass.
- Minecraft Season Pass.

Medidas:

- Soma de Total Value.
- Contagem de Subscriber ID.
- Soma de Coupon Value.
- Soma de EA Play Season Pass Price.
- Soma de Minecraft Season Pass Price.

### Segmentações

O dashboard possui segmentações para:

- Plano.
- Renovação automática.
- Periodicidade.
- EA Play.
- Minecraft.

As segmentações filtram o pivô mestre. O Dashboard Engine agrega os resultados filtrados e aplica o período selecionado.

### Períodos de análise

O usuário pode escolher:

- Ano completo.
- 1º semestre.
- 2º semestre.
- 1º trimestre.
- 2º trimestre.
- 3º trimestre.
- 4º trimestre.

Cada opção possui:

- Data inicial.
- Data final exclusiva.
- Início do período anterior.
- Final do período anterior.

A comparação é realizada contra um período anterior de mesma duração.

---

## Fórmulas principais

***Receita total**

```excel
=SUM(Tabela1[Total Value])
```

***Quantidade de assinantes**

```excel
=ROWS(Tabela1[Subscriber ID])
```

***Cupons concedidos**

```excel
=SUM(Tabela1[Coupon Value])
```

***Renovação anual**

```excel
=SUMIFS(
    Tabela1[Total Value],
    Tabela1[Subscription Type],"Annual",
    Tabela1[Auto Renewal],"Yes"
)
/SUMIFS(
    Tabela1[Total Value],
    Tabela1[Subscription Type],"Annual"
)
```

***Agregação por período**

```excel
=SUMIFS(
    'Pivot Engine'!$G:$G,
    'Pivot Engine'!$A:$A,">="&DataInicial,
    'Pivot Engine'!$A:$A,"<"&DataFinal
)
```

***Variação contra o período anterior**

```excel
=IF(
    OR(Anterior="",Anterior=0),
    "",
    Atual/Anterior-1
)
```

Formato: `+0,0%;-0,0%;0,0%`

***Tratamento de ausência de dados**

```excel
=IF(SUM(Intervalo)=0,"SEM DADOS",Resultado)
```

***Caixas de texto dinâmicas**

O Excel não calcula fórmulas completas diretamente em caixas de texto. A fórmula deve permanecer em uma célula e a caixa deve ser vinculada a ela:

```excel
=Dashboard!D15
```

***Formatação monetária**

Para evitar a exibição incorreta `R$`, utilize:

```
[$R$-pt-BR] #.##0,00
```

Sem centavos:

```
[$R$-pt-BR] #.##0
```

Não combine o símbolo regional automático com outro texto `R$`.

---

## Gráficos

| Gráfico | Tipo | Objetivo |
| --- | --- | --- |
| Faturamento anual | Barras | Receita dos planos anuais. |
| Renovação automática | Rosca | Receita anual por renovação. |
| EA Play | Barras | Receita filtrada do passe. |
| Minecraft | Barras | Receita filtrada do passe. |
| Desempenho por plano | Colunas | Comparar os planos. |
| Evolução mensal | Linha | Mostrar a receita mensal. |
| Impacto dos cupons | Colunas | Comparar cupons, líquido e bruto. |

---

## Identidade visual

| Cor | Utilização |
| --- | --- |
| `#107C10` | Verde principal Xbox. |
| `#9BC848` | Destaques. |
| `#22C55E` | Acentos positivos e EA Play. |
| `#1F1F1F` | Cabeçalhos escuros. |
| `#E8E6E9` | Elementos neutros. |
| `#F2F2F2` | Fundo. |

Padrões:

- Fonte Aptos.
- Cartões brancos.
- Cabeçalhos escuros.
- Barra lateral verde.
- Formatação monetária brasileira.

---

## Instruções de utilização

1. Abra a planilha `Dashboard`.
2. Escolha o período na barra lateral.
3. Aplique uma ou mais segmentações.
4. Observe a atualização dos cartões.
5. Analise os gráficos.
6. Compare com o período anterior.
7. Limpe os filtros para retornar à visão completa.

### Recomendações

- Altere um filtro por vez para compreender seu impacto.
- Combine filtros para análises específicas.
- Se aparecer `SEM DADOS`, remova um filtro ou amplie o período.
- Limpe os filtros antes de iniciar uma nova análise.

---

## Reprodução do projeto

### 1. Preparar a base

- Crie uma tabela chamada `Tabela1`.
- Utilize os campos do dicionário de dados.
- Garanta que datas e moedas sejam valores numéricos reais.
- Adicione `Month Start` e `Year`.

### 2. Criar o pivô mestre

- Insira uma Tabela Dinâmica baseada em `Tabela1`.
- Adicione as dimensões às linhas.
- Adicione as medidas aos valores.
- Utilize layout tabular.
- Remova subtotais.
- Nomeie o pivô como `PT_Master`.

### 3. Criar as segmentações

- Crie as cinco segmentações categóricas.
- Conecte-as ao pivô mestre.
- Posicione-as na lateral esquerda.

### 4. Criar períodos rápidos

- Crie uma tabela de parâmetros.
- Informe início e fim de cada período.
- Informe o período anterior equivalente.
- Crie uma lista suspensa com as opções.

### 5. Criar o motor do dashboard

- Agregue os resultados do pivô com `SUMIFS`.
- Calcule os indicadores atuais.
- Calcule os indicadores anteriores.
- Calcule as variações.
- Crie as fontes dos gráficos.
- Trate resultados vazios e divisões por zero.

### 6. Criar o dashboard

- Crie os quatro cartões.
- Insira os sete gráficos.
- Vincule caixas de texto às células.
- Aplique a paleta Xbox.
- Formate rótulos, eixos e legendas.

### 7. Validar

Teste:

- Ano completo.
- Primeiro e segundo semestres.
- Quatro trimestres.
- Cada segmentação individualmente.
- Combinações de filtros.
- Combinações sem dados.
- Limpeza completa dos filtros.

---

## Controles de qualidade

Antes da entrega, confirme:

- Ausência de `#REF!`.
- Ausência de `#DIV/0!`.
- Ausência de `#NOME?`.
- Ausência de `#N/D`.
- Total anual de 295 assinantes.
- Receita total de R$ 7.633,00.
- Cupons totais de R$ 2.122,00.
- Sete gráficos conectados.
- Segmentações sem sobreposição.
- Cartões atualizados pelos filtros.
- Rótulos exibindo somente R$.

---

## Manutenção

Ao adicionar registros:

- Insira-os ao final de `Tabela1`.
- Confirme a propagação das colunas calculadas.
- Atualize a Tabela Dinâmica.
- Verifique novos meses e anos.
- Atualize a tabela de períodos rápidos, se necessário.
- Ao adicionar um novo ano, atualize os limites dos períodos e valide as comparações.

---

## Resultados de referência

Na visão anual original:

- Receita total: R$ 7.633,00
- Assinantes: 295
- Cupons: R$ 2.122,00
- Renovação anual: 87,6%
- EA Play: 98 adesões / R$ 2.940,00
- Minecraft: 194 adesões / R$ 3.880,00
- Plano com maior receita: Ultimate / R$ 5.388,00

## Contexto acadêmico

Projeto desenvolvido como desafio do bootcamp **Santander — Excel com IA**, oferecido pela **DIO**, com foco em organização de dados, automação sem macros e experiência do usuário no Excel.
