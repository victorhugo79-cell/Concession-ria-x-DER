# 🚗📊 Sistema de Consolidação de Acidentes Viários

<div align="center">

![Demon Slayer](https://media0.giphy.com/media/v1.Y2lkPTc5MGI3NjExdTdubWk2Zmg4ZWdjendmZ3V4a2hodzNtbHZ3anB1d3gwcWZocXBzZSZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/jh7F7XwHTywg85ekdl/giphy.gif)

**Análise Inteligente de Acidentes por Trecho Rodoviário**

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Pandas-1.3+-green.svg)](https://pandas.pydata.org/)
[![OpenPyXL](https://img.shields.io/badge/OpenPyXL-3.0+-orange.svg)](https://openpyxl.readthedocs.io/)

</div>

---

## 📋 Índice

- [🎯 Sobre o Projeto](#-sobre-o-projeto)
- [✨ Funcionalidades](#-funcionalidades)
- [🛠️ Requisitos](#️-requisitos)
- [📦 Instalação](#-instalação)
- [🚀 Como Usar](#-como-usar)
- [📊 Estrutura dos Dados](#-estrutura-dos-dados)
- [📈 Métricas e Indicadores](#-métricas-e-indicadores)
- [🎨 Formatação](#-formatação)
- [⚠️ Solução de Problemas](#️-solução-de-problemas)
- [📝 Exemplo de Saída](#-exemplo-de-saída)

---

## 🎯 Sobre o Projeto

Este sistema foi desenvolvido para **consolidar e analisar dados de acidentes viários** por trecho rodoviário, gerando relatórios Excel formatados e prontos para análise estratégica.

<div align="center">

![Tanjiro](https://media.giphy.com/media/v1.Y2lkPWVjZjA1ZTQ3MHR3dGtmOG00cnA1OWpzazhkZXcwbnAyZ2N2YWhzODAwNnNvbjA1YyZlcD12MV9naWZzX3JlbGF0ZWQmY3Q9Zw/WkeXjFoubG449UdbGh/giphy.gif)

*"Respiração da Água: Primeira Forma - Análise de Dados!"* 💧

</div>

### 🎭 O que este script faz?

O script processa planilhas Excel com dados de acidentes e gera uma **consolidação inteligente** que:

- 🔍 **Agrupa acidentes por trecho** (Rodovia + KM + Sentido)
- 📊 **Calcula métricas de criticidade** (ICR - Índice de Criticidade de Risco)
- 🕐 **Analisa padrões temporais** (dia/noite, dias da semana)
- 🚗 **Identifica tipos de veículos envolvidos**
- 🌦️ **Correlaciona condições climáticas**
- 📈 **Gera paretos e resumos estatísticos**

---

## ✨ Funcionalidades

### 🎯 Funcionalidades Principais

| Funcionalidade | Descrição | Emoji |
|---------------|-----------|-------|
| **Consolidação por Trecho** | Agrupa acidentes por rodovia, km e sentido | 🗺️ |
| **Cálculo de ICR** | Índice que prioriza trechos mais críticos | 📊 |
| **Análise Temporal** | Distribuição por período (dia/noite) e dias da semana | 🕐 |
| **Detecção de Veículos** | Identifica motos, carros, caminhões e ônibus | 🚗 |
| **Classificação de Acidentes** | Agrupa tipos similares (Choque, Queda, etc.) | 🚨 |
| **Condições Climáticas** | Análise por condições do tempo | 🌦️ |
| **Formatação Automática** | Excel formatado e pronto para uso | ✨ |

### 🎨 Recursos Visuais

<div align="center">

![Nezuko](https://media.giphy.com/media/3o7aD2saQqX1gL5Qy8/giphy.gif)

</div>

- ✅ **Cabeçalhos formatados** com cores e negrito
- ✅ **Largura automática** de colunas
- ✅ **Bordas e alinhamento** profissional
- ✅ **Cabeçalho congelado** para navegação fácil
- ✅ **Múltiplas planilhas** organizadas

---

## 🛠️ Requisitos

### 📦 Dependências Python

```bash
pandas >= 1.3.0
numpy >= 1.20.0
openpyxl >= 3.0.0
```

### 💻 Sistema Operacional

- ✅ Windows 10/11
- ✅ Linux (Ubuntu, Debian, etc.)
- ✅ macOS

### 💾 Espaço em Disco

- Mínimo: **100 MB** livres
- Recomendado: **500 MB** livres
- O script verifica automaticamente o espaço disponível

---

## 📦 Instalação

### 1️⃣ Instalar Python

Certifique-se de ter Python 3.8 ou superior instalado:

```bash
python --version
```

### 2️⃣ Instalar Dependências

```bash
pip install pandas numpy openpyxl
```

Ou usando o arquivo `requirements.txt` (se disponível):

```bash
pip install -r requirements.txt
```

### 3️⃣ Preparar Arquivos

Coloque o arquivo Excel de entrada na mesma pasta do script:

```
📁 Pasta do Projeto/
  ├── 📄 main.py
  ├── 📊 L23_LESTE PAULISTA.xlsx  ← Arquivo de entrada
  └── 📄 README.md
```

---

## 🚀 Como Usar

### 📝 Passo a Passo

<div align="center">

![Zenitsu](https://media.giphy.com/media/v1.Y2lkPWVjZjA1ZTQ3MHR3dGtmOG00cnA1OWpzazhkZXcwbnAyZ2N2YWhzODAwNnNvbjA1YyZlcD12MV9naWZzX3JlbGF0ZWQmY3Q9Zw/WkeXjFoubG449UdbGh/giphy.gif)
*"Respiração do Trovão: Primeira Forma - Execução Rápida!"* ⚡

</div>

#### 1. Preparar o Arquivo de Entrada

O arquivo Excel deve conter uma planilha chamada **`ACIDENTES__MITS`** com as seguintes colunas (ou similares):

| Coluna Necessária | Exemplos de Nomes Aceitos |
|-------------------|---------------------------|
| **Rodovia** | `RODOVIA`, `rodo`, `rodovia` |
| **KM** | `KM`, `km` |
| **Sentido** | `SENTIDO`, `sentido` |
| **Data/Hora** | `DATA HORA`, `data_hora`, `datahora` |
| **Hora** | `HORA`, `hora` |
| **Veículos** | `VEICULOS_ENVOLVIDOS`, `veiculos` |
| **Tipo de Acidente** | `TIPO DO ACIDENTE`, `tipo`, `classificacao` |
| **Condição Climática** | `COND_CLIMATICAS`, `cond`, `climaticas` |

#### 2. Executar o Script

**Windows (PowerShell):**
```powershell
python main.py
```

**Windows (CMD):**
```cmd
python main.py
```

**Linux/macOS:**
```bash
python3 main.py
```

#### 3. Aguardar Processamento

O script irá:

1. ✅ Detectar automaticamente as colunas
2. ✅ Processar e normalizar os dados
3. ✅ Calcular métricas e indicadores
4. ✅ Gerar planilhas formatadas
5. ✅ Aplicar formatação visual

#### 4. Verificar o Resultado

O arquivo **`Trechos_Consolidados.xlsx`** será gerado na mesma pasta com:

- 📊 **trechos_consolidados** - Planilha principal com análises
- 📈 **pareto_veiculos** - Distribuição de veículos
- 📉 **pareto_acidentes** - Distribuição de tipos de acidentes
- 📅 **resumo_periodo** - Resumo por período (dia/noite)
- 📖 **instrucoes** - Documentação das fórmulas
- 📋 **raw_original** - Dados originais (backup)

---

## 📊 Estrutura dos Dados

### 📈 Planilha Principal: `trechos_consolidados`

#### Colunas de Identificação

| Coluna | Descrição | Tipo |
|--------|-----------|------|
| `rodovia` | Nome da rodovia | Texto |
| `km_int` | KM arredondado para baixo | Número |
| `sentido` | Direção do tráfego | Texto |

#### Colunas de Métricas

| Coluna | Descrição | Fórmula/Origem |
|--------|-----------|----------------|
| `sinistros_total` | Total de acidentes no trecho | Contagem |
| `total_vitimas` | Total de vítimas | Soma |
| `total_fatais` | Total de vítimas fatais | Soma |
| `ICR` | Índice de Criticidade de Risco | Ver fórmula abaixo |

#### Colunas de Veículos

| Coluna | Descrição |
|--------|-----------|
| `motos` | Quantidade de acidentes com motos |
| `carros` | Quantidade de acidentes com carros |
| `caminhoes` | Quantidade de acidentes com caminhões |
| `onibus` | Quantidade de acidentes com ônibus |

#### Colunas Temporais

| Coluna | Descrição |
|--------|-----------|
| `dia` | Acidentes ocorridos entre 06:00 e 17:59 |
| `noite` | Acidentes ocorridos entre 18:00 e 05:59 |
| `unknown` | Acidentes sem hora válida |
| `Segunda` a `Domingo` | Contagem por dia da semana |
| `mes_ano` | Mês/Ano mais frequente (formato MM/AAAA) |

#### Colunas de Tipos de Acidentes

Colunas dinâmicas criadas automaticamente baseadas nos tipos encontrados:

- `Choque` - Todos os tipos de choque/colisão
- `Colisão` - Colisões específicas
- `Queda` - Quedas e tombamentos
- `Atropelamento` - Atropelamentos (exceto animais)
- `Atropelamento de Animal` - Atropelamentos de animais
- `Tombamento`, `Capotamento`, `Saída de Pista`, etc.

#### Colunas Climáticas

| Coluna | Descrição |
|--------|-----------|
| `cond_boa` | Acidentes em condições climáticas boas |
| `cond_chuva` | Acidentes durante chuva |
| `cond_neblina` | Acidentes com neblina |
| `cond_nublado` | Acidentes com céu nublado |

---

## 📈 Métricas e Indicadores

### 🎯 ICR - Índice de Criticidade de Risco

O **ICR** é o indicador principal para priorização de ações de segurança viária.

<div align="center">

![Inosuke](https://media.giphy.com/media/3o7aD2saQqX1gL5Qy8/giphy.gif)

*"Respiração da Fera: Análise de Criticidade!"* 🐗

</div>

#### 📐 Fórmula do ICR

```
ICR = 10 × total_fatais 
    + 3 × (total_vitimas - total_fatais)
    + 1 × (sinistros_total - total_vitimas)
```

#### 🎯 Interpretação

| Valor ICR | Criticidade | Ação Recomendada |
|-----------|-------------|------------------|
| **0-10** | 🟢 Baixa | Monitoramento |
| **11-50** | 🟡 Média | Análise detalhada |
| **51-100** | 🟠 Alta | Ações preventivas |
| **>100** | 🔴 Crítica | Intervenção imediata |

#### 📊 Ordenação

A planilha é **automaticamente ordenada por ICR decrescente**, colocando os trechos mais críticos no topo.

---

### 🕐 Análise Temporal

#### Período do Dia

- **Dia**: 06:00 às 17:59
- **Noite**: 18:00 às 05:59
- **Unknown**: Hora não disponível ou inválida

#### Dias da Semana

A contagem é feita para cada dia:
- Segunda-feira
- Terça-feira
- Quarta-feira
- Quinta-feira
- Sexta-feira
- Sábado
- Domingo

---

### 🚗 Classificação de Veículos

O sistema detecta automaticamente veículos mencionados nos registros:

| Veículo | Palavras-chave Detectadas |
|---------|---------------------------|
| 🏍️ **Moto** | moto, motocicleta, motoc |
| 🚗 **Carro** | carro, auto, automóvel, passageiro |
| 🚛 **Caminhão** | caminhão, carreta, truck, tracto |
| 🚌 **Ônibus** | ônibus, bus |

---

### 🚨 Classificação de Acidentes

O sistema agrupa tipos similares de acidentes:

| Grupo | Tipos Agrupados |
|-------|-----------------|
| **Choque** | choque, batida, embate, impacto |
| **Colisão** | colisão, colisao |
| **Queda** | queda, caiu, caída |
| **Atropelamento** | atropelamento, atropelou, atropelo |
| **Atropelamento de Animal** | atropelamento de animal (separado) |
| **Tombamento** | tombamento, tombou, virou |
| **Capotamento** | capotamento, capotou |
| **Saída de Pista** | saída de pista, saiu da pista |
| **Incêndio** | incêndio, queimou |
| **Abalroamento** | abalroamento, abalroou |

---

## 🎨 Formatação

### ✨ Recursos Visuais Aplicados

<div align="center">

![Kanao](https://media.giphy.com/media/3o7aD2saQqX1gL5Qy8/giphy.gif)

*"Formatação Perfeita como a Respiração da Flor!"* 🌸

</div>

#### 🎨 Cabeçalho

- **Cor de fundo**: Azul escuro (#366092)
- **Texto**: Branco, negrito, tamanho 11
- **Alinhamento**: Centralizado
- **Bordas**: Todas as células

#### 📏 Colunas

- **Largura automática**: Ajustada ao conteúdo
- **Largura mínima**: 10 caracteres
- **Largura máxima**: 50 caracteres

#### 📐 Células

- **Bordas**: Todas as células têm bordas
- **Números**: Alinhados à direita
- **Texto**: Alinhado à esquerda
- **Quebra de texto**: Automática

#### 🔒 Navegação

- **Cabeçalho congelado**: Primeira linha sempre visível
- **Rolagem facilitada**: Dados sempre acessíveis

---

## ⚠️ Solução de Problemas

### 🚨 Erros Comuns

#### 1. Arquivo não encontrado

```
FileNotFoundError: Arquivo não encontrado: L23_LESTE PAULISTA.xlsx
```

**Solução:**
- ✅ Verifique se o arquivo está na mesma pasta do script
- ✅ Confirme o nome exato do arquivo (case-sensitive)
- ✅ Verifique se o arquivo não está aberto em outro programa

#### 2. Sem espaço em disco

```
ERRO CRÍTICO: Sem espaço em disco suficiente!
```

**Solução:**
- ✅ Libere espaço no disco (mínimo 100-200 MB)
- ✅ Exclua arquivos temporários da pasta Temp
- ✅ Execute o script em outro local com mais espaço
- ✅ Limpe a pasta de destino

#### 3. Arquivo aberto no Excel

```
PermissionError: O arquivo está aberto em outro programa
```

**Solução:**
- ✅ Feche o arquivo Excel se estiver aberto
- ✅ Feche outros programas que possam estar usando o arquivo
- ✅ Execute o script novamente

#### 4. Colunas não detectadas

```
RuntimeError: Não foi possível localizar coluna 'rodovia'
```

**Solução:**
- ✅ Verifique se a planilha se chama `ACIDENTES__MITS`
- ✅ Confirme que as colunas existem na planilha
- ✅ Verifique se os nomes das colunas estão corretos

#### 5. Erro de formatação

Se houver erro ao aplicar formatação, o script continuará normalmente, mas o arquivo pode não estar formatado.

**Solução:**
- ✅ Verifique se o openpyxl está atualizado: `pip install --upgrade openpyxl`
- ✅ O arquivo será gerado mesmo sem formatação

---

### 🔧 Dicas de Performance

#### ⚡ Otimização

- **Arquivos grandes**: O processamento pode levar alguns minutos
- **Memória**: Certifique-se de ter RAM suficiente (mínimo 4GB)
- **Processador**: Processamento é single-threaded

#### 💾 Economia de Espaço

O script automaticamente:
- ✅ Remove planilhas grandes se não houver espaço
- ✅ Verifica espaço antes de salvar
- ✅ Avisa sobre espaço insuficiente

---

## 📝 Exemplo de Saída

### 📊 Estrutura do Arquivo Gerado

```
Trechos_Consolidados.xlsx
│
├── 📊 trechos_consolidados (Planilha Principal)
│   ├── Colunas de identificação (rodovia, km_int, sentido)
│   ├── Métricas (sinistros_total, total_vitimas, ICR)
│   ├── Veículos (motos, carros, caminhoes, onibus)
│   ├── Períodos (dia, noite, unknown)
│   ├── Dias da semana (Segunda a Domingo)
│   ├── Tipos de acidentes (Choque, Colisão, etc.)
│   └── Condições climáticas
│
├── 📈 pareto_veiculos
│   └── Distribuição de veículos envolvidos
│
├── 📉 pareto_acidentes
│   └── Distribuição de tipos de acidentes
│
├── 📅 resumo_periodo
│   └── Resumo por período (dia/noite/unknown)
│
├── 📖 instrucoes
│   └── Documentação das fórmulas e conceitos
│
└── 📋 raw_original
    └── Dados originais (backup)
```

### 📈 Exemplo de Dados

| rodovia | km_int | sentido | sinistros_total | ICR | dia | noite | Segunda | Terça | ... |
|---------|--------|---------|-----------------|-----|-----|-------|---------|-------|-----|
| SP070 | 23 | OESTE | 972 | 3456 | 600 | 372 | 150 | 120 | ... |
| SP070 | 24 | OESTE | 450 | 1890 | 280 | 170 | 80 | 65 | ... |

---

## 🎓 Conceitos Importantes

### 📐 KM Int (KM Inteiro)

O **KM Int** é o quilômetro arredondado **para baixo**:

- `26.1` → `26`
- `26.9` → `26`
- `27.0` → `27`

Isso agrupa acidentes próximos no mesmo trecho.

### 🕐 Periodização

- **Dia**: 06:00 às 17:59 (12 horas)
- **Noite**: 18:00 às 05:59 (12 horas)
- **Unknown**: Sem hora válida

### 🚗 Detecção de Veículos

A detecção é feita por **heurística textual**, procurando palavras-chave nos registros. Pode haver falsos positivos ou negativos dependendo da qualidade dos dados.

### 🚨 Agrupamento de Acidentes

Tipos similares são agrupados automaticamente. Por exemplo:
- "Choque frontal" → `Choque`
- "Colisão traseira" → `Colisão`
- "Atropelamento de animal" → `Atropelamento de Animal` (separado)

---

## 📞 Suporte

<div align="center">

![Hashira](https://media.giphy.com/media/3o7aD2saQqX1gL5Qy8/giphy.gif)

*"Respiração da Água: Décima Primeira Forma - Suporte Técnico!"* 💧

</div>

### 🆘 Precisa de Ajuda?

1. ✅ Verifique a seção [Solução de Problemas](#️-solução-de-problemas)
2. ✅ Confirme que todos os requisitos estão instalados
3. ✅ Verifique os logs de erro no console
4. ✅ Certifique-se de que o arquivo de entrada está correto

---

## 📜 Licença

Este projeto é de uso interno para análise de dados de segurança viária.

---

## 🙏 Agradecimentos

<div align="center">

![Demon Slayer Corps](https://media.giphy.com/media/3o7aD2saQqX1gL5Qy8/giphy.gif)

*"Que todos os dados sejam processados com sucesso!"* ⚔️

**Desenvolvido com dedicação para melhorar a segurança viária** 🛡️

</div>

---

<div align="center">

**Made with ❤️ and 🐍 Python**

*"Respiração Total: Concentração Constante!"* 🌊

</div>




