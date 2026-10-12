import { useEffect, useMemo, useRef, useState } from "react";
import { AnimatePresence, motion } from "framer-motion";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import html2canvas from "html2canvas";
import jsPDF from "jspdf";
import {
  ArcElement,
  BarElement,
  CategoryScale,
  Chart,
  Legend,
  LinearScale,
  LineElement,
  PointElement,
  Tooltip,
} from "chart.js";
import { Bar, Doughnut, Line, Scatter } from "react-chartjs-2";
import {
  BookOpen,
  Brain,
  CheckCircle2,
  Download,
  FileChartColumn,
  Filter,
  GraduationCap,
  Home,
  Import,
  LayoutTemplate,
  Moon,
  Settings,
  Target,
  Trophy,
  Sun,
  Upload,
  X,
} from "lucide-react";

Chart.register(CategoryScale, LinearScale, BarElement, LineElement, PointElement, ArcElement, Tooltip, Legend);

type DataRow = Record<string, string | number>;
type ThemeMode = "dark" | "light";
type TabKey =
  | "dashboard"
  | "importar"
  | "editor"
  | "filtros"
  | "metricas"
  | "templates"
  | "guia"
  | "config";

type WidgetType =
  | "barVertical"
  | "barHorizontal"
  | "line"
  | "donut"
  | "area"
  | "kpi"
  | "table"
  | "scatter"
  | "funnel"
  | "gauge"
  | "heatmap"
  | "combo"
  | "text"
  | "shapeRect"
  | "shapeLine"
  | "image";

type Widget = {
  id: string;
  type: WidgetType;
  x: number;
  y: number;
  w: number;
  h: number;
  title: string;
  color: string;
  xField?: string;
  yField?: string;
  valueField?: string;
  legendField?: string;
  text?: string;
  imageUrl?: string;
};

type ReportPage = {
  id: string;
  name: string;
  background: string;
  widgets: Widget[];
};

type MeasureType =
  | "SOMA"
  | "MEDIA"
  | "CONTAGEM"
  | "MINIMO"
  | "MAXIMO"
  | "PORCENTAGEM"
  | "VARIACAO"
  | "ACUMULADO"
  | "YOY";

type Measure = {
  id: string;
  nome: string;
  funcao: MeasureType;
  campo: string;
  formula: string;
};

type CourseLesson = {
  id: string;
  titulo: string;
  explicacao: string;
  passos: string[];
};

type CourseTask = {
  id: string;
  titulo: string;
  instrucoes: string;
  pontos: number;
  dica: string;
  tipo?: "escrita" | "pratica";
  criterio?: string[];
};

type CourseQuiz = {
  q: string;
  a: string;
  opcoes: string[];
  explicacao?: string;
};

type CourseModule = {
  id: string;
  titulo: string;
  descricao: string;
  nivel: "Iniciante" | "Intermediario" | "Avancado";
  duracao: string;
  aulas: CourseLesson[];
  tarefas: CourseTask[];
  quiz: CourseQuiz[];
};

const STORAGE_KEY = "pbc_state_v1";
const USER_KEY = "pbc_user_name";

const tabs: { key: TabKey; label: string; icon: typeof Home }[] = [
  { key: "dashboard", label: "Dashboard", icon: Home },
  { key: "importar", label: "Importar Dados", icon: Import },
  { key: "editor", label: "Editor de Relatorios", icon: FileChartColumn },
  { key: "filtros", label: "Filtros", icon: Filter },
  { key: "metricas", label: "Metricas", icon: Brain },
  { key: "templates", label: "Templates", icon: LayoutTemplate },
  { key: "guia", label: "Guia de Aprendizado", icon: BookOpen },
  { key: "config", label: "Configuracoes", icon: Settings },
];

const templates = [
  "Dashboard de Vendas",
  "Dashboard Financeiro",
  "Dashboard de RH",
  "Dashboard de Marketing",
  "Dashboard de Estoque",
  "Dashboard de Atendimento",
  "Dashboard de Producao",
  "Dashboard Executivo",
];

const learningModules: CourseModule[] = [
  {
    id: "m1",
    titulo: "Modulo 1: Introducao ao Power BI",
    descricao: "Entenda o que e BI, por que o Power BI e importante e como a ferramenta e usada nas empresas.",
    nivel: "Iniciante",
    duracao: "50 min",
    aulas: [
      {
        id: "a1",
        titulo: "BI sem complicacao",
        explicacao: "Business Intelligence e o processo de transformar dados em decisao. No dia a dia, isso significa olhar para numeros de vendas, custos e clientes para decidir com seguranca.",
        passos: ["Defina o problema de negocio", "Escolha os dados certos", "Transforme dados em visualizacoes objetivas"],
      },
      {
        id: "a2",
        titulo: "Ecossistema Power BI",
        explicacao: "Power BI Desktop cria relatorios, Service publica e compartilha, Mobile acompanha em qualquer lugar.",
        passos: ["Desktop: modelagem e DAX", "Service: publicacao e governanca", "Mobile: consulta rapida e alertas"],
      },
    ],
    tarefas: [
      { id: "t1", titulo: "Mapa mental BI", instrucoes: "Escreva em 5 linhas como BI ajudaria um negocio local da sua cidade.", pontos: 30, dica: "Pense em vendas, clientes e estoque." },
      { id: "t2", titulo: "Comparar Desktop x Service", instrucoes: "Crie uma tabela com 3 diferencas entre Desktop e Service.", pontos: 20, dica: "Foque em objetivo de cada ferramenta." },
    ],
    quiz: [
      { q: "Power BI e uma ferramenta de:", a: "Business Intelligence", opcoes: ["Business Intelligence", "Edicao de video", "Banco de dados"] },
      { q: "Onde voce cria relatorios localmente?", a: "Power BI Desktop", opcoes: ["Power BI Desktop", "Power BI Mobile", "Excel Online"] },
    ],
  },
  {
    id: "m2",
    titulo: "Modulo 2: Primeiros Passos",
    descricao: "Aprenda o caminho inicial: instalar, importar dados e criar seu primeiro relatorio guiado.",
    nivel: "Iniciante",
    duracao: "65 min",
    aulas: [
      { id: "a1", titulo: "Conhecendo a interface", explicacao: "Ribbon com comandos, Canvas para montar visuals e paineis para campos/filtros.", passos: ["Abra o Desktop", "Explore a guia Home", "Identifique Campos, Visualizacoes e Filtros"] },
      { id: "a2", titulo: "Primeira importacao", explicacao: "Voce pode importar CSV, Excel e bases SQL. Nesta plataforma use a aba Importar Dados para praticar.", passos: ["Carregue dados exemplo", "Revise tipos", "Abra no Editor"] },
    ],
    tarefas: [
      { id: "t1", titulo: "Primeiro grafico", instrucoes: "Crie um grafico de barras com Categoria no eixo X e Valor_Total no eixo Y.", pontos: 35, dica: "Use o Editor de Relatorios." },
      { id: "t2", titulo: "Titulo profissional", instrucoes: "Renomeie o visual para 'Receita por Categoria'.", pontos: 15, dica: "Use o painel de propriedades." },
    ],
    quiz: [
      { q: "Qual menu importa dados no Power BI Desktop?", a: "Obter Dados", opcoes: ["Obter Dados", "Exibir", "Modelagem"] },
      { q: "Para que serve o Canvas?", a: "Montar visualizacoes", opcoes: ["Montar visualizacoes", "Editar SQL", "Instalar plugins"] },
    ],
  },
  {
    id: "m3",
    titulo: "Modulo 3: Transformacao de Dados (Power Query)",
    descricao: "Limpe e transforme dados para evitar erros nos indicadores.",
    nivel: "Iniciante",
    duracao: "75 min",
    aulas: [
      { id: "a1", titulo: "Limpeza essencial", explicacao: "Remover duplicatas e corrigir tipos e o primeiro passo para analises confiaveis.", passos: ["Verifique nulos", "Remova duplicatas", "Padronize tipos"] },
      { id: "a2", titulo: "Merge, Append e Unpivot", explicacao: "Merge junta tabelas, Append empilha, Unpivot normaliza colunas para analise temporal.", passos: ["Escolha chave correta", "Teste o resultado", "Valide total de linhas"] },
    ],
    tarefas: [
      { id: "t1", titulo: "Faxina dos dados", instrucoes: "Na aba de Importar, exclua 1 coluna irrelevante e renomeie outra coluna.", pontos: 25, dica: "Exemplo: renomeie Valor_Total para Receita." },
      { id: "t2", titulo: "Filtro antes de analisar", instrucoes: "Aplique filtro por Regiao para analisar apenas Sudeste.", pontos: 20, dica: "Use filtros globais na aba Filtros." },
    ],
    quiz: [{ q: "Merge une tabelas usando:", a: "Chaves em comum", opcoes: ["Chaves em comum", "Cor das colunas", "Ordem alfabetica"] }],
  },
  {
    id: "m4",
    titulo: "Modulo 4: Modelagem de Dados",
    descricao: "Monte modelos claros com fato e dimensao para acelerar analises.",
    nivel: "Intermediario",
    duracao: "70 min",
    aulas: [
      { id: "a1", titulo: "Fato x Dimensao", explicacao: "Tabela fato guarda eventos (vendas), dimensoes descrevem contexto (produto, regiao, data).", passos: ["Separe entidades", "Evite colunas repetidas", "Conecte com chaves"] },
      { id: "a2", titulo: "Cardinalidade e filtros", explicacao: "Entender 1:N evita duplicidades e calculos errados.", passos: ["Defina relacoes", "Teste filtros", "Revise direcao"] },
    ],
    tarefas: [
      { id: "t1", titulo: "Desenho estrela", instrucoes: "Descreva em texto como ficaria um modelo estrela para TechStore Brasil.", pontos: 30, dica: "Fato vendas no centro e dimensoes ao redor." },
      { id: "t2", titulo: "Valide contexto", instrucoes: "Explique por que filtros de dimensao devem atingir tabela fato.", pontos: 20, dica: "Pense no impacto em medidas." },
    ],
    quiz: [{ q: "Modelo mais recomendado para BI:", a: "Esquema estrela", opcoes: ["Esquema estrela", "Modelo circular", "Tabela unica gigante"] }],
  },
  {
    id: "m5",
    titulo: "Modulo 5: DAX",
    descricao: "Crie medidas robustas para responder perguntas do negocio com agilidade.",
    nivel: "Intermediario",
    duracao: "95 min",
    aulas: [
      { id: "a1", titulo: "Contexto de filtro e linha", explicacao: "DAX depende do contexto. O mesmo calculo muda quando voce aplica filtros.", passos: ["Compare medida com e sem filtro", "Teste por categoria", "Observe variacao"] },
      { id: "a2", titulo: "Funcoes que mais caem no trabalho", explicacao: "SUM, CALCULATE, DIVIDE, SAMEPERIODLASTYEAR e TOTALYTD resolvem boa parte dos cenarios reais.", passos: ["Comece por soma base", "Derive margem", "Crie comparacao anual"] },
    ],
    tarefas: [
      { id: "t1", titulo: "Medida de margem", instrucoes: "Crie uma medida de variacao percentual na aba Metricas.", pontos: 35, dica: "Use funcao VARIACAO e campo Lucro." },
      { id: "t2", titulo: "Acumulado", instrucoes: "Crie uma medida de ACUMULADO para Valor_Total.", pontos: 25, dica: "Compare com SOMA para entender diferenca conceitual." },
    ],
    quiz: [
      { q: "Funcao para acumulado no ano:", a: "TOTALYTD", opcoes: ["TOTALYTD", "GROUPBY", "ROUNDUP"] },
      { q: "Comparacao com ano anterior:", a: "SAMEPERIODLASTYEAR", opcoes: ["SAMEPERIODLASTYEAR", "SUMMARIZE", "DATEVALUE"] },
    ],
  },
  {
    id: "m6",
    titulo: "Modulo 6: Visualizacoes",
    descricao: "Escolha o grafico certo e comunique a historia dos dados com clareza.",
    nivel: "Intermediario",
    duracao: "80 min",
    aulas: [
      { id: "a1", titulo: "Escolha do visual", explicacao: "Linha para tendencia, barras para comparacao, dispersao para correlacao.", passos: ["Defina a pergunta", "Escolha visual", "Ajuste titulo e cores"] },
      { id: "a2", titulo: "Interacao e filtros", explicacao: "Com cross-filtering, um clique em categoria impacta o restante do dashboard.", passos: ["Clique no grafico", "Veja outros visuais mudarem", "Limpe o filtro"] },
    ],
    tarefas: [
      { id: "t1", titulo: "Dashboard com 3 visuais", instrucoes: "Adicione barra, linha e tabela na mesma pagina.", pontos: 30, dica: "Use os tipos no painel esquerdo." },
      { id: "t2", titulo: "Padrao visual", instrucoes: "Aplique mesma paleta de cor em todos os visuais.", pontos: 20, dica: "Edite cor no painel de propriedades." },
    ],
    quiz: [{ q: "Melhor grafico para tendencia temporal:", a: "Linha", opcoes: ["Linha", "Pizza", "Gauge"] }],
  },
  {
    id: "m7",
    titulo: "Modulo 7: Publicacao e Compartilhamento",
    descricao: "Leve seus relatorios para o time com seguranca e atualizacao constante.",
    nivel: "Intermediario",
    duracao: "60 min",
    aulas: [
      { id: "a1", titulo: "Service e Workspaces", explicacao: "No Service, voce organiza conteudo por times e controla permissoes.", passos: ["Criar workspace", "Publicar relatorio", "Definir acesso"] },
      { id: "a2", titulo: "RLS e gateway", explicacao: "RLS limita dados por usuario e gateway conecta dados locais com nuvem.", passos: ["Criar regra de perfil", "Testar como usuario", "Configurar gateway"] },
    ],
    tarefas: [
      { id: "t1", titulo: "Plano de compartilhamento", instrucoes: "Liste quem pode ver, editar e compartilhar um relatorio executivo.", pontos: 20, dica: "Pense em perfis: diretoria, gerencia, analista." },
      { id: "t2", titulo: "Checklist de publicacao", instrucoes: "Monte checklist com 5 itens antes de publicar.", pontos: 20, dica: "Inclua validação de filtros e medidas." },
    ],
    quiz: [{ q: "RLS significa:", a: "Row Level Security", opcoes: ["Row Level Security", "Report Level Style", "Raw Logic System"] }],
  },
  {
    id: "m8",
    titulo: "Modulo 8: Dicas Avancadas e Boas Praticas",
    descricao: "Aprimore performance, design e narrativa para dashboards de alto nivel.",
    nivel: "Avancado",
    duracao: "85 min",
    aulas: [
      { id: "a1", titulo: "Performance", explicacao: "Modelos enxutos, boas relacoes e DAX eficiente melhoram tempo de resposta.", passos: ["Remova colunas sem uso", "Evite medidas redundantes", "Teste desempenho"] },
      { id: "a2", titulo: "Storytelling com dados", explicacao: "Todo dashboard deve contar uma historia objetiva para tomada de decisao.", passos: ["Defina mensagem central", "Organize hierarquia visual", "Finalize com conclusao"] },
    ],
    tarefas: [
      { id: "t1", titulo: "Otimizar pagina", instrucoes: "Reduza em 20% a quantidade de elementos de uma pagina do editor.", pontos: 30, dica: "Menos ruído, mais foco." },
      { id: "t2", titulo: "Resumo executivo", instrucoes: "Escreva um texto de 4 linhas com principal insight do seu dashboard.", pontos: 20, dica: "Comece pelo impacto no negocio." },
    ],
    quiz: [{ q: "Modo que costuma ter melhor performance:", a: "Import", opcoes: ["Import", "DirectQuery", "Streaming"] }],
  },
  {
    id: "m9",
    titulo: "Modulo 9: Projetos Praticos",
    descricao: "Consolide tudo com projetos completos de Vendas, Financeiro e RH.",
    nivel: "Avancado",
    duracao: "120 min",
    aulas: [
      { id: "a1", titulo: "Projeto 1: Vendas", explicacao: "Construa um painel com metas, receita, lucro e desempenho por regiao.", passos: ["Importe dados", "Modele medidas", "Monte tela executiva"] },
      { id: "a2", titulo: "Projeto 2 e 3", explicacao: "Repita metodo para Financeiro e RH, ajustando KPIs e filtros.", passos: ["Escolha KPIs", "Monte visuais", "Valide com checklist"] },
    ],
    tarefas: [
      { id: "t1", titulo: "Projeto final", instrucoes: "Crie um dashboard executivo com no minimo 4 visuais e 2 filtros ativos.", pontos: 60, dica: "Pode usar templates como base." },
      { id: "t2", titulo: "Apresentacao final", instrucoes: "Escreva uma conclusao com 3 recomendacoes de negocio baseadas no dashboard.", pontos: 40, dica: "Fale de ação, risco e oportunidade." },
    ],
    quiz: [{ q: "Quantos projetos praticos principais existem no modulo 9?", a: "3", opcoes: ["2", "3", "4"] }],
  },
];

const daxExamples = [
  "Receita Total = SUM(Vendas[Valor_Total])",
  "Ticket Medio = DIVIDE([Receita Total], [Qtd Pedidos], 0)",
  "Lucro % = DIVIDE(SUM(Vendas[Lucro]), SUM(Vendas[Valor_Total]), 0)",
  "Vendas Ano Anterior = CALCULATE([Receita Total], SAMEPERIODLASTYEAR(Calendario[Data]))",
  "Acumulado Ano = TOTALYTD([Receita Total], Calendario[Data])",
];

const bonusQuizByModule: Record<string, CourseQuiz[]> = {
  m1: [
    {
      q: "Qual principal objetivo de um dashboard de BI?",
      a: "Apoiar decisoes com dados",
      opcoes: ["Apoiar decisoes com dados", "Substituir planilhas sem analise", "Criar artes visuais"],
      explicacao: "Dashboard bom nao e so bonito, ele orienta decisoes com rapidez.",
    },
  ],
  m2: [
    {
      q: "Depois de importar os dados, qual passo vem antes de criar graficos?",
      a: "Validar qualidade dos dados",
      opcoes: ["Validar qualidade dos dados", "Publicar no Service", "Compartilhar por e-mail"],
      explicacao: "Sempre confira tipos, nulos e duplicatas antes de visualizar.",
    },
  ],
  m3: [
    {
      q: "O Unpivot e usado para:",
      a: "Transformar colunas em linhas",
      opcoes: ["Transformar colunas em linhas", "Somar tabelas automaticamente", "Criar relacionamento N:N"],
      explicacao: "Unpivot ajuda a normalizar formatos para analise por periodo.",
    },
  ],
  m4: [
    {
      q: "No modelo estrela, a tabela fato geralmente contem:",
      a: "Medidas numericas e chaves",
      opcoes: ["Medidas numericas e chaves", "Apenas textos descritivos", "Apenas listas de usuarios"],
      explicacao: "A fato concentra eventos mensuraveis do negocio.",
    },
  ],
  m5: [
    {
      q: "Qual funcao evita erro de divisao por zero em DAX?",
      a: "DIVIDE",
      opcoes: ["DIVIDE", "VALUE", "IFERRORSQL"],
      explicacao: "DIVIDE permite definir retorno alternativo em divisao invalida.",
    },
  ],
  m6: [
    {
      q: "Para comparar participacao de categorias no total, qual visual e comum?",
      a: "Donut/Pizza",
      opcoes: ["Donut/Pizza", "Dispersao", "Gauge"],
      explicacao: "Apesar de limites, donut/pizza e util para proporcao com poucas categorias.",
    },
  ],
  m7: [
    {
      q: "Qual recurso agenda atualizacao automatica de dados locais?",
      a: "Gateway",
      opcoes: ["Gateway", "Bookmark", "Tooltip"],
      explicacao: "Gateway faz ponte segura entre ambiente local e Service.",
    },
  ],
  m8: [
    {
      q: "Boa pratica de performance e:",
      a: "Remover colunas nao usadas",
      opcoes: ["Remover colunas nao usadas", "Duplicar tabelas grandes", "Criar medidas sem padrao"],
      explicacao: "Modelos menores tendem a consultar e atualizar mais rapido.",
    },
  ],
  m9: [
    {
      q: "Projeto final de BI deve terminar com:",
      a: "Recomendacoes de negocio",
      opcoes: ["Recomendacoes de negocio", "Apenas print de graficos", "Sem conclusao"],
      explicacao: "Insight sem recomendacao gera pouco impacto no negocio.",
    },
  ],
};

const moduleDeepDive: Record<string, { objetivo: string; resultados: string[]; roteiro: string[]; erros: string[] }> = {
  m1: {
    objetivo: "Construir base conceitual para entender porque dados bem analisados geram vantagem competitiva.",
    resultados: ["Explicar BI com suas palavras", "Diferenciar produtos Power BI", "Identificar valor de um dashboard"],
    roteiro: ["Entender problema de negocio", "Mapear indicadores", "Planejar narrativa executiva"],
    erros: ["Olhar apenas grafico bonito", "Nao definir pergunta de negocio", "Ignorar contexto do usuario final"],
  },
  m2: {
    objetivo: "Fazer o primeiro relatorio de ponta a ponta com seguranca e clareza.",
    resultados: ["Importar dados corretamente", "Criar visualizacao inicial", "Organizar layout basico"],
    roteiro: ["Carregar fonte", "Conferir dados", "Criar visual", "Salvar e revisar"],
    erros: ["Importar e sair criando graficos", "Nao validar tipo de dados", "Titulos sem contexto"],
  },
  m3: {
    objetivo: "Dominar limpeza e transformacao para evitar decisoes baseadas em dados inconsistentes.",
    resultados: ["Padronizar tipos", "Remover ruido", "Aplicar merge e append com criterio"],
    roteiro: ["Auditoria de colunas", "Tratamento de qualidade", "Validacao final"],
    erros: ["Transformar sem rastrear passos", "Mesclar sem chave", "Perder historico de mudancas"],
  },
  m4: {
    objetivo: "Modelar dados para escalar analises com performance e consistencia.",
    resultados: ["Desenhar esquema estrela", "Definir cardinalidade", "Compreender fluxo de filtro"],
    roteiro: ["Separar fato/dimensao", "Criar relacionamentos", "Testar contexto"],
    erros: ["Relacionamentos ambiguos", "Muitas tabelas desconectadas", "Filtro em direcao incorreta"],
  },
  m5: {
    objetivo: "Criar medidas DAX uteis para responder perguntas de negocio com rapidez.",
    resultados: ["Criar medidas base", "Aplicar inteligencia temporal", "Comparar periodos"],
    roteiro: ["Medidas base", "Medidas derivadas", "Validar numeros com filtros"],
    erros: ["Confundir coluna e medida", "Ignorar contexto de filtro", "Formulas sem padrao"],
  },
  m6: {
    objetivo: "Escolher visual correto e reforcar leitura de insights sem poluicao.",
    resultados: ["Mapear visual ideal", "Aplicar paleta consistente", "Usar interacao com foco"],
    roteiro: ["Pergunta -> visual", "Ajustes de legibilidade", "Teste com usuario"],
    erros: ["Graficos demais", "Cores sem hierarquia", "Falta de titulo objetivo"],
  },
  m7: {
    objetivo: "Publicar com governanca, seguranca e atualizacao confiavel.",
    resultados: ["Compartilhar por perfil", "Entender RLS", "Planejar atualizacao"],
    roteiro: ["Preparar versao", "Publicar", "Aplicar seguranca", "Monitorar uso"],
    erros: ["Publicar sem controle de acesso", "Esquecer gateway", "Nao testar RLS"],
  },
  m8: {
    objetivo: "Levar dashboards para nivel profissional com performance e storytelling.",
    resultados: ["Reduzir complexidade", "Melhorar velocidade", "Escrever conclusoes acionaveis"],
    roteiro: ["Diagnostico", "Refino tecnico", "Refino narrativo"],
    erros: ["Texto longo e vago", "Muitas medidas redundantes", "Sem recomendacao final"],
  },
  m9: {
    objetivo: "Consolidar aprendizado em projetos que simulam cenarios reais do mercado.",
    resultados: ["Entregar painel executivo", "Argumentar com dados", "Defender plano de acao"],
    roteiro: ["Planejar escopo", "Construir painel", "Validar com checklist", "Apresentar insights"],
    erros: ["Sem objetivo de negocio", "Sem storytelling", "Sem plano de acompanhamento"],
  },
};

const openQuestionByModule: Record<string, string> = {
  m1: "Em ate 8 linhas, explique como BI pode melhorar uma decisao de marketing em uma empresa pequena.",
  m2: "Descreva o passo a passo para criar seu primeiro relatorio sem erros comuns.",
  m3: "Escreva um mini roteiro de limpeza de dados para evitar inconsistencias.",
  m4: "Explique como voce montaria um esquema estrela para vendas online.",
  m5: "Escreva quando voce usaria CALCULATE e por que.",
  m6: "Descreva como escolher o visual ideal para mostrar tendencia, comparacao e proporcao.",
  m7: "Monte uma politica resumida de compartilhamento seguro para um dashboard.",
  m8: "Explique 3 boas praticas para melhorar performance e leitura do dashboard.",
  m9: "Escreva sua apresentacao executiva final com 3 recomendacoes de negocio.",
};

function randomOf<T>(list: T[]) {
  return list[Math.floor(Math.random() * list.length)];
}

function formatCurrency(value: number) {
  return value.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}

function inferType(value: string | number | undefined) {
  if (value === undefined || value === null || value === "") return "texto";
  if (typeof value === "number") return "numero";
  if (/^\d{4}-\d{2}-\d{2}$/.test(String(value))) return "data";
  if (/^R\$/.test(String(value))) return "moeda";
  return "texto";
}

function parseNumber(value: string | number | undefined) {
  const num = typeof value === "number" ? value : Number(value ?? 0);
  return Number.isFinite(num) ? num : 0;
}

function generateSampleData() {
  const produtos = [
    { nome: "Notebook Pro 14", categoria: "Eletronicos", preco: 4900, custo: 3600 },
    { nome: "Monitor 27 4K", categoria: "Perifericos", preco: 2100, custo: 1450 },
    { nome: "Teclado Mecanico", categoria: "Perifericos", preco: 520, custo: 280 },
    { nome: "Mouse Gamer", categoria: "Perifericos", preco: 280, custo: 130 },
    { nome: "Headset NoiseX", categoria: "Audio", preco: 790, custo: 430 },
    { nome: "Smartphone Max", categoria: "Mobile", preco: 3300, custo: 2450 },
    { nome: "Tablet Vision", categoria: "Mobile", preco: 2400, custo: 1700 },
    { nome: "Cadeira Ergonomica", categoria: "Moveis", preco: 1900, custo: 1200 },
  ];
  const regioes = ["Sudeste", "Sul", "Nordeste", "Norte", "Centro-Oeste"];
  const estados = ["SP", "RJ", "MG", "PR", "RS", "BA", "PE", "DF", "GO", "AM"];
  const vendedores = ["Ana", "Bruno", "Camila", "Diego", "Erica", "Felipe", "Giovana", "Henrique"];

  const rows: DataRow[] = [];
  for (let i = 1; i <= 120; i += 1) {
    const item = randomOf(produtos);
    const quantidade = Math.floor(Math.random() * 8) + 1;
    const ajuste = Number((0.9 + Math.random() * 0.25).toFixed(2));
    const valorUnitario = Number((item.preco * ajuste).toFixed(2));
    const valorTotal = Number((valorUnitario * quantidade).toFixed(2));
    const custo = Number((item.custo * quantidade).toFixed(2));
    const lucro = Number((valorTotal - custo).toFixed(2));
    const data = new Date(2025, Math.floor(Math.random() * 12), Math.floor(Math.random() * 28) + 1)
      .toISOString()
      .slice(0, 10);

    rows.push({
      ID: i,
      Data: data,
      Produto: item.nome,
      Categoria: item.categoria,
      Regiao: randomOf(regioes),
      Estado: randomOf(estados),
      Vendedor: randomOf(vendedores),
      Quantidade: quantidade,
      Valor_Unitario: valorUnitario,
      Valor_Total: valorTotal,
      Custo: custo,
      Lucro: lucro,
    });
  }
  return rows;
}

function buildDefaultPage(columns: string[]): ReportPage {
  return {
    id: crypto.randomUUID(),
    name: "Pagina 1",
    background: "#12152b",
    widgets: [
      {
        id: crypto.randomUUID(),
        type: "barVertical",
        x: 30,
        y: 30,
        w: 460,
        h: 280,
        title: "Vendas por Categoria",
        color: "#00d4ff",
        xField: columns.includes("Categoria") ? "Categoria" : columns[0],
        yField: columns.includes("Valor_Total") ? "Valor_Total" : columns[1],
      },
      {
        id: crypto.randomUUID(),
        type: "line",
        x: 510,
        y: 30,
        w: 460,
        h: 280,
        title: "Tendencia de Receita",
        color: "#e94560",
        xField: columns.includes("Data") ? "Data" : columns[0],
        yField: columns.includes("Valor_Total") ? "Valor_Total" : columns[1],
      },
    ],
  };
}

export default function App() {
  const [activeTab, setActiveTab] = useState<TabKey>("guia");
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [theme, setTheme] = useState<ThemeMode>("dark");
  const [highContrast, setHighContrast] = useState(false);
  const [textScale, setTextScale] = useState(1);
  const [userName, setUserName] = useState("");
  const [nameInput, setNameInput] = useState("");
  const [rows, setRows] = useState<DataRow[]>([]);
  const [columns, setColumns] = useState<string[]>([]);
  const [columnAliases, setColumnAliases] = useState<Record<string, string>>({});
  const [deletedColumns, setDeletedColumns] = useState<string[]>([]);
  const [searchImport, setSearchImport] = useState("");
  const [importColumn, setImportColumn] = useState("");
  const [importValue, setImportValue] = useState("");
  const [pages, setPages] = useState<ReportPage[]>([]);
  const [activePageId, setActivePageId] = useState("");
  const [selectedWidgetId, setSelectedWidgetId] = useState("");
  const [zoom, setZoom] = useState(1);
  const [showGrid, setShowGrid] = useState(true);
  const [draggingId, setDraggingId] = useState<string | null>(null);
  const [filters, setFilters] = useState({
    dateStart: "",
    dateEnd: "",
    category: "Todos",
    regiao: "Todos",
    valorMax: 100000,
  });
  const [crossFilter, setCrossFilter] = useState<{ field: string; value: string } | null>(null);
  const [measures, setMeasures] = useState<Measure[]>([]);
  const [measureForm, setMeasureForm] = useState({ nome: "", funcao: "SOMA" as MeasureType, campo: "", formula: "" });
  const [completedLessons, setCompletedLessons] = useState<string[]>([]);
  const [completedTasks, setCompletedTasks] = useState<string[]>([]);
  const [passedQuizzes, setPassedQuizzes] = useState<string[]>([]);
  const [quizAnswers, setQuizAnswers] = useState<Record<string, string>>({});
  const [taskResponses, setTaskResponses] = useState<Record<string, string>>({});
  const [openQuizResponses, setOpenQuizResponses] = useState<Record<string, string>>({});
  const [searchGuide, setSearchGuide] = useState("");
  const [selectedModuleId, setSelectedModuleId] = useState("m1");
  const [guideView, setGuideView] = useState<"modulo" | "aula" | "tarefa" | "quiz">("modulo");
  const [selectedLessonId, setSelectedLessonId] = useState("");
  const [selectedTaskId, setSelectedTaskId] = useState("");
  const [profile, setProfile] = useState({ nome: "Analista", cargo: "BI Analyst", foto: "" });
  const reportRef = useRef<HTMLDivElement | null>(null);

  const palette = useMemo(() => {
    if (theme === "light") {
      return {
        bg: highContrast ? "#ffffff" : "#f5f5f7",
        panel: "#ffffff",
        panelSoft: "#f0f2f8",
        sidebar: "#ffffff",
        text: "#222631",
        textMuted: "#5f6777",
        border: highContrast ? "#000000" : "#d8deec",
      };
    }
    return {
      bg: highContrast ? "#000000" : "#0f0f1a",
      panel: highContrast ? "#101010" : "#16213e",
      panelSoft: highContrast ? "#141414" : "#1a1a3e",
      sidebar: highContrast ? "#050505" : "#162447",
      text: "#ffffff",
      textMuted: highContrast ? "#d0d0d0" : "#a0a0b0",
      border: highContrast ? "#ffffff" : "#2a2a4a",
    };
  }, [theme, highContrast]);

  const activePage = useMemo(() => pages.find((p) => p.id === activePageId), [pages, activePageId]);
  const selectedWidget = useMemo(
    () => activePage?.widgets.find((w) => w.id === selectedWidgetId),
    [activePage, selectedWidgetId],
  );

  useEffect(() => {
    const storedName = localStorage.getItem(USER_KEY);
    if (storedName) setUserName(storedName);

    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      const state = JSON.parse(saved);
      setTheme(state.theme ?? "dark");
      setHighContrast(Boolean(state.highContrast));
      setTextScale(state.textScale ?? 1);
      setRows(state.rows ?? generateSampleData());
      const cols = state.columns ?? Object.keys((state.rows ?? generateSampleData())[0]);
      setColumns(cols);
      setColumnAliases(state.columnAliases ?? {});
      setDeletedColumns(state.deletedColumns ?? []);
      setPages(state.pages ?? [buildDefaultPage(cols)]);
      setActivePageId(state.activePageId ?? (state.pages?.[0]?.id ?? ""));
      setMeasures(state.measures ?? []);
      setCompletedLessons(state.completedLessons ?? []);
      setCompletedTasks(state.completedTasks ?? []);
      setPassedQuizzes(state.passedQuizzes ?? []);
      setTaskResponses(state.taskResponses ?? {});
      setOpenQuizResponses(state.openQuizResponses ?? {});
      setProfile(state.profile ?? { nome: "Analista", cargo: "BI Analyst", foto: "" });
      setSelectedModuleId(state.selectedModuleId ?? "m1");
      setGuideView(state.guideView ?? "modulo");
      setSelectedLessonId(state.selectedLessonId ?? "");
      setSelectedTaskId(state.selectedTaskId ?? "");
    } else {
      const sample = generateSampleData();
      const cols = Object.keys(sample[0]);
      setRows(sample);
      setColumns(cols);
      const firstPage = buildDefaultPage(cols);
      setPages([firstPage]);
      setActivePageId(firstPage.id);
    }
  }, []);

  useEffect(() => {
    if (!rows.length || !pages.length) return;
    const state = {
      theme,
      highContrast,
      textScale,
      rows,
      columns,
      columnAliases,
      deletedColumns,
      pages,
      activePageId,
      measures,
      completedLessons,
      completedTasks,
      passedQuizzes,
      taskResponses,
      openQuizResponses,
      selectedModuleId,
      guideView,
      selectedLessonId,
      selectedTaskId,
      profile,
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }, [
    theme,
    highContrast,
    textScale,
    rows,
    columns,
    columnAliases,
    deletedColumns,
    pages,
    activePageId,
    measures,
    completedLessons,
    completedTasks,
    passedQuizzes,
    taskResponses,
    openQuizResponses,
    selectedModuleId,
    guideView,
    selectedLessonId,
    selectedTaskId,
    profile,
  ]);

  useEffect(() => {
    const id = setInterval(() => {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify({
          theme,
          highContrast,
          textScale,
          rows,
          columns,
          columnAliases,
          deletedColumns,
          pages,
          activePageId,
          measures,
          completedLessons,
          completedTasks,
          passedQuizzes,
          taskResponses,
          openQuizResponses,
          selectedModuleId,
          guideView,
          selectedLessonId,
          selectedTaskId,
          profile,
        }),
      );
    }, 120000);
    return () => clearInterval(id);
  }, [theme, highContrast, textScale, rows, columns, columnAliases, deletedColumns, pages, activePageId, measures, completedLessons, completedTasks, passedQuizzes, taskResponses, openQuizResponses, selectedModuleId, guideView, selectedLessonId, selectedTaskId, profile]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.ctrlKey && e.key.toLowerCase() === "s") {
        e.preventDefault();
        localStorage.setItem(STORAGE_KEY, JSON.stringify({ theme, highContrast, textScale, rows, columns, columnAliases, deletedColumns, pages, activePageId, measures, completedLessons, completedTasks, passedQuizzes, taskResponses, openQuizResponses, selectedModuleId, guideView, selectedLessonId, selectedTaskId, profile }));
      }
      if (e.ctrlKey && e.key === "1") setActiveTab("dashboard");
      if (e.ctrlKey && e.key === "2") setActiveTab("importar");
      if (e.ctrlKey && e.key === "3") setActiveTab("editor");
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [theme, highContrast, textScale, rows, columns, columnAliases, deletedColumns, pages, activePageId, measures, completedLessons, completedTasks, passedQuizzes, taskResponses, openQuizResponses, selectedModuleId, guideView, selectedLessonId, selectedTaskId, profile]);

  const visibleColumns = useMemo(() => columns.filter((c) => !deletedColumns.includes(c)), [columns, deletedColumns]);

  const dataAfterImportFilter = useMemo(() => {
    return rows.filter((row) => {
      if (!searchImport) return true;
      return visibleColumns.some((c) => String(row[c] ?? "").toLowerCase().includes(searchImport.toLowerCase()));
    });
  }, [rows, visibleColumns, searchImport]);

  const dataFiltered = useMemo(() => {
    return dataAfterImportFilter.filter((row) => {
      const data = String(row.Data ?? "");
      const categoria = String(row.Categoria ?? "");
      const regiao = String(row.Regiao ?? "");
      const valor = parseNumber(row.Valor_Total);
      const passDateStart = !filters.dateStart || data >= filters.dateStart;
      const passDateEnd = !filters.dateEnd || data <= filters.dateEnd;
      const passCategory = filters.category === "Todos" || categoria === filters.category;
      const passRegiao = filters.regiao === "Todos" || regiao === filters.regiao;
      const passValor = valor <= filters.valorMax;
      const passCross = !crossFilter || String(row[crossFilter.field]) === crossFilter.value;
      const passImport = !importColumn || !importValue || String(row[importColumn] ?? "").toLowerCase().includes(importValue.toLowerCase());
      return passDateStart && passDateEnd && passCategory && passRegiao && passValor && passCross && passImport;
    });
  }, [dataAfterImportFilter, filters, crossFilter, importColumn, importValue]);

  const categoryOptions = useMemo(
    () => ["Todos", ...new Set(rows.map((r) => String(r.Categoria ?? "Sem categoria")))],
    [rows],
  );
  const regiaoOptions = useMemo(() => ["Todos", ...new Set(rows.map((r) => String(r.Regiao ?? "")))], [rows]);

  const kpis = useMemo(() => {
    const receita = dataFiltered.reduce((acc, row) => acc + parseNumber(row.Valor_Total), 0);
    const vendas = dataFiltered.reduce((acc, row) => acc + parseNumber(row.Quantidade), 0);
    const clientes = new Set(dataFiltered.map((r) => String(r.Vendedor))).size * 14;
    const lucro = dataFiltered.reduce((acc, row) => acc + parseNumber(row.Lucro), 0);
    const crescimento = receita ? (lucro / receita) * 100 : 0;
    return { receita, vendas, clientes, crescimento, lucro };
  }, [dataFiltered]);

  function applyRows(newRows: DataRow[]) {
    if (!newRows.length) return;
    const cols = Object.keys(newRows[0]);
    setRows(newRows);
    setColumns(cols);
    setDeletedColumns([]);
    const aliases: Record<string, string> = {};
    cols.forEach((c) => {
      aliases[c] = c;
    });
    setColumnAliases(aliases);
    if (!pages.length) {
      const page = buildDefaultPage(cols);
      setPages([page]);
      setActivePageId(page.id);
    }
  }

  function handleFileUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    const ext = file.name.split(".").pop()?.toLowerCase();

    if (ext === "csv" || ext === "txt") {
      Papa.parse(file, {
        header: true,
        dynamicTyping: true,
        skipEmptyLines: true,
        complete: (result: Papa.ParseResult<DataRow>) => applyRows(result.data as DataRow[]),
      });
      return;
    }

    if (ext === "json") {
      const reader = new FileReader();
      reader.onload = () => {
        const content = JSON.parse(String(reader.result ?? "[]"));
        applyRows(Array.isArray(content) ? content : content.rows ?? []);
      };
      reader.readAsText(file);
      return;
    }

    if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = (e) => {
        const buffer = e.target?.result;
        const workbook = XLSX.read(buffer, { type: "array" });
        const firstSheet = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheet];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        applyRows(jsonData as DataRow[]);
      };
      reader.readAsArrayBuffer(file);
    }
  }

  function addWidget(type: WidgetType) {
    if (!activePage || !visibleColumns.length) return;
    const newWidget: Widget = {
      id: crypto.randomUUID(),
      type,
      x: 60,
      y: 60,
      w: 380,
      h: 240,
      title: `Novo ${type}`,
      color: "#00d4ff",
      xField: visibleColumns[0],
      yField: visibleColumns[1] ?? visibleColumns[0],
      text: "Digite seu texto aqui",
      imageUrl: "",
    };
    setPages((prev) => prev.map((p) => (p.id === activePage.id ? { ...p, widgets: [...p.widgets, newWidget] } : p)));
    setSelectedWidgetId(newWidget.id);
  }

  function updateWidget(id: string, patch: Partial<Widget>) {
    setPages((prev) =>
      prev.map((p) =>
        p.id === activePageId
          ? {
              ...p,
              widgets: p.widgets.map((w) => (w.id === id ? { ...w, ...patch } : w)),
            }
          : p,
      ),
    );
  }

  function duplicateWidget(id: string) {
    if (!activePage) return;
    const widget = activePage.widgets.find((w) => w.id === id);
    if (!widget) return;
    const copy = { ...widget, id: crypto.randomUUID(), x: widget.x + 20, y: widget.y + 20 };
    setPages((prev) => prev.map((p) => (p.id === activePage.id ? { ...p, widgets: [...p.widgets, copy] } : p)));
  }

  function deleteWidget(id: string) {
    setPages((prev) => prev.map((p) => (p.id === activePageId ? { ...p, widgets: p.widgets.filter((w) => w.id !== id) } : p)));
    if (selectedWidgetId === id) setSelectedWidgetId("");
  }

  function alignWidgets(mode: "left" | "center" | "right") {
    if (!activePage) return;
    setPages((prev) =>
      prev.map((p) => {
        if (p.id !== activePage.id) return p;
        return {
          ...p,
          widgets: p.widgets.map((w) => {
            if (mode === "left") return { ...w, x: 20 };
            if (mode === "center") return { ...w, x: 480 - w.w / 2 };
            return { ...w, x: 940 - w.w };
          }),
        };
      }),
    );
  }

  async function exportCurrentAsImage() {
    if (!reportRef.current) return;
    const canvas = await html2canvas(reportRef.current, { scale: 2 });
    const link = document.createElement("a");
    link.download = "dashboard-pbc.png";
    link.href = canvas.toDataURL("image/png");
    link.click();
  }

  async function exportCurrentAsPdf() {
    if (!reportRef.current) return;
    const canvas = await html2canvas(reportRef.current, { scale: 2 });
    const img = canvas.toDataURL("image/png");
    const pdf = new jsPDF("landscape", "pt", "a4");
    const width = pdf.internal.pageSize.getWidth();
    const height = (canvas.height * width) / canvas.width;
    pdf.addImage(img, "PNG", 0, 20, width, height);
    pdf.save("relatorio-pbc.pdf");
  }

  function exportDataCsv() {
    const csv = Papa.unparse(dataFiltered);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.setAttribute("download", "dados-filtrados.csv");
    link.click();
    URL.revokeObjectURL(url);
  }

  function exportProjectJson() {
    const payload = {
      theme,
      highContrast,
      textScale,
      rows,
      columns,
      columnAliases,
      deletedColumns,
      pages,
      activePageId,
      measures,
      completedLessons,
      completedTasks,
      passedQuizzes,
      taskResponses,
      openQuizResponses,
      selectedModuleId,
      guideView,
      selectedLessonId,
      selectedTaskId,
      profile,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "projeto-power-bi-chrystopher.json";
    a.click();
    URL.revokeObjectURL(url);
  }

  function importProjectJson(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const state = JSON.parse(String(reader.result || "{}"));
      setTheme(state.theme ?? "dark");
      setHighContrast(Boolean(state.highContrast));
      setTextScale(state.textScale ?? 1);
      setRows(state.rows ?? []);
      setColumns(state.columns ?? []);
      setColumnAliases(state.columnAliases ?? {});
      setDeletedColumns(state.deletedColumns ?? []);
      setPages(state.pages ?? []);
      setActivePageId(state.activePageId ?? state.pages?.[0]?.id ?? "");
      setMeasures(state.measures ?? []);
      setCompletedLessons(state.completedLessons ?? []);
      setCompletedTasks(state.completedTasks ?? []);
      setPassedQuizzes(state.passedQuizzes ?? []);
      setTaskResponses(state.taskResponses ?? {});
      setOpenQuizResponses(state.openQuizResponses ?? {});
      setSelectedModuleId(state.selectedModuleId ?? "m1");
      setGuideView(state.guideView ?? "modulo");
      setSelectedLessonId(state.selectedLessonId ?? "");
      setSelectedTaskId(state.selectedTaskId ?? "");
      setProfile(state.profile ?? { nome: "Analista", cargo: "BI Analyst", foto: "" });
    };
    reader.readAsText(file);
  }

  function runMeasure(measure: Measure) {
    const values = dataFiltered.map((r) => parseNumber(r[measure.campo]));
    const sum = values.reduce((a, b) => a + b, 0);
    if (measure.funcao === "SOMA") return sum;
    if (measure.funcao === "MEDIA") return values.length ? sum / values.length : 0;
    if (measure.funcao === "CONTAGEM") return values.length;
    if (measure.funcao === "MINIMO") return Math.min(...values);
    if (measure.funcao === "MAXIMO") return Math.max(...values);
    if (measure.funcao === "PORCENTAGEM") return sum ? (sum / (kpis.receita || 1)) * 100 : 0;
    if (measure.funcao === "VARIACAO") return kpis.receita ? (kpis.lucro / kpis.receita) * 100 : 0;
    if (measure.funcao === "ACUMULADO") return sum;
    if (measure.funcao === "YOY") return (Math.random() * 16 - 4).toFixed(2);

    try {
      const val = Number(Function("SUM", `return ${measure.formula}`)(sum));
      return Number.isFinite(val) ? val : 0;
    } catch {
      return 0;
    }
  }

  function renderWidget(widget: Widget) {
    const xField = widget.xField ?? visibleColumns[0];
    const yField = widget.yField ?? visibleColumns[1];
    const labels = [...new Set(dataFiltered.map((r) => String(r[xField] ?? "Sem valor")))].slice(0, 12);
    const values = labels.map((label) =>
      dataFiltered
        .filter((r) => String(r[xField]) === label)
        .reduce((sum, row) => sum + parseNumber(row[yField]), 0),
    );

    const chartData = {
      labels,
      datasets: [
        {
          label: widget.title,
          data: values,
          borderColor: widget.color,
          backgroundColor: `${widget.color}88`,
          fill: widget.type === "area",
          tension: 0.35,
        },
      ],
    };

    const commonProps = {
      data: chartData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { labels: { color: palette.textMuted } } },
        scales: {
          x: { ticks: { color: palette.textMuted }, grid: { color: `${palette.border}55` } },
          y: { ticks: { color: palette.textMuted }, grid: { color: `${palette.border}55` } },
        },
        onClick: (_: unknown, elements: { index: number }[]) => {
          if (elements?.length) {
            const label = labels[elements[0].index];
            setCrossFilter({ field: xField, value: label });
          }
        },
      },
    };

    if (widget.type === "barVertical") return <Bar {...commonProps} />;
    if (widget.type === "barHorizontal") return <Bar {...commonProps} options={{ ...commonProps.options, indexAxis: "y" as const }} />;
    if (widget.type === "line") return <Line {...commonProps} />;
    if (widget.type === "area") return <Line {...commonProps} />;
    if (widget.type === "donut") return <Doughnut {...commonProps} options={{ ...commonProps.options, scales: undefined }} />;
    if (widget.type === "combo") {
      return (
        <Bar
          data={{
            labels,
            datasets: [
              { type: "bar", label: "Barras", data: values, backgroundColor: `${widget.color}66` },
              { type: "line", label: "Linha", data: values.map((v) => Number((v * 0.8).toFixed(2))), borderColor: "#e94560" },
            ],
          } as never}
          options={commonProps.options as never}
        />
      );
    }
    if (widget.type === "scatter") {
      const points = dataFiltered.slice(0, 80).map((r) => ({ x: parseNumber(r[xField]), y: parseNumber(r[yField]) }));
      return (
        <Scatter
          data={{ datasets: [{ label: widget.title, data: points, backgroundColor: `${widget.color}99` }] }}
          options={{ ...commonProps.options, scales: { x: { ticks: { color: palette.textMuted } }, y: { ticks: { color: palette.textMuted } } } }}
        />
      );
    }
    if (widget.type === "funnel") {
      const sorted = [...values].sort((a, b) => b - a).slice(0, 6);
      return (
        <Bar
          data={{ labels: labels.slice(0, sorted.length), datasets: [{ label: "Funil", data: sorted, backgroundColor: `${widget.color}88` }] }}
          options={{ ...commonProps.options, indexAxis: "y" as const }}
        />
      );
    }
    if (widget.type === "gauge") {
      const ratio = Math.max(0, Math.min(100, kpis.crescimento));
      return (
        <div className="flex h-full flex-col items-center justify-center">
          <div className="h-36 w-36">
            <Doughnut
              data={{ labels: ["Atual", "Restante"], datasets: [{ data: [ratio, 100 - ratio], backgroundColor: ["#00c853", `${palette.border}66`], borderWidth: 0 }] }}
              options={{ circumference: 180, rotation: 270, cutout: "70%", plugins: { legend: { display: false } } }}
            />
          </div>
          <p className="text-lg font-semibold" style={{ color: palette.text }}>{ratio.toFixed(1)}%</p>
        </div>
      );
    }
    if (widget.type === "heatmap") {
      const yCats = [...new Set(dataFiltered.map((r) => String(r[yField] ?? "-")))].slice(0, 6);
      return (
        <div className="grid h-full gap-1 overflow-auto" style={{ gridTemplateColumns: `repeat(${Math.min(6, labels.length)}, minmax(80px,1fr))` }}>
          {labels.slice(0, 6).map((lx) =>
            yCats.map((ly) => {
              const count = dataFiltered.filter((r) => String(r[xField]) === lx && String(r[yField]) === ly).length;
              const alpha = Math.min(0.95, count / 12 + 0.1);
              return (
                <div key={`${lx}-${ly}`} className="rounded p-2 text-xs" style={{ background: `rgba(0,212,255,${alpha})`, color: "#001" }}>
                  <div>{lx}</div>
                  <div>{ly}</div>
                  <strong>{count}</strong>
                </div>
              );
            }),
          )}
        </div>
      );
    }
    if (widget.type === "kpi") {
      const val = dataFiltered.reduce((acc, row) => acc + parseNumber(row[yField]), 0);
      return (
        <div className="flex h-full items-center justify-center text-center">
          <div>
            <p className="text-sm" style={{ color: palette.textMuted }}>{widget.title}</p>
            <p className="text-4xl font-bold" style={{ color: widget.color }}>{formatCurrency(val)}</p>
          </div>
        </div>
      );
    }
    if (widget.type === "table") {
      return (
        <div className="h-full overflow-auto">
          <table className="w-full text-xs">
            <thead>
              <tr>
                {visibleColumns.slice(0, 5).map((col) => (
                  <th key={col} className="border-b p-1 text-left" style={{ borderColor: palette.border }}>{col}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {dataFiltered.slice(0, 10).map((row, idx) => (
                <tr key={idx}>
                  {visibleColumns.slice(0, 5).map((col) => (
                    <td key={col} className="border-b p-1" style={{ borderColor: palette.border }}>{String(row[col] ?? "")}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
    }
    if (widget.type === "text") {
      return <div className="p-3 text-sm whitespace-pre-wrap">{widget.text}</div>;
    }
    if (widget.type === "shapeRect") return <div className="h-full w-full rounded border-2" style={{ borderColor: widget.color }} />;
    if (widget.type === "shapeLine") return <div className="mt-1 h-[2px] w-full" style={{ background: widget.color }} />;
    if (widget.type === "image") return widget.imageUrl ? <img src={widget.imageUrl} alt="Imagem do relatorio" className="h-full w-full object-contain" /> : <p className="p-2 text-xs">Adicione uma URL de imagem.</p>;
    return null;
  }

  function saveLesson(moduleId: string, lessonId: string) {
    const key = `${moduleId}:${lessonId}`;
    if (!completedLessons.includes(key)) setCompletedLessons((prev) => [...prev, key]);
  }

  function isWritingTask(task: CourseTask) {
    if (task.tipo) return task.tipo === "escrita";
    return /escreva|descreva|explique|liste|conclusao/i.test(task.instrucoes);
  }

  function toggleTask(moduleId: string, taskId: string) {
    const moduleData = learningModules.find((mod) => mod.id === moduleId);
    const task = moduleData?.tarefas.find((item) => item.id === taskId);
    const key = `${moduleId}:task:${taskId}`;
    const textKey = `${moduleId}:taskText:${taskId}`;
    if (task && isWritingTask(task)) {
      const answer = (taskResponses[textKey] ?? "").trim();
      if (answer.length < 60) return;
    }
    setCompletedTasks((prev) => (prev.includes(key) ? prev.filter((item) => item !== key) : [...prev, key]));
  }

  function getModuleQuiz(moduleId: string) {
    const mod = learningModules.find((item) => item.id === moduleId);
    if (!mod) return [] as CourseQuiz[];
    return [...mod.quiz, ...(bonusQuizByModule[moduleId] ?? [])];
  }

  function evaluateQuiz(moduleId: string) {
    const questions = getModuleQuiz(moduleId);
    if (!questions.length) return { correct: 0, total: 0, pct: 0 };
    const correct = questions.reduce((acc, qz, idx) => {
      const key = `${moduleId}-${idx}`;
      return quizAnswers[key]?.toLowerCase() === qz.a.toLowerCase() ? acc + 1 : acc;
    }, 0);
    const total = questions.length;
    return { correct, total, pct: total ? Math.round((correct / total) * 100) : 0 };
  }

  function submitQuiz(moduleId: string) {
    const result = evaluateQuiz(moduleId);
    const openAnswer = (openQuizResponses[`${moduleId}:open`] ?? "").trim();
    const quizKey = `${moduleId}:quiz`;
    if (result.pct >= 70 && openAnswer.length >= 80 && !passedQuizzes.includes(quizKey)) {
      setPassedQuizzes((prev) => [...prev, quizKey]);
    }
  }

  const progressByModule = useMemo(() => {
    return learningModules.map((mod) => {
      const lessonTotal = mod.aulas.length;
      const lessonDone = mod.aulas.filter((aula) => completedLessons.includes(`${mod.id}:${aula.id}`)).length;
      const taskTotal = mod.tarefas.length;
      const taskDone = mod.tarefas.filter((task) => completedTasks.includes(`${mod.id}:task:${task.id}`)).length;
      const quizDone = passedQuizzes.includes(`${mod.id}:quiz`) ? 1 : 0;
      const totalUnits = lessonTotal + taskTotal + 1;
      const doneUnits = lessonDone + taskDone + quizDone;
      return {
        id: mod.id,
        lessonDone,
        lessonTotal,
        taskDone,
        taskTotal,
        quizDone,
        pct: totalUnits ? Math.round((doneUnits / totalUnits) * 100) : 0,
      };
    });
  }, [completedLessons, completedTasks, passedQuizzes]);

  const totalPoints = useMemo(() => {
    const lessonPoints = completedLessons.length * 10;
    const taskPoints = learningModules.reduce((acc, mod) => {
      return (
        acc +
        mod.tarefas.reduce((sum, task) => {
          return completedTasks.includes(`${mod.id}:task:${task.id}`) ? sum + task.pontos : sum;
        }, 0)
      );
    }, 0);
    const quizPoints = passedQuizzes.length * 40;
    return lessonPoints + taskPoints + quizPoints;
  }, [completedLessons, completedTasks, passedQuizzes]);

  const selectedModule = useMemo(
    () => learningModules.find((mod) => mod.id === selectedModuleId) ?? learningModules[0],
    [selectedModuleId],
  );

  const selectedLesson = useMemo(
    () => selectedModule.aulas.find((aula) => aula.id === selectedLessonId) ?? selectedModule.aulas[0],
    [selectedModule, selectedLessonId],
  );

  const selectedTask = useMemo(
    () => selectedModule.tarefas.find((task) => task.id === selectedTaskId) ?? selectedModule.tarefas[0],
    [selectedModule, selectedTaskId],
  );

  const selectedQuizQuestions = useMemo(() => getModuleQuiz(selectedModule.id), [selectedModule.id]);
  const selectedQuizResult = useMemo(
    () => evaluateQuiz(selectedModule.id),
    [selectedModule.id, quizAnswers],
  );

  useEffect(() => {
    if (!selectedModule.aulas.some((aula) => aula.id === selectedLessonId)) {
      setSelectedLessonId(selectedModule.aulas[0]?.id ?? "");
    }
    if (!selectedModule.tarefas.some((task) => task.id === selectedTaskId)) {
      setSelectedTaskId(selectedModule.tarefas[0]?.id ?? "");
    }
  }, [selectedModule, selectedLessonId, selectedTaskId]);

  if (!userName) {
    return (
      <main className="min-h-screen overflow-hidden" style={{ background: "radial-gradient(circle at 20% 20%, #533483, #0f0f1a 60%)" }}>
        <div className="mx-auto flex min-h-screen max-w-5xl items-center px-6 py-14">
          <motion.section
            initial={{ opacity: 0, y: 24 }}
            animate={{ opacity: 1, y: 0 }}
            className="w-full rounded-2xl border p-8 md:p-12"
            style={{ borderColor: "#2a2a4a", background: "rgba(22,33,62,0.86)", color: "#fff" }}
          >
            <div className="mb-10 flex items-center gap-4">
              <div className="flex h-14 w-14 items-center justify-center rounded-xl bg-cyan-400 font-black text-slate-900">PBC</div>
              <div>
                <p className="text-sm uppercase tracking-[0.2em] text-cyan-300">Power BI Chrystopher</p>
                <h1 className="text-3xl font-semibold">Sua plataforma de BI no navegador</h1>
              </div>
            </div>
            <p className="max-w-2xl text-slate-200">
              "Sem dados, voce e apenas mais uma pessoa com opinioes." Entre para criar dashboards, analisar vendas e evoluir com um mini-curso completo de Power BI.
            </p>
            <div className="mt-8 flex flex-col gap-3 md:flex-row">
              <input
                aria-label="Nome do usuario"
                className="h-12 flex-1 rounded-lg border px-4 text-base outline-none"
                style={{ background: "#0f0f1a", borderColor: "#2a2a4a" }}
                placeholder="Digite seu nome"
                value={nameInput}
                onChange={(e) => setNameInput(e.target.value)}
              />
              <button
                className="h-12 rounded-lg bg-cyan-400 px-6 font-semibold text-slate-900 transition hover:brightness-95"
                onClick={() => {
                  if (!nameInput.trim()) return;
                  localStorage.setItem(USER_KEY, nameInput.trim());
                  setUserName(nameInput.trim());
                }}
              >
                Entrar na Plataforma
              </button>
            </div>
          </motion.section>
        </div>
      </main>
    );
  }

  return (
    <div style={{ background: palette.bg, color: palette.text, fontSize: `${textScale}rem` }} className="min-h-screen">
      <div className="flex min-h-screen">
        <aside
          className="border-r px-3 py-4 transition-all"
          style={{ width: sidebarCollapsed ? 82 : 272, background: palette.sidebar, borderColor: palette.border }}
        >
          <div className="mb-6 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-cyan-400 font-black text-slate-900">PBC</div>
              {!sidebarCollapsed && <div className="text-sm leading-tight"><strong>Power BI</strong><br />Chrystopher</div>}
            </div>
            <button aria-label="Colapsar menu" onClick={() => setSidebarCollapsed((v) => !v)} className="rounded p-1 hover:bg-white/10">{sidebarCollapsed ? ">" : "<"}</button>
          </div>
          <nav className="space-y-1">
            {tabs.map((tab) => {
              const Icon = tab.icon;
              return (
                <button
                  key={tab.key}
                  onClick={() => setActiveTab(tab.key)}
                  className="flex w-full items-center gap-3 rounded-lg px-3 py-2 text-left transition"
                  style={{ background: activeTab === tab.key ? "rgba(0,212,255,0.2)" : "transparent" }}
                >
                  <Icon size={18} />
                  {!sidebarCollapsed && <span>{tab.label}</span>}
                </button>
              );
            })}
          </nav>
          <div className="mt-8 space-y-2">
            <button className="flex w-full items-center justify-center gap-2 rounded-lg bg-cyan-400 py-2 text-sm font-semibold text-slate-900" onClick={() => setTheme((t) => (t === "dark" ? "light" : "dark"))}>
              {theme === "dark" ? <Sun size={16} /> : <Moon size={16} />} {sidebarCollapsed ? "" : theme === "dark" ? "Tema Claro" : "Tema Escuro"}
            </button>
          </div>
        </aside>

        <section className="flex min-h-screen flex-1 flex-col">
          <header className="flex flex-wrap items-center justify-between gap-3 border-b px-5 py-3" style={{ borderColor: palette.border, background: palette.panel }}>
            <div>
              <h2 className="text-xl font-semibold">{tabs.find((t) => t.key === activeTab)?.label}</h2>
              <p className="text-sm" style={{ color: palette.textMuted }}>Bem-vindo, {userName}. {new Date().toLocaleDateString("pt-BR")} - {new Date().toLocaleTimeString("pt-BR")}</p>
            </div>
            <div className="flex flex-wrap items-center gap-2">
              <button onClick={exportCurrentAsImage} className="rounded border px-3 py-1 text-sm" style={{ borderColor: palette.border }}>PNG</button>
              <button onClick={exportCurrentAsPdf} className="rounded border px-3 py-1 text-sm" style={{ borderColor: palette.border }}>PDF</button>
              <button onClick={exportDataCsv} className="rounded border px-3 py-1 text-sm" style={{ borderColor: palette.border }}>CSV</button>
              <button onClick={() => window.print()} className="rounded border px-3 py-1 text-sm" style={{ borderColor: palette.border }}>Imprimir</button>
              {crossFilter && (
                <button onClick={() => setCrossFilter(null)} className="rounded border px-3 py-1 text-xs" style={{ borderColor: palette.border }}>
                  Limpar filtro visual ({crossFilter.field}: {crossFilter.value})
                </button>
              )}
            </div>
          </header>

          <AnimatePresence mode="wait">
            <motion.main
              key={activeTab}
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              transition={{ duration: 0.2 }}
              className="flex-1 overflow-auto p-5"
            >
              {activeTab === "dashboard" && (
                <div className="space-y-6">
                  <div className="rounded-xl border p-4" style={{ borderColor: palette.border, background: palette.panel }}>
                    <h3 className="text-lg font-semibold">Ola, {profile.nome || userName}</h3>
                    <p style={{ color: palette.textMuted }}>Dica do dia: use medidas com CALCULATE para controlar contexto de filtro e criar analises mais confiaveis.</p>
                  </div>

                  <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
                    {[
                      { label: "Receita", value: formatCurrency(kpis.receita), trend: "+12%" },
                      { label: "Vendas", value: kpis.vendas.toLocaleString("pt-BR"), trend: "+8%" },
                      { label: "Clientes", value: kpis.clientes.toLocaleString("pt-BR"), trend: "+4%" },
                      { label: "Crescimento", value: `${kpis.crescimento.toFixed(2)}%`, trend: "+2.1%" },
                    ].map((card) => (
                      <motion.div whileHover={{ y: -3 }} key={card.label} className="rounded-xl border p-4" style={{ borderColor: palette.border, background: palette.panelSoft }}>
                        <div className="mb-3 flex items-center justify-between">
                          <span className="text-sm" style={{ color: palette.textMuted }}>{card.label}</span>
                          <span className="text-xs text-emerald-400">{card.trend}</span>
                        </div>
                        <p className="mb-3 text-2xl font-semibold">{card.value}</p>
                        <div className="h-20">
                          <Line
                            data={{ labels: ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun"], datasets: [{ data: [4, 6, 5, 8, 9, 12], borderColor: "#00d4ff", pointRadius: 0 }] }}
                            options={{ responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { display: false }, y: { display: false } } }}
                          />
                        </div>
                      </motion.div>
                    ))}
                  </div>

                  <div className="grid gap-4 lg:grid-cols-3">
                    <div className="rounded-xl border p-4 lg:col-span-2" style={{ borderColor: palette.border, background: palette.panel }}>
                      <h3 className="mb-2 text-lg font-semibold">Dashboards Recentes</h3>
                      <ul className="space-y-2 text-sm">
                        {pages.map((p) => (
                          <li key={p.id} className="flex items-center justify-between rounded border px-3 py-2" style={{ borderColor: palette.border }}>
                            <span>{p.name}</span>
                            <button className="text-cyan-400" onClick={() => { setActivePageId(p.id); setActiveTab("editor"); }}>Abrir</button>
                          </li>
                        ))}
                      </ul>
                    </div>
                    <div className="space-y-3">
                      <button className="flex w-full items-center justify-center gap-2 rounded-xl bg-cyan-400 p-3 font-semibold text-slate-900" onClick={() => setActiveTab("editor")}>Criar Novo Dashboard</button>
                      <button className="flex w-full items-center justify-center gap-2 rounded-xl border p-3" style={{ borderColor: palette.border }} onClick={() => setActiveTab("importar")}>Importar Dados</button>
                      <div className="rounded-xl border p-3 text-sm" style={{ borderColor: palette.border, background: palette.panel }}>
                        <p className="font-medium">Relogio Atual</p>
                        <p className="text-2xl">{new Date().toLocaleTimeString("pt-BR")}</p>
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {activeTab === "importar" && (
                <div className="space-y-4">
                  <div className="rounded-xl border p-4" style={{ borderColor: palette.border, background: palette.panel }}>
                    <div className="flex flex-wrap items-center gap-3">
                      <label className="inline-flex cursor-pointer items-center gap-2 rounded-lg bg-cyan-400 px-4 py-2 font-semibold text-slate-900">
                        <Upload size={16} /> Upload CSV/XLSX/JSON/TXT
                        <input type="file" accept=".csv,.xlsx,.xls,.json,.txt" onChange={handleFileUpload} className="hidden" />
                      </label>
                      <button onClick={() => applyRows(generateSampleData())} className="rounded-lg border px-4 py-2" style={{ borderColor: palette.border }}>
                        Usar Dados de Exemplo (TechStore Brasil)
                      </button>
                    </div>
                    <div className="mt-4 grid gap-3 md:grid-cols-3">
                      <input className="rounded border px-3 py-2" style={{ borderColor: palette.border }} placeholder="Pesquisar dados" value={searchImport} onChange={(e) => setSearchImport(e.target.value)} />
                      <select className="rounded border px-3 py-2" style={{ borderColor: palette.border }} value={importColumn} onChange={(e) => setImportColumn(e.target.value)}>
                        <option value="">Filtrar por coluna</option>
                        {visibleColumns.map((c) => (<option key={c}>{c}</option>))}
                      </select>
                      <input className="rounded border px-3 py-2" style={{ borderColor: palette.border }} placeholder="Contem valor" value={importValue} onChange={(e) => setImportValue(e.target.value)} />
                    </div>
                  </div>

                  <div className="rounded-xl border p-4" style={{ borderColor: palette.border, background: palette.panel }}>
                    <h3 className="mb-2 text-lg font-semibold">Tipos de dados detectados</h3>
                    <div className="grid gap-2 md:grid-cols-3">
                      {visibleColumns.map((col) => (
                        <div key={col} className="rounded border p-2 text-sm" style={{ borderColor: palette.border }}>
                          <div className="mb-1 font-medium">{col}</div>
                          <div style={{ color: palette.textMuted }}>Tipo: {inferType(rows[0]?.[col])}</div>
                          <div className="mt-2 flex gap-2">
                            <input value={columnAliases[col] ?? col} onChange={(e) => setColumnAliases((prev) => ({ ...prev, [col]: e.target.value }))} className="w-full rounded border px-2 py-1 text-xs" style={{ borderColor: palette.border }} aria-label={`Renomear ${col}`} />
                            <button className="rounded border px-2 text-xs" style={{ borderColor: palette.border }} onClick={() => setDeletedColumns((prev) => [...prev, col])}>Excluir</button>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="rounded-xl border p-4" style={{ borderColor: palette.border, background: palette.panel }}>
                    <h3 className="mb-2 text-lg font-semibold">Preview dos dados (20 linhas)</h3>
                    <div className="overflow-auto">
                      <table className="w-full text-sm">
                        <thead>
                          <tr>
                            {visibleColumns.map((c) => (
                              <th key={c} className="border-b px-2 py-1 text-left" style={{ borderColor: palette.border }}>{columnAliases[c] ?? c}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {dataFiltered.slice(0, 20).map((row, idx) => (
                            <tr key={idx}>
                              {visibleColumns.map((c) => (
                                <td key={c} className="border-b px-2 py-1" style={{ borderColor: palette.border }}>{String(row[c] ?? "")}</td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}

              {activeTab === "editor" && (
                <div className="grid gap-4 xl:grid-cols-[220px_1fr_280px]">
                  <div className="space-y-4 rounded-xl border p-3" style={{ borderColor: palette.border, background: palette.panel }}>
                    <h3 className="font-semibold">Campos</h3>
                    <div className="max-h-48 overflow-auto text-sm">
                      {visibleColumns.map((c) => (
                        <div key={c} className="mb-1 rounded border px-2 py-1" style={{ borderColor: palette.border }}>{c}</div>
                      ))}
                    </div>
                    <h3 className="font-semibold">Adicionar visual</h3>
                    <div className="grid grid-cols-2 gap-2 text-xs">
                      {[
                        ["barVertical", "Barra V"], ["barHorizontal", "Barra H"], ["line", "Linha"], ["donut", "Donut"],
                        ["area", "Area"], ["kpi", "KPI"], ["table", "Tabela"], ["scatter", "Scatter"], ["funnel", "Funil"],
                        ["gauge", "Gauge"], ["heatmap", "Heatmap"], ["combo", "Combinado"], ["text", "Texto"], ["shapeRect", "Retangulo"], ["shapeLine", "Linha Div"], ["image", "Imagem"],
                      ].map(([type, label]) => (
                        <button key={type} className="rounded border px-2 py-1" style={{ borderColor: palette.border }} onClick={() => addWidget(type as WidgetType)}>{label}</button>
                      ))}
                    </div>
                  </div>

                  <div>
                    <div className="mb-3 flex flex-wrap items-center gap-2">
                      <button className="rounded border px-2 py-1 text-xs" style={{ borderColor: palette.border }} onClick={() => alignWidgets("left")}>Alinhar esquerda</button>
                      <button className="rounded border px-2 py-1 text-xs" style={{ borderColor: palette.border }} onClick={() => alignWidgets("center")}>Centralizar</button>
                      <button className="rounded border px-2 py-1 text-xs" style={{ borderColor: palette.border }} onClick={() => alignWidgets("right")}>Alinhar direita</button>
                      <button className="rounded border px-2 py-1 text-xs" style={{ borderColor: palette.border }} onClick={() => setZoom((z) => Math.max(0.6, z - 0.1))}>Zoom -</button>
                      <button className="rounded border px-2 py-1 text-xs" style={{ borderColor: palette.border }} onClick={() => setZoom((z) => Math.min(1.4, z + 0.1))}>Zoom +</button>
                      <button className="rounded border px-2 py-1 text-xs" style={{ borderColor: palette.border }} onClick={() => setShowGrid((v) => !v)}>{showGrid ? "Ocultar" : "Mostrar"} grade</button>
                    </div>

                    <div className="overflow-auto rounded-xl border p-2" style={{ borderColor: palette.border, background: palette.panel }}>
                      <div ref={reportRef} className="relative mx-auto h-[720px] w-[980px] origin-top-left overflow-hidden rounded-lg" style={{ transform: `scale(${zoom})`, transformOrigin: "top left", background: activePage?.background ?? "#12152b", backgroundImage: showGrid ? "linear-gradient(rgba(255,255,255,0.05) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,0.05) 1px, transparent 1px)" : "none", backgroundSize: "20px 20px" }}>
                        {activePage?.widgets.map((widget) => (
                          <div
                            key={widget.id}
                            className="absolute rounded border"
                            style={{ left: widget.x, top: widget.y, width: widget.w, height: widget.h, borderColor: selectedWidgetId === widget.id ? "#00d4ff" : palette.border, background: "rgba(11,15,34,0.7)" }}
                            onClick={() => setSelectedWidgetId(widget.id)}
                            onMouseDown={(e) => {
                              if ((e.target as HTMLElement).closest("canvas")) return;
                              setDraggingId(widget.id);
                            }}
                            onMouseUp={() => setDraggingId(null)}
                            onMouseMove={(e) => {
                              if (draggingId !== widget.id) return;
                              const parent = (e.currentTarget.parentElement as HTMLDivElement).getBoundingClientRect();
                              updateWidget(widget.id, { x: e.clientX - parent.left - widget.w / 2, y: e.clientY - parent.top - 18 });
                            }}
                          >
                            <div className="flex items-center justify-between border-b px-2 py-1 text-xs" style={{ borderColor: palette.border }}>
                              <span>{widget.title}</span>
                              <div className="flex gap-1">
                                <button className="rounded bg-white/10 px-1" onClick={() => duplicateWidget(widget.id)}>D</button>
                                <button className="rounded bg-white/10 px-1" onClick={() => deleteWidget(widget.id)}>X</button>
                              </div>
                            </div>
                            <div className="h-[calc(100%-26px)] p-1">{renderWidget(widget)}</div>
                          </div>
                        ))}
                      </div>
                    </div>

                    <div className="mt-3 flex items-center gap-2 overflow-auto">
                      {pages.map((p) => (
                        <button key={p.id} className="rounded border px-3 py-1 text-sm" style={{ borderColor: p.id === activePageId ? "#00d4ff" : palette.border }} onClick={() => setActivePageId(p.id)}>{p.name}</button>
                      ))}
                      <button className="rounded border px-3 py-1 text-sm" style={{ borderColor: palette.border }} onClick={() => {
                        const np: ReportPage = { id: crypto.randomUUID(), name: `Pagina ${pages.length + 1}`, background: "#12152b", widgets: [] };
                        setPages((prev) => [...prev, np]);
                        setActivePageId(np.id);
                      }}>+ Nova pagina</button>
                    </div>
                  </div>

                  <div className="space-y-3 rounded-xl border p-3" style={{ borderColor: palette.border, background: palette.panel }}>
                    <h3 className="font-semibold">Propriedades</h3>
                    {selectedWidget ? (
                      <>
                        <input className="w-full rounded border px-2 py-1 text-sm" style={{ borderColor: palette.border }} value={selectedWidget.title} onChange={(e) => updateWidget(selectedWidget.id, { title: e.target.value })} />
                        <label className="text-xs">Cor</label>
                        <input type="color" className="h-9 w-full rounded" value={selectedWidget.color} onChange={(e) => updateWidget(selectedWidget.id, { color: e.target.value })} />
                        <label className="text-xs">Eixo X</label>
                        <select className="w-full rounded border px-2 py-1 text-sm" style={{ borderColor: palette.border }} value={selectedWidget.xField} onChange={(e) => updateWidget(selectedWidget.id, { xField: e.target.value })}>
                          {visibleColumns.map((c) => (<option key={c}>{c}</option>))}
                        </select>
                        <label className="text-xs">Eixo Y/Valor</label>
                        <select className="w-full rounded border px-2 py-1 text-sm" style={{ borderColor: palette.border }} value={selectedWidget.yField} onChange={(e) => updateWidget(selectedWidget.id, { yField: e.target.value })}>
                          {visibleColumns.map((c) => (<option key={c}>{c}</option>))}
                        </select>
                        <div className="grid grid-cols-2 gap-2">
                          <input type="number" className="rounded border px-2 py-1 text-sm" style={{ borderColor: palette.border }} value={selectedWidget.w} onChange={(e) => updateWidget(selectedWidget.id, { w: Number(e.target.value) })} />
                          <input type="number" className="rounded border px-2 py-1 text-sm" style={{ borderColor: palette.border }} value={selectedWidget.h} onChange={(e) => updateWidget(selectedWidget.id, { h: Number(e.target.value) })} />
                        </div>
                        {selectedWidget.type === "text" && <textarea className="w-full rounded border px-2 py-1 text-sm" style={{ borderColor: palette.border }} value={selectedWidget.text} onChange={(e) => updateWidget(selectedWidget.id, { text: e.target.value })} />}
                        {selectedWidget.type === "image" && <input className="w-full rounded border px-2 py-1 text-sm" style={{ borderColor: palette.border }} placeholder="https://..." value={selectedWidget.imageUrl} onChange={(e) => updateWidget(selectedWidget.id, { imageUrl: e.target.value })} />}
                      </>
                    ) : (
                      <p className="text-sm" style={{ color: palette.textMuted }}>Selecione um visual no canvas.</p>
                    )}
                    <hr style={{ borderColor: palette.border }} />
                    <label className="text-xs">Fundo da pagina</label>
                    <input type="color" className="h-9 w-full rounded" value={activePage?.background ?? "#12152b"} onChange={(e) => setPages((prev) => prev.map((p) => (p.id === activePageId ? { ...p, background: e.target.value } : p)))} />
                  </div>
                </div>
              )}

              {activeTab === "filtros" && (
                <div className="space-y-4">
                  <div className="rounded-xl border p-4" style={{ borderColor: palette.border, background: palette.panel }}>
                    <h3 className="mb-3 text-lg font-semibold">Barra de Filtros / Segmentacao</h3>
                    <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-5">
                      <div>
                        <label className="mb-1 block text-xs">Data Inicio</label>
                        <input type="date" className="w-full rounded border px-2 py-1" style={{ borderColor: palette.border }} value={filters.dateStart} onChange={(e) => setFilters((f) => ({ ...f, dateStart: e.target.value }))} />
                      </div>
                      <div>
                        <label className="mb-1 block text-xs">Data Fim</label>
                        <input type="date" className="w-full rounded border px-2 py-1" style={{ borderColor: palette.border }} value={filters.dateEnd} onChange={(e) => setFilters((f) => ({ ...f, dateEnd: e.target.value }))} />
                      </div>
                      <div>
                        <label className="mb-1 block text-xs">Categoria</label>
                        <select className="w-full rounded border px-2 py-1" style={{ borderColor: palette.border }} value={filters.category} onChange={(e) => setFilters((f) => ({ ...f, category: e.target.value }))}>
                          {categoryOptions.map((opt) => (<option key={opt}>{opt}</option>))}
                        </select>
                      </div>
                      <div>
                        <label className="mb-1 block text-xs">Regiao</label>
                        <select className="w-full rounded border px-2 py-1" style={{ borderColor: palette.border }} value={filters.regiao} onChange={(e) => setFilters((f) => ({ ...f, regiao: e.target.value }))}>
                          {regiaoOptions.map((opt) => (<option key={opt}>{opt}</option>))}
                        </select>
                      </div>
                      <div>
                        <label className="mb-1 block text-xs">Valor Maximo: {formatCurrency(filters.valorMax)}</label>
                        <input type="range" min={1000} max={100000} step={500} value={filters.valorMax} onChange={(e) => setFilters((f) => ({ ...f, valorMax: Number(e.target.value) }))} className="w-full" />
                      </div>
                    </div>
                  </div>
                  <div className="rounded-xl border p-4" style={{ borderColor: palette.border, background: palette.panel }}>
                    <h3 className="mb-2 font-semibold">Filtros visuais rapidos (cross-filtering simplificado)</h3>
                    <div className="flex flex-wrap gap-2">
                      {categoryOptions.slice(1).map((cat) => (
                        <button key={cat} className="rounded border px-3 py-1 text-sm" style={{ borderColor: palette.border }} onClick={() => setCrossFilter({ field: "Categoria", value: cat })}>{cat}</button>
                      ))}
                      <button className="rounded border px-3 py-1 text-sm" style={{ borderColor: palette.border }} onClick={() => setCrossFilter(null)}>Limpar</button>
                    </div>
                    <p className="mt-3 text-sm" style={{ color: palette.textMuted }}>Registros apos filtros: {dataFiltered.length}</p>
                  </div>
                </div>
              )}

              {activeTab === "metricas" && (
                <div className="grid gap-4 lg:grid-cols-2">
                  <div className="rounded-xl border p-4" style={{ borderColor: palette.border, background: palette.panel }}>
                    <h3 className="mb-3 text-lg font-semibold">Criar Medidas (DAX simplificado)</h3>
                    <div className="space-y-2">
                      <input className="w-full rounded border px-3 py-2" style={{ borderColor: palette.border }} placeholder="Nome da medida" value={measureForm.nome} onChange={(e) => setMeasureForm((m) => ({ ...m, nome: e.target.value }))} />
                      <select className="w-full rounded border px-3 py-2" style={{ borderColor: palette.border }} value={measureForm.funcao} onChange={(e) => setMeasureForm((m) => ({ ...m, funcao: e.target.value as MeasureType }))}>
                        {(["SOMA", "MEDIA", "CONTAGEM", "MINIMO", "MAXIMO", "PORCENTAGEM", "VARIACAO", "ACUMULADO", "YOY"] as MeasureType[]).map((f) => (<option key={f}>{f}</option>))}
                      </select>
                      <select className="w-full rounded border px-3 py-2" style={{ borderColor: palette.border }} value={measureForm.campo} onChange={(e) => setMeasureForm((m) => ({ ...m, campo: e.target.value }))}>
                        <option value="">Selecione o campo</option>
                        {visibleColumns.map((c) => (<option key={c}>{c}</option>))}
                      </select>
                      <textarea className="h-24 w-full rounded border px-3 py-2 font-mono text-xs" style={{ borderColor: palette.border }} placeholder="Formula avancada opcional, ex: SUM * 1.15" value={measureForm.formula} onChange={(e) => setMeasureForm((m) => ({ ...m, formula: e.target.value }))} />
                      <button
                        className="rounded bg-cyan-400 px-4 py-2 font-semibold text-slate-900"
                        onClick={() => {
                          if (!measureForm.nome || !measureForm.campo) return;
                          setMeasures((prev) => [...prev, { id: crypto.randomUUID(), ...measureForm }]);
                          setMeasureForm({ nome: "", funcao: "SOMA", campo: "", formula: "" });
                        }}
                      >
                        Salvar medida
                      </button>
                    </div>
                  </div>
                  <div className="rounded-xl border p-4" style={{ borderColor: palette.border, background: palette.panel }}>
                    <h3 className="mb-3 text-lg font-semibold">Biblioteca de medidas</h3>
                    <ul className="space-y-2 text-sm">
                      {measures.map((m) => (
                        <li key={m.id} className="rounded border p-2" style={{ borderColor: palette.border }}>
                          <div className="flex items-center justify-between">
                            <strong>{m.nome}</strong>
                            <button onClick={() => setMeasures((prev) => prev.filter((item) => item.id !== m.id))}><X size={14} /></button>
                          </div>
                          <div style={{ color: palette.textMuted }}>{m.funcao} em {m.campo}</div>
                          <div className="mt-1">Preview: <span className="font-semibold text-cyan-300">{Number(runMeasure(m)).toLocaleString("pt-BR")}</span></div>
                        </li>
                      ))}
                    </ul>
                  </div>
                </div>
              )}

              {activeTab === "templates" && (
                <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
                  {templates.map((tpl) => (
                    <motion.div whileHover={{ scale: 1.02 }} key={tpl} className="rounded-xl border p-4" style={{ borderColor: palette.border, background: palette.panel }}>
                      <div className="mb-3 h-28 rounded bg-gradient-to-br from-cyan-500/25 to-fuchsia-500/25" />
                      <h3 className="font-semibold">{tpl}</h3>
                      <p className="mb-3 text-sm" style={{ color: palette.textMuted }}>Layout pre-montado com indicadores iniciais.</p>
                      <button
                        className="w-full rounded bg-cyan-400 py-2 text-sm font-semibold text-slate-900"
                        onClick={() => {
                          const page = buildDefaultPage(visibleColumns);
                          page.name = tpl;
                          if (tpl.includes("Financeiro")) page.widgets.push({ id: crypto.randomUUID(), type: "gauge", x: 60, y: 340, w: 300, h: 220, title: "Margem", color: "#00c853", xField: "Categoria", yField: "Lucro" });
                          if (tpl.includes("Marketing")) page.widgets.push({ id: crypto.randomUUID(), type: "scatter", x: 390, y: 340, w: 300, h: 220, title: "Conversao", color: "#7b2ff7", xField: "Quantidade", yField: "Lucro" });
                          if (tpl.includes("Executivo")) page.widgets.push({ id: crypto.randomUUID(), type: "kpi", x: 720, y: 340, w: 220, h: 220, title: "Receita Geral", color: "#00d4ff", xField: "Data", yField: "Valor_Total" });
                          setPages((prev) => [...prev, page]);
                          setActivePageId(page.id);
                          setActiveTab("editor");
                        }}
                      >
                        Usar este Template
                      </button>
                    </motion.div>
                  ))}
                </div>
              )}

              {activeTab === "guia" && (
                <div className="space-y-4">
                  <div className="rounded-xl border p-4" style={{ borderColor: palette.border, background: palette.panel }}>
                    <div className="flex flex-wrap items-center justify-between gap-2">
                      <h3 className="flex items-center gap-2 text-lg font-semibold"><GraduationCap size={18} /> Curso Completo de Power BI</h3>
                      <div className="flex items-center gap-4 text-sm">
                        <span className="inline-flex items-center gap-1"><Trophy size={14} /> <strong>{totalPoints} pts</strong></span>
                        <span className="inline-flex items-center gap-1"><Target size={14} /> {Math.round(progressByModule.reduce((acc, mod) => acc + mod.pct, 0) / progressByModule.length)}% geral</span>
                      </div>
                    </div>
                    <p className="mt-2 text-sm" style={{ color: palette.textMuted }}>
                      Curso passo a passo para iniciantes e intermediarios: aulas explicadas em linguagem simples, tarefas praticas e pontuacao por evolucao.
                    </p>
                    <input className="mt-3 w-full rounded border px-3 py-2" style={{ borderColor: palette.border }} placeholder="Buscar modulo, aula ou tema: DAX, Gateway, Star Schema" value={searchGuide} onChange={(e) => setSearchGuide(e.target.value)} />
                  </div>

                  <div className="grid gap-4 xl:grid-cols-[300px_1fr_260px]">
                    <div className="space-y-2 rounded-xl border p-3" style={{ borderColor: palette.border, background: palette.panel }}>
                      <h4 className="font-semibold">Trilha de Modulos</h4>
                      {learningModules
                        .filter((mod) => mod.titulo.toLowerCase().includes(searchGuide.toLowerCase()) || mod.descricao.toLowerCase().includes(searchGuide.toLowerCase()) || mod.aulas.some((aula) => aula.titulo.toLowerCase().includes(searchGuide.toLowerCase())))
                        .map((mod) => {
                          const progress = progressByModule.find((p) => p.id === mod.id);
                          return (
                            <button
                              key={mod.id}
                              className="w-full rounded-lg border p-3 text-left"
                              style={{ borderColor: selectedModule.id === mod.id ? "#00d4ff" : palette.border, background: selectedModule.id === mod.id ? "rgba(0,212,255,0.12)" : "transparent" }}
                              onClick={() => {
                                setSelectedModuleId(mod.id);
                                setGuideView("modulo");
                              }}
                            >
                              <div className="text-xs" style={{ color: palette.textMuted }}>{mod.nivel} • {mod.duracao}</div>
                              <p className="text-sm font-medium">{mod.titulo}</p>
                              <div className="mt-2 h-1.5 w-full rounded bg-white/10">
                                <div className="h-1.5 rounded bg-cyan-400" style={{ width: `${progress?.pct ?? 0}%` }} />
                              </div>
                            </button>
                          );
                        })}
                    </div>

                    <div className="space-y-3 rounded-xl border p-4" style={{ borderColor: palette.border, background: palette.panel }}>
                      <div>
                        <p className="text-xs uppercase tracking-wider" style={{ color: palette.textMuted }}>{selectedModule.nivel} • {selectedModule.duracao}</p>
                        <h4 className="text-xl font-semibold">{selectedModule.titulo}</h4>
                        <p className="mt-1 text-sm" style={{ color: palette.textMuted }}>{selectedModule.descricao}</p>
                      </div>

                      <div className="flex flex-wrap gap-2">
                        <button className="rounded border px-3 py-1 text-xs" style={{ borderColor: guideView === "modulo" ? "#00d4ff" : palette.border }} onClick={() => setGuideView("modulo")}>Pagina do modulo</button>
                        <button className="rounded border px-3 py-1 text-xs" style={{ borderColor: guideView === "aula" ? "#00d4ff" : palette.border }} onClick={() => setGuideView("aula")}>Pagina da aula</button>
                        <button className="rounded border px-3 py-1 text-xs" style={{ borderColor: guideView === "tarefa" ? "#00d4ff" : palette.border }} onClick={() => setGuideView("tarefa")}>Pagina da tarefa</button>
                        <button className="rounded border px-3 py-1 text-xs" style={{ borderColor: guideView === "quiz" ? "#00d4ff" : palette.border }} onClick={() => setGuideView("quiz")}>Pagina do quiz</button>
                      </div>

                      {guideView === "modulo" && (
                        <section className="space-y-4">
                          <div className="rounded-lg border p-4" style={{ borderColor: palette.border }}>
                            <h5 className="text-lg font-semibold">Visao geral completa</h5>
                            <p className="mt-2 text-sm" style={{ color: palette.textMuted }}>{moduleDeepDive[selectedModule.id]?.objetivo}</p>
                            <div className="mt-3 grid gap-3 md:grid-cols-2">
                              <div>
                                <p className="text-sm font-medium">Resultados esperados</p>
                                <ul className="mt-1 space-y-1 text-xs" style={{ color: palette.textMuted }}>
                                  {(moduleDeepDive[selectedModule.id]?.resultados ?? []).map((item) => (<li key={item}>- {item}</li>))}
                                </ul>
                              </div>
                              <div>
                                <p className="text-sm font-medium">Erros comuns para evitar</p>
                                <ul className="mt-1 space-y-1 text-xs" style={{ color: palette.textMuted }}>
                                  {(moduleDeepDive[selectedModule.id]?.erros ?? []).map((item) => (<li key={item}>- {item}</li>))}
                                </ul>
                              </div>
                            </div>
                            <div className="mt-3">
                              <p className="text-sm font-medium">Roteiro de estudo recomendado</p>
                              <ol className="mt-1 list-decimal space-y-1 pl-4 text-xs" style={{ color: palette.textMuted }}>
                                {(moduleDeepDive[selectedModule.id]?.roteiro ?? []).map((item) => (<li key={item}>{item}</li>))}
                              </ol>
                            </div>
                          </div>

                          <div className="rounded-lg border p-4" style={{ borderColor: palette.border }}>
                            <h6 className="font-semibold">Abrir paginas de aula</h6>
                            <div className="mt-2 grid gap-2 md:grid-cols-2">
                              {selectedModule.aulas.map((aula) => (
                                <button key={aula.id} className="rounded border p-2 text-left text-sm" style={{ borderColor: palette.border }} onClick={() => { setSelectedLessonId(aula.id); setGuideView("aula"); }}>{aula.titulo}</button>
                              ))}
                            </div>
                          </div>

                          <div className="rounded-lg border p-4" style={{ borderColor: palette.border }}>
                            <h6 className="font-semibold">Abrir paginas de tarefa</h6>
                            <div className="mt-2 grid gap-2 md:grid-cols-2">
                              {selectedModule.tarefas.map((task) => (
                                <button key={task.id} className="rounded border p-2 text-left text-sm" style={{ borderColor: palette.border }} onClick={() => { setSelectedTaskId(task.id); setGuideView("tarefa"); }}>{task.titulo}</button>
                              ))}
                            </div>
                          </div>
                        </section>
                      )}

                      {guideView === "aula" && selectedLesson && (
                        <section className="space-y-4">
                          <div className="flex flex-wrap gap-2">
                            {selectedModule.aulas.map((aula) => (
                              <button key={aula.id} className="rounded border px-2 py-1 text-xs" style={{ borderColor: selectedLesson.id === aula.id ? "#00d4ff" : palette.border }} onClick={() => setSelectedLessonId(aula.id)}>{aula.titulo}</button>
                            ))}
                          </div>
                          <article className="rounded-lg border p-5" style={{ borderColor: palette.border }}>
                            <h5 className="text-xl font-semibold">{selectedLesson.titulo}</h5>
                            <p className="mt-2 text-sm leading-relaxed" style={{ color: palette.textMuted }}>{selectedLesson.explicacao}</p>
                            <div className="mt-4 grid gap-4 md:grid-cols-2">
                              <div>
                                <h6 className="text-sm font-semibold">Passo a passo pratico</h6>
                                <ol className="mt-2 list-decimal space-y-2 pl-4 text-sm" style={{ color: palette.textMuted }}>
                                  {selectedLesson.passos.map((passo) => (<li key={passo}>{passo}</li>))}
                                </ol>
                              </div>
                              <div>
                                <h6 className="text-sm font-semibold">Aplicacao real</h6>
                                <p className="mt-2 text-sm" style={{ color: palette.textMuted }}>
                                  Simule esse conceito nos dados da TechStore Brasil e observe como ele muda a leitura de receita, lucro e tendencia.
                                  O objetivo e enxergar como cada tecnica impacta decisoes reais.
                                </p>
                              </div>
                            </div>
                            <button className="mt-4 rounded border px-3 py-1 text-sm" style={{ borderColor: palette.border }} onClick={() => saveLesson(selectedModule.id, selectedLesson.id)}>
                              {completedLessons.includes(`${selectedModule.id}:${selectedLesson.id}`) ? "Aula concluida" : "Marcar aula como concluida (+10 pts)"}
                            </button>
                          </article>
                        </section>
                      )}

                      {guideView === "tarefa" && selectedTask && (
                        <section className="space-y-4">
                          <div className="flex flex-wrap gap-2">
                            {selectedModule.tarefas.map((task) => (
                              <button key={task.id} className="rounded border px-2 py-1 text-xs" style={{ borderColor: selectedTask.id === task.id ? "#00d4ff" : palette.border }} onClick={() => setSelectedTaskId(task.id)}>{task.titulo}</button>
                            ))}
                          </div>
                          <article className="rounded-lg border p-5" style={{ borderColor: palette.border }}>
                            <div className="flex flex-wrap items-center justify-between gap-2">
                              <h5 className="text-xl font-semibold">{selectedTask.titulo}</h5>
                              <span className="text-sm text-cyan-300">{selectedTask.pontos} pontos</span>
                            </div>
                            <p className="mt-2 text-sm" style={{ color: palette.textMuted }}>{selectedTask.instrucoes}</p>
                            <p className="mt-2 text-xs" style={{ color: palette.textMuted }}>Dica: {selectedTask.dica}</p>
                            <div className="mt-3 rounded border p-3" style={{ borderColor: palette.border }}>
                              <p className="text-xs font-semibold">Criterio de avaliacao</p>
                              <ul className="mt-1 space-y-1 text-xs" style={{ color: palette.textMuted }}>
                                {(selectedTask.criterio ?? ["Resposta clara", "Aplicacao pratica", "Conclusao objetiva"]).map((item) => (<li key={item}>- {item}</li>))}
                              </ul>
                            </div>
                            {isWritingTask(selectedTask) && (
                              <div className="mt-3">
                                <label className="mb-1 block text-xs">Resposta da tarefa (minimo 60 caracteres)</label>
                                <textarea
                                  className="h-36 w-full rounded border px-3 py-2 text-sm"
                                  style={{ borderColor: palette.border }}
                                  value={taskResponses[`${selectedModule.id}:taskText:${selectedTask.id}`] ?? ""}
                                  onChange={(e) => setTaskResponses((prev) => ({ ...prev, [`${selectedModule.id}:taskText:${selectedTask.id}`]: e.target.value }))}
                                  placeholder="Escreva aqui sua resposta da tarefa..."
                                />
                              </div>
                            )}
                            <div className="mt-3 flex items-center gap-2">
                              <button className="rounded bg-cyan-400 px-3 py-1 text-xs font-semibold text-slate-900" onClick={() => toggleTask(selectedModule.id, selectedTask.id)}>
                                {completedTasks.includes(`${selectedModule.id}:task:${selectedTask.id}`) ? "Tarefa concluida" : "Concluir tarefa"}
                              </button>
                              <span className="text-xs" style={{ color: palette.textMuted }}>
                                {isWritingTask(selectedTask) ? "Tarefas de escrita exigem resposta no campo acima." : "Tarefa pratica pronta para validacao."}
                              </span>
                            </div>
                          </article>
                        </section>
                      )}

                      {guideView === "quiz" && (
                        <section className="space-y-4">
                          <article className="rounded-lg border p-5" style={{ borderColor: palette.border }}>
                            <h5 className="text-xl font-semibold">Quiz completo do modulo</h5>
                            <p className="mt-1 text-sm" style={{ color: palette.textMuted }}>Agora o quiz tem questoes mais aprofundadas com feedback explicativo.</p>
                            <div className="mt-3 space-y-3">
                              {selectedQuizQuestions.map((qz, idx) => {
                                const qKey = `${selectedModule.id}-${idx}`;
                                const answered = quizAnswers[qKey] ?? "";
                                const isCorrect = answered && answered.toLowerCase() === qz.a.toLowerCase();
                                return (
                                  <div key={qKey} className="rounded border p-3" style={{ borderColor: palette.border }}>
                                    <p className="text-sm font-medium">{idx + 1}. {qz.q}</p>
                                    <select
                                      className="mt-2 w-full rounded border px-2 py-1 text-sm"
                                      style={{ borderColor: palette.border }}
                                      value={answered}
                                      onChange={(e) => setQuizAnswers((prev) => ({ ...prev, [qKey]: e.target.value }))}
                                    >
                                      <option value="">Selecione uma opcao</option>
                                      {qz.opcoes.map((opcao) => (
                                        <option key={opcao} value={opcao}>{opcao}</option>
                                      ))}
                                    </select>
                                    {answered && (
                                      <p className="mt-2 text-xs" style={{ color: isCorrect ? "#00c853" : "#ffab00" }}>
                                        {isCorrect ? "Resposta correta." : `Resposta esperada: ${qz.a}.`} {qz.explicacao ?? ""}
                                      </p>
                                    )}
                                  </div>
                                );
                              })}
                            </div>

                            <div className="mt-4 rounded border p-3" style={{ borderColor: palette.border }}>
                              <p className="text-sm font-semibold">Pergunta aberta (obrigatoria para aprovacao)</p>
                              <p className="mt-1 text-xs" style={{ color: palette.textMuted }}>{openQuestionByModule[selectedModule.id]}</p>
                              <textarea
                                className="mt-2 h-28 w-full rounded border px-2 py-1 text-sm"
                                style={{ borderColor: palette.border }}
                                value={openQuizResponses[`${selectedModule.id}:open`] ?? ""}
                                onChange={(e) => setOpenQuizResponses((prev) => ({ ...prev, [`${selectedModule.id}:open`]: e.target.value }))}
                                placeholder="Escreva sua resposta discursiva aqui..."
                              />
                            </div>

                            <div className="mt-3 flex flex-wrap items-center gap-2">
                              <button className="rounded bg-cyan-400 px-3 py-1 text-xs font-semibold text-slate-900" onClick={() => submitQuiz(selectedModule.id)}>Validar quiz</button>
                              <span className="text-xs" style={{ color: palette.textMuted }}>
                                Resultado: {selectedQuizResult.correct}/{selectedQuizResult.total} ({selectedQuizResult.pct}%)
                              </span>
                            </div>
                            <p className="mt-2 text-xs" style={{ color: palette.textMuted }}>
                              Regra de aprovacao: minimo 70% no quiz e resposta aberta com ao menos 80 caracteres.
                            </p>
                            {passedQuizzes.includes(`${selectedModule.id}:quiz`) && (
                              <p className="mt-2 inline-flex items-center gap-1 text-xs text-emerald-400"><CheckCircle2 size={14} /> Quiz aprovado (+40 pts)</p>
                            )}
                          </article>
                        </section>
                      )}
                    </div>

                    <aside className="space-y-3">
                      <div className="rounded-xl border p-3" style={{ borderColor: palette.border, background: palette.panel }}>
                        <h4 className="mb-2 font-semibold">Glossario BI</h4>
                        <ul className="space-y-1 text-xs" style={{ color: palette.textMuted }}>
                          <li><strong>DAX:</strong> linguagem de expressoes analiticas.</li>
                          <li><strong>RLS:</strong> seguranca por linha.</li>
                          <li><strong>Star Schema:</strong> modelo fato/dimensao.</li>
                          <li><strong>Slicer:</strong> segmentador visual de filtro.</li>
                        </ul>
                      </div>
                      <div className="rounded-xl border p-3" style={{ borderColor: palette.border, background: palette.panel }}>
                        <h4 className="mb-2 font-semibold">Resumo de progresso</h4>
                        <p className="text-xs" style={{ color: palette.textMuted }}>Aulas concluidas: {completedLessons.length}</p>
                        <p className="text-xs" style={{ color: palette.textMuted }}>Tarefas concluidas: {completedTasks.length}</p>
                        <p className="text-xs" style={{ color: palette.textMuted }}>Quizzes aprovados: {passedQuizzes.length}</p>
                      </div>
                      <div className="rounded-xl border p-3" style={{ borderColor: palette.border, background: palette.panel }}>
                        <h4 className="mb-2 font-semibold">DAX Highlight</h4>
                        <pre className="overflow-auto rounded bg-black/30 p-2 text-xs text-cyan-300">
                          {daxExamples.join("\n")}
                        </pre>
                      </div>
                      <div className="rounded-xl border p-3" style={{ borderColor: palette.border, background: palette.panel }}>
                        <h4 className="mb-2 font-semibold">Diagrama visual</h4>
                        <svg viewBox="0 0 280 120" className="w-full">
                          <rect x="10" y="20" width="80" height="40" rx="6" fill="#00d4ff" opacity="0.75" />
                          <rect x="102" y="10" width="80" height="40" rx="6" fill="#7b2ff7" opacity="0.8" />
                          <rect x="194" y="20" width="80" height="40" rx="6" fill="#e94560" opacity="0.8" />
                          <line x1="90" y1="40" x2="102" y2="30" stroke="#fff" strokeWidth="2" />
                          <line x1="182" y1="30" x2="194" y2="40" stroke="#fff" strokeWidth="2" />
                          <text x="26" y="45" fill="#001">Fonte</text>
                          <text x="123" y="35" fill="#fff">Modelo</text>
                          <text x="220" y="45" fill="#fff">Dashboard</text>
                        </svg>
                      </div>
                    </aside>
                  </div>
                </div>
              )}

              {activeTab === "config" && (
                <div className="grid gap-4 lg:grid-cols-2">
                  <div className="space-y-3 rounded-xl border p-4" style={{ borderColor: palette.border, background: palette.panel }}>
                    <h3 className="text-lg font-semibold">Perfil do usuario</h3>
                    <input className="w-full rounded border px-3 py-2" style={{ borderColor: palette.border }} placeholder="Nome" value={profile.nome} onChange={(e) => setProfile((p) => ({ ...p, nome: e.target.value }))} />
                    <input className="w-full rounded border px-3 py-2" style={{ borderColor: palette.border }} placeholder="Cargo" value={profile.cargo} onChange={(e) => setProfile((p) => ({ ...p, cargo: e.target.value }))} />
                    <input className="w-full rounded border px-3 py-2" style={{ borderColor: palette.border }} placeholder="URL da foto" value={profile.foto} onChange={(e) => setProfile((p) => ({ ...p, foto: e.target.value }))} />
                    <div className="grid grid-cols-2 gap-2">
                      <button className="rounded border px-3 py-2" style={{ borderColor: palette.border }} onClick={() => setTheme((t) => (t === "dark" ? "light" : "dark"))}>Alternar Tema</button>
                      <button className="rounded border px-3 py-2" style={{ borderColor: palette.border }} onClick={() => setHighContrast((h) => !h)}>{highContrast ? "Desativar" : "Ativar"} alto contraste</button>
                      <label className="col-span-2 text-sm">Tamanho do texto</label>
                      <input className="col-span-2" type="range" min={0.9} max={1.2} step={0.05} value={textScale} onChange={(e) => setTextScale(Number(e.target.value))} />
                    </div>
                  </div>
                  <div className="space-y-3 rounded-xl border p-4" style={{ borderColor: palette.border, background: palette.panel }}>
                    <h3 className="text-lg font-semibold">Dados e Projeto</h3>
                    <button className="flex w-full items-center justify-center gap-2 rounded border px-3 py-2" style={{ borderColor: palette.border }} onClick={exportProjectJson}><Download size={16} /> Exportar projeto JSON</button>
                    <label className="flex cursor-pointer items-center justify-center gap-2 rounded border px-3 py-2" style={{ borderColor: palette.border }}>
                      <Upload size={16} /> Importar projeto JSON
                      <input type="file" accept=".json" className="hidden" onChange={importProjectJson} />
                    </label>
                    <button
                      className="flex w-full items-center justify-center gap-2 rounded border px-3 py-2"
                      style={{ borderColor: palette.border }}
                      onClick={() => {
                        localStorage.removeItem(STORAGE_KEY);
                        localStorage.removeItem(USER_KEY);
                        window.location.reload();
                      }}
                    >
                      Limpar cache e dados locais
                    </button>
                    <div className="rounded border p-3 text-sm" style={{ borderColor: palette.border }}>
                      <p><strong>Idioma:</strong> Portugues (Brasil)</p>
                      <p><strong>Versao:</strong> v1.0.0</p>
                      <p><strong>Atalhos:</strong> Ctrl+S salvar, Ctrl+1 Dashboard, Ctrl+2 Importar, Ctrl+3 Editor.</p>
                    </div>
                  </div>
                </div>
              )}
            </motion.main>
          </AnimatePresence>

          <footer className="border-t px-5 py-3 text-sm" style={{ borderColor: palette.border, color: palette.textMuted, background: palette.panel }}>
            Power BI Chrystopher © 2025 | Desenvolvido para aprendizado e analise de dados | Sobre | Ajuda | GitHub | v1.0.0
          </footer>
        </section>
      </div>
    </div>
  );
}
