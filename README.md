# Sistema de Mapeamento de Dores & Monitoramento Operacional

Aplicação web local integrada diretamente à sua base de dados Microsoft Excel (`Monitoramento_Dores - FINAL.xlsx`) com suporte completo a **anexos de arquivos**, **dashboard executivo com gráficos temporais e de Pareto**, **filtros por período**, **ordenação por colunas** e **exportação em CSV**.

---

## 🚀 Como Executar

O servidor já está **ativo** em:
👉 **[http://127.0.0.1:5000](http://127.0.0.1:5000)**

Para iniciar o sistema no futuro, basta:
1. Dar um duplo clique no arquivo: `iniciar_sistema.bat`
2. O sistema verifica automaticamente as dependências do Python e abre o navegador!

---

## 📎 Sistema de Anexos (Novo!)

### 1. No Momento do Cadastro:
- **Área Drag & Drop:** Arraste arquivos diretamente para a área indicada ou clique para selecionar.
- **Múltiplos Arquivos:** Suporta envio simultâneo de PDFs, imagens (PNG, JPG), planilhas (XLSX, CSV), notas fiscais (XML), documentos (DOCX) e arquivos compactados (ZIP) de até 50MB.
- **Pré-visualização:** Mostra a lista de arquivos selecionados com ícone, tamanho em KB/MB e botão para remover antes de salvar.

### 2. Na Consulta & Detalhes:
- **Indicador na Tabela:** Ícone de clipe com contador (ex.: `📎 2`).
- **Cards de Anexos no Modal de Detalhes:**
  - Miniaturas para imagens/fotos.
  - Botão **Visualizar** (abre em nova aba com visualizador nativo para PDF e imagens).
  - Botão **Baixar** (download direto).
  - Botão **Excluir** anexo.
  - Botão **"+ Anexar Arquivo"** para adicionar novas evidências a qualquer momento sem precisar recadastrar a demanda.
- **Sincronização com o Excel:** Os nomes dos arquivos anexados são salvos na **Coluna 31 (AE - Anexos)** da planilha Excel.

---

## 🌟 Principais Melhorias Implementadas

### 1. Dashboard Executivo Aprimorado
- **Filtro de Período Temporal:** Selecione entre `Todo o Histórico`, `Últimos 7 dias`, `Últimos 30 dias`, `Este Mês` ou intervalo `Personalizado` (com data inicial e final) recalculando KPIs e gráficos em tempo real.
- **Novo Gráfico: Evolução Temporal das Dores:** Linha do tempo mostrando a curva de ocorrências cadastradas versus concluídas por mês.
- **Novo Gráfico: Top 5 Maiores Impactos Financeiros (Pareto):** Identifica visualmente onde estão os maiores riscos financeiros.
- **Card KPI de Evidências:** Totalizador de arquivos anexados no sistema.

### 2. Gestão & Tabela de Consulta
- **Ordenação Clicável por Coluna:** Clique em qualquer cabeçalho da tabela (ID, Datas, Setor, Filial, Prioridade, Status, Valor, Anexos) para ordenar de forma crescente ou decrescente (▲ / ▼).
- **Filtro Exclusivo "Com Anexos":** Botão para exibir somente as pendências que possuem evidências anexadas.
- **Ação Rápida de Conclusão (✓):** Botão direto na tabela para marcar uma pendência como Concluída com 1 clique (gravando data de encerramento no Excel).
- **Paginação:** Seletor de 10, 25, 50 ou Todos os registros por página com navegação fluida.
- **Exportação CSV Estruturada:** Botão "Exportar CSV" que gera relatório com codificação UTF-8 BOM (compatível com acentuação no Excel brasileiro).
- **Impressão Executiva:** Botão "Imprimir Ficha" que formata uma ficha limpa para impressão ou PDF.

### 3. Atalhos de Teclado
- `Ctrl + K` ou `/`: Foca imediatamente na barra de pesquisa.
- `N`: Abre o modal de cadastro de nova pendência.
- `Esc`: Fecha qualquer modal aberto.

---

## 📁 Estrutura de Arquivos
```text
C:\Users\peewxx\.gemini\antigravity\scratch\gestao_dores\
  ├── app.py                     # Servidor Flask com endpoints REST e upload
  ├── excel_manager.py           # Gestor da planilha, fórmulas e anexos
  ├── iniciar_sistema.bat        # Inicializador rápido para Windows
  ├── requirements.txt           # Dependências Python
  ├── uploads/                   # Pasta com os arquivos anexados organizados por ID
  │   └── <id>/<arquivo>
  ├── templates/
  │   └── index.html             # Interface web completa
  ├── static/
  │   ├── css/
  │   │   ├── tailwind.min.css   # Framework CSS local
  │   │   └── style.css          # Estilos customizados e dropzones
  │   └── js/
  │       ├── chart.min.js       # Gráficos locais
  │       ├── lucide.min.js      # Ícones locais
  │       └── app.js             # Lógica e interatividade do sistema
  ├── data/                      # Cópia de trabalho da base Excel
  └── backups/                   # Histórico de backups automáticos
```
