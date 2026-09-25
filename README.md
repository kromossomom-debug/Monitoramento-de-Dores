# Sistema de Mapeamento de Dores & Monitoramento Operacional

Aplicação web local integrada diretamente à sua base de dados Microsoft Excel (`Monitoramento_Dores - FINAL.xlsx`).

---

## 🚀 Como Executar

O servidor já está **ativo** em:
👉 **[http://127.0.0.1:5000](http://127.0.0.1:5000)**

Para iniciar o sistema no futuro, basta:
1. Dar um duplo clique no arquivo: `iniciar_sistema.bat`
2. O navegador abrirá automaticamente com o painel pronto para uso!

---

## 📊 Principais Funcionalidades

### 1. Dashboard Executivo & Gráficos
- **8 Cards de Indicadores (KPIs):**
  - Total de Ocorrências
  - Ocorrências Abertas / Em Tratativa
  - Alerta de SLAs Vencidos
  - Alerta de SLAs A Vencer (&le; 2 dias)
  - Ocorrências Concluídas e Taxa de Resolução (%)
  - Valor Financeiro Total das Notas (R$)
  - Tempo Médio em Aberto (dias)
  - Tempo Médio de Resolução (dias)
- **6 Gráficos Interativos (Chart.js):**
  - Distribuição por Status (Aberto, Em Análise, Em Tratativa, Aguardando Terceiros, Concluído, Cancelado)
  - Situação de SLA (No prazo, A vencer, Vencido, Concluído)
  - Volume de Dores por Categoria (Fiscal, Estoque, Financeiro, Faturamento, Logística, etc.)
  - Setor Responsável atribuído
  - Distribuição por Criticidade / Prioridade
  - Volume por Filial (Matupá, Rondonópolis, Paranaguá)

### 2. Cadastro de Nova Pendência
- Formulário intuitivo estruturado em 4 blocos:
  1. **Dados Gerais:** Datas, Responsável, Filial, Setor Responsável, Setor Impactado e Categoria da Dor.
  2. **Problema & Criticidade:** Descrição detalhada, Causa Raiz, Impacto no Negócio, Prioridade (com cálculo automático de SLA: Crítica = 2d, Alta = 5d, Média = 10d, Baixa = 15d) e Recorrência.
  3. **Dados Fiscais & Operacionais:** Nota Fiscal, Série, Qtd em kg, Valor em R$ e Nº do Chamado.
  4. **Plano de Ação & Solução:** Status inicial, Plano de Ação, Prazo acordado e Responsável pela Solução.
- **Gravação Direta no Excel:** Cada pendência cadastrada é gravada imediatamente na linha correta da aba `Monitoramento de Dores`, mantendo todas as fórmulas de cálculo do Excel intactas (`ID`, `Dias em Aberto`, `SLA (dias)`, `Data Limite SLA`, `Situação SLA`, `Dias para Vencimento`).

### 3. Consulta & Gestão Interativa
- **Pesquisa Instantânea:** Filtro textual em tempo real por descrição, setor, responsável, nota fiscal, número de chamado, etc.
- **Filtros Rápidos:** Botões de 1 clique para `Abertos`, `SLA Vencido`, `SLA A Vencer`, `Prioridade Crítica` e `Concluídos`.
- **Filtros Combinados:** Por Status, Prioridade, Filial e SLA.
- **Visualização Completa de Detalhes:** Modal detalhado com todos os 30 campos da base.
- **Tratamento & Conclusão:** Permite alterar o status para `Concluído` (preenchendo a data de conclusão automaticamente), atualizar o plano de ação, adicionar observações e salvar direto no Excel.

### 4. Segurança e Integridade da Base Excel
- **Caminho Ativo:** `C:\Users\peewxx\Downloads\Monitoramento_Dores - FINAL.xlsx`
- **Backups Automáticos:** Antes de qualquer alteração, uma cópia com carimbo de data/hora é salva na pasta `backups/`.
- **Botão "Abrir no Excel":** Abre o arquivo nativamente no Microsoft Excel.
- **Botão "Baixar Excel":** Permite fazer o download da planilha atualizada diretamente pelo navegador.
- **Totalmente Offline:** Todos os scripts de gráficos e estilos estão salvos localmente.

---

## 📁 Estrutura do Projeto
```text
C:\Users\peewxx\.gemini\antigravity\scratch\gestao_dores\
  ├── app.py                     # Servidor Flask com endpoints REST
  ├── excel_manager.py           # Leitura, escrita, fórmulas e backups da planilha
  ├── iniciar_sistema.bat        # Inicializador rápido para Windows
  ├── requirements.txt           # Dependências Python
  ├── templates/
  │   └── index.html             # Interface web moderna SPA
  ├── static/
  │   ├── css/
  │   │   ├── tailwind.min.css   # Framework CSS local
  │   │   └── style.css          # Estilos personalizados e badges
  │   └── js/
  │       ├── chart.min.js       # Gráficos interativos locais
  │       ├── lucide.min.js      # Ícones locais
  │       └── app.js             # Lógica e interatividade do sistema
  ├── data/                      # Cópia sincronizada da base
  └── backups/                   # Histórico de backups automáticos
```
