# VBA Excel/Word Report Automation Engine 🚀

Solução avançada de automação para a geração de relatórios institucionais, integrando o Microsoft Excel como banco de dados e o Microsoft Word como motor de renderização de documentos.

## 📑 Visão Geral

O projeto foi desenvolvido para eliminar o erro humano e o trabalho manual na consolidação de dados de auditoria. Ele utiliza a técnica de **Template Tagging**, onde um documento `.docx` serve como base e o VBA preenche as informações dinamicamente, mantendo a integridade visual do relatório original.



## 🛠️ Arquitetura do Código

O código está estruturado em módulos funcionais para garantir escalabilidade:

### 1. Motor de Integração (`Motor_Gerador_Relatorio`)
* **Late Binding:** Não requer ativação de bibliotecas (References) manuais, garantindo que o código funcione em qualquer versão do Office.
* **Multi-Source Data:** Consolida dados de três fontes diferentes (`Main_Database`, `Findings_Data`, `Structural_Data`).

### 2. Funções de Inteligência em Tabelas (`ColarTabelaWord`)
Diferente de uma colagem simples, este motor:
* **Auto-Ajuste:** Redimensiona tabelas automaticamente para a largura da página (Window Fit).
* **Limpeza de Dados:** Executa uma varredura para excluir linhas onde os valores são zero, nulos ou apenas hífens, garantindo que apenas dados relevantes apareçam.
* **Alinhamento Dinâmico:** Aplica regras específicas por coluna (ex: descrições à esquerda, valores centralizados).

### 3. Tratamento de Strings e Texto
* **`CleanWordCell`**: Remove caracteres de controle invisíveis do Word (Chr 7, 13, 160) que frequentemente corrompem lógicas de comparação.
* **`ApagarBlocoSeguro`**: Remove seções inteiras do documento se não houver dados correspondentes, evitando parágrafos órfãos ou títulos sem conteúdo.

## ☁️ Cloud & OneDrive Synchronization
Um dos maiores desafios da automação Office é a latência de sincronização em nuvem. Este projeto implementa um **Salvamento Robusto**:
* Uso de `DoEvents` para liberar o processamento do SO.
* `Application.Wait` estratégico para garantir que o Word finalize a escrita em disco antes de fechar a instância, evitando corrupção de arquivos em ambientes OneDrive/SharePoint.

## 📂 Como Configurar

1.  Mantenha os arquivos na mesma pasta conforme a nomenclatura:
    * `Main_Database.xlsm` (Controle central)
    * `Findings_Data.xlsm` (Dados de apontamentos)
    * `Structural_Data.xlsm` (Dados cadastrais)
    * `Report_Template.docx` (Modelo com as tags `[...]`)
2.  Abra o `Main_Database.xlsm`, habilite as Macros e execute via interface.

## 🔍 Tags Suportadas (Exemplos)
* `[QUADROS_NOME]`: Nome da entidade.
* `[QUADROS_TABELA_ESTOQUE]`: Inserção de tabela dinâmica.
* `[APONT_3.3_FRASE]`: Lógica de texto condicional.

---
**Tecnologias:** VBA (Excel/Word Object Model), Microsoft Office.  
**Foco:** Produtividade, Compliance, Automação de Backoffice.

## 📄 Licença

Este projeto está sob a licença MIT. Consulte o arquivo `LICENSE` para mais detalhes.
