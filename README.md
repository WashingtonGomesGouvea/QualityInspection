
## 📋 Formulário Dinâmico de Inspeção de Qualidade (Streamlit) ✨

Esta aplicação Streamlit fornece um sistema de formulários dinâmicos para inspeções de qualidade, concebido para ser modular, escalável e com foco em dispositivos móveis 📱. Integra-se com o SharePoint para carregar rotinas de inspeção e guardar relatórios ☁️.

## ✨ Funcionalidades Principais

*   **📝 Geração Dinâmica de Formulários:** Os formulários de inspeção são renderizados dinamicamente com base num ficheiro de configuração JSON (`roteiros_final_v4.json`).
*   **🔗 Integração com SharePoint:**
    *   🔄 Carrega as configurações das rotinas de inspeção primariamente a partir de um URL do SharePoint especificado.
    *   💾 Guarda os relatórios de inspeção concluídos (em formato Excel) num diretório designado do SharePoint, anexando a ficheiros existentes ou criando novos se não existirem.
*   **💻 Fallback Local e Download:**
    *   📄 Utiliza um ficheiro JSON local para as rotinas caso o SharePoint esteja indisponível.
    *   📊 Permite aos utilizadores descarregar um relatório Excel de todas as inspeções concluídas durante a sessão atual.
*   **⏳ Cálculo Reativo de Validade da Solução:** Calcula e exibe automaticamente a data de validade para soluções com base na data de preparação e no tipo de solução. Este cálculo é atualizado reativamente à medida que o utilizador altera os campos de entrada relevantes (a data de preparação e o tipo de solução são colocados fora do formulário principal para permitir esta reatividade).
*   **🔒 Gestão Segura de Credenciais:** Utiliza a gestão de segredos incorporada do Streamlit (`.streamlit/secrets.toml`) para as credenciais do SharePoint.
*   **🎨 Interface Amigável:**
    *   👤 Barra lateral para dados iniciais do inspetor, setor de inspeção e seleção de processo.
    *   🎛️ Suporta uma vasta variedade de tipos de campos: texto, data, seleção (dropdown), multisseleção, caixas de verificação, botões de rádio e múltiplos uploads de ficheiros.
    *   👁️ Exibição condicional de campos do formulário com base numa lógica "show_if" ligada a valores de outros campos.
*   **✔️ Validação de Dados:** Valida campos obrigatórios antes de permitir a submissão do formulário.
*   **📄 Saída Excel Padronizada:** Gera relatórios Excel com uma estrutura consistente, adequados para análise de dados e integração com Business Intelligence (BI).
*   **📱 Design Mobile-First:** Construído com um layout responsivo para usabilidade em vários dispositivos.

## 🚀 Pré-requisitos

*   🐍 Python 3.8+
*   📦 Pip (instalador de pacotes Python)

## 🛠️ Instalação e Configuração

1.  **📥 Clonar o repositório (ou descarregar os ficheiros):**
    ```bash
    # Se tiver um repositório git
    # git clone <repository-url>
    # cd <repository-directory>
    ```

2.  **🌿 Criar um ambiente virtual (recomendado):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # No Windows use `venv\Scripts\activate`
    ```

3.  **🧩 Instalar os pacotes Python necessários:**
    ```bash
    pip install -r requirements.txt
    ```
    O ficheiro `requirements.txt` deve conter:
    ```
    streamlit
pandas
openpyxl
Office365-REST-Python-Client
    ```

4.  **🔑 Configuração:**

    *   **🔒 Credenciais do SharePoint (`secrets.toml`):**
        Crie um ficheiro chamado `secrets.toml` num diretório `.streamlit` na raiz do seu projeto (ou seja, `.streamlit/secrets.toml`). Adicione as suas credenciais e URLs do SharePoint a este ficheiro:
        ```toml
        [sharepoint]
        email = "your_sharepoint_email@example.com"
        password = "your_sharepoint_password"
        site_url = "https://yourtenant.sharepoint.com/sites/your_site_or_personal_site/"

        # URL para o ficheiro JSON que contém as rotinas de inspeção no SharePoint
        # Este deve ser um link de download direto ou um link do qual o UniqueId possa ser extraído.
        roteiros_file_url = "https://yourtenant.sharepoint.com/personal/your_user/_layouts/15/download.aspx?UniqueId=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx&e=xxxxxx"

        # (Opcional) Nome do ficheiro para o histórico consolidado de inspeções em Excel no SharePoint
        # O padrão é "registos_inspecoes_V2.xlsx" se não for especificado.
        # Este ficheiro será armazenado em "Documents/Inspeção Qualidade/" relativo ao site_url.
        historico_inspecoes_filename = "nome_do_seu_ficheiro_de_historico.xlsx"
        ```
        **⚠️ Nota de Segurança Importante:** Certifique-se de que o ficheiro `secrets.toml` está incluído no seu ficheiro `.gitignore` se estiver a usar Git, para evitar a exposição acidental de credenciais.

    *   **📝 Ficheiro de Rotinas de Inspeção (`roteiros_final_v4.json`):**
        *   **☁️ SharePoint (Primário):** Carregue o seu ficheiro `roteiros_final_v4.json` para uma localização no SharePoint e atualize o `roteiros_file_url` em `secrets.toml` para apontar para o seu link de download direto.
        *   **💻 Local Fallback (Secundário):** Coloque o ficheiro `roteiros_final_v4.json` no mesmo diretório que `QualityInspection_V2.py` (ou atualize `ROTEIROS_LOCAL_PATH` no script se o colocar noutro local). Este ficheiro local será usado se a versão do SharePoint não puder ser acedida.

        A estrutura deste ficheiro JSON define os campos iniciais do inspetor, setores, processos e os campos específicos para cada processo de inspeção. Elementos chave incluem:
        *   `informacoes_iniciais`: Campos para detalhes do inspetor (nome, email, empresa, data da inspeção).
        *   `setores_inspecao`: Uma lista de setores, cada um contendo:
            *   `nome`: Nome de exibição do setor.
            *   `key`: Identificador único para o setor.
            *   `processos`: Uma lista de processos dentro desse setor, cada um contendo:
                *   `nome`: Nome de exibição do processo.
                *   `key`: Identificador único para o processo.
                *   `campos`: Uma lista de definições de campos para o formulário de inspeção, incluindo `label`, `key`, `tipo` (texto, data, seleção, etc.), `obrigatorio`, `opcoes` (para tipos de seleção) e `condicional` (para exibição condicional).
        *   `regras_validade_solucoes`: Define regras para calcular as datas de validade das soluções (usado pelo processo "Soluções").

## 🏃 Como Executar a Aplicação

1.  Certifique-se de que o seu ambiente virtual está ativado (se criou um).
2.  Navegue para o diretório que contém `QualityInspection_V2.py`.
3.  Execute a aplicação Streamlit a partir do seu terminal:
    ```bash
    streamlit run QualityInspection_V2.py
    ```
4.  O Streamlit normalmente abrirá a aplicação automaticamente no seu navegador web padrão. Caso contrário, exibirá um URL local (por exemplo, `http://localhost:8501`) que pode abrir manualmente.

## 📖 Como Utilizar a Aplicação

1.  **👤 Barra Lateral:** Preencha os detalhes do inspetor e selecione o setor de inspeção e o processo na barra lateral.
2.  **🔄 Carregar Roteiro:** Clique em "Carregar Roteiro de Inspeção".
3.  **📝 Formulário Dinâmico:** A área principal exibirá o formulário de inspeção correspondente ao processo selecionado.
    *   Para o processo "Soluções": Os campos "Data de Preparo da Solução" e "Tipo da Solução" aparecerão acima do formulário principal. A alteração destes atualizará reativamente o campo "Data de Validade da Solução (Calculada)" mostrado no formulário.
4.  **✍️ Preencher Formulário:** Complete todos os campos obrigatórios. Carregue ficheiros de evidência quando necessário.
5.  **✅ Submeter:** Clique em "Finalizar e Submeter Inspeção".
    *   Os dados da inspeção serão guardados no SharePoint (se configurado e acessível).
    *   A inspeção será adicionada a uma lista de inspeções realizadas na sessão atual.
6.  **📊 Exportar Dados da Sessão:** Pode descarregar um ficheiro Excel de todas as inspeções realizadas na sessão atual usando o botão "Download Excel da Sessão" na barra lateral.

## 🤔 Resolução de Problemas (Troubleshooting)

*   **🔌 Problemas de Conexão com o SharePoint:** Verifique novamente o seu email, palavra-passe e `site_url` em `secrets.toml`. Certifique-se de que o utilizador tem as permissões necessárias para o site do SharePoint e a biblioteca de documentos (`Documents/Inspeção Qualidade/`).
*   **🗺️ Roteiros Não Carregam:** Verifique se o `roteiros_file_url` em `secrets.toml` está correto e acessível. Certifique-se de que o ficheiro local `roteiros_final_v4.json` existe no caminho correto se o acesso ao SharePoint falhar.
*   **💾 Erros de Exportação/Gravação do Excel:** Certifique-se de que `openpyxl` está instalado. Verifique as permissões do SharePoint se a gravação no SharePoint falhar.

---

Este README fornece um guia abrangente para configurar e executar a aplicação Formulário Dinâmico de Inspeção de Qualidade. Para personalização ou desenvolvimento adicional, consulte os comentários no script `QualityInspection_V2.py` e a estrutura do ficheiro de configuração `roteiros_final_v4.json`.

