# Plano de Produção - Sistema de Gerenciamento de Pedidos

Este é um sistema para gestão de planos de produção, com funcionalidades de carregamento de pedidos, definição de prioridades, validação de prazos e geração de relatórios.

## Índice

- [Recursos](#recursos)
- [Requisitos do Sistema](#requisitos-do-sistema)
- [Instalação](#instalação)
  - [Windows](#windows)
  - [macOS](#macos)
- [Como Usar](#como-usar)
- [Funcionalidades Principais](#funcionalidades-principais)
- [Configurações](#configurações)
- [Resolução de Problemas](#resolução-de-problemas)

## Recursos

- Interface de linha de comando interativa
- Carregamento de pedidos a partir de arquivos Excel
- Definição de tipos de corte para pedidos
- Priorização personalizada de pedidos
- Criação automática de planos de produção
- Validação de prazos
- Exportação de relatórios
- Configurações personalizáveis

## Requisitos do Sistema

- Python 3.6 ou superior
- Acesso à linha de comando (Terminal no macOS, Prompt de Comando ou PowerShell no Windows)

## Instalação

### Windows

1. **Instale o Python**:
   - Baixe o instalador mais recente do Python em [python.org](https://www.python.org/downloads/windows/)
   - Execute o instalador
   - **IMPORTANTE**: Marque a opção "Add Python to PATH" durante a instalação
   - Clique em "Install Now"

2. **Baixe o projeto**:
   - Baixe os arquivos do projeto para uma pasta em seu computador

3. **Abra o Prompt de Comando**:
   - Pressione `Win + R`, digite `cmd` e pressione Enter
   - Navegue até a pasta do projeto usando o comando `cd caminho\para\o\projeto`

4. **Configure o ambiente virtual** (recomendado):
   ```cmd
   python -m venv venv
   venv\Scripts\activate
   ```

5. **Instale as dependências**:
   ```cmd
   pip install -r requirements.txt
   ```

### macOS

1. **Instale o Python** (caso já não tenha):
   - Muitos sistemas macOS já vêm com Python, mas recomenda-se a instalação da versão mais recente
   - Usando Homebrew: `brew install python`
   - Ou baixe o instalador em [python.org](https://www.python.org/downloads/mac-osx/)

2. **Baixe o projeto**:
   - Baixe os arquivos do projeto para uma pasta em seu computador

3. **Abra o Terminal**:
   - Abra o aplicativo Terminal
   - Navegue até a pasta do projeto usando o comando `cd caminho/para/o/projeto`

4. **Configure o ambiente virtual** (recomendado):
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   ```

5. **Instale as dependências**:
   ```bash
   pip3 install -r requirements.txt
   ```

## Como Usar

1. **Ative o ambiente virtual** (se estiver usando):
   - Windows: `venv\Scripts\activate`
   - macOS: `source venv/bin/activate`

2. **Execute o programa**:
   - Windows: `python main.py`
   - macOS: `python3 main.py`

3. **Menu Interativo**:
   - Use as setas do teclado para navegar pelo menu
   - Pressione Enter para selecionar uma opção

## Funcionalidades Principais

### 📥 Carregar Pedidos
- Permite selecionar um arquivo Excel contendo a lista de pedidos
- Define tipos de corte para cada pedido
- Estabelece prioridades para processamento
- Valida prazos e cria um plano de produção

### 📊 Exportar Relatórios
- Gera relatórios baseados em dados de produção
- Permite escolher o arquivo para exportação

### ⚙️ Configurações
- Visualiza e altera parâmetros do sistema
- Alterações são salvas automaticamente

### 🚪 Sair
- Encerra o programa com segurança

## Configurações

O sistema utiliza um arquivo de configuração (CSV com codificação UTF-16) para armazenar parâmetros. Os parâmetros podem ser:
- Valores numéricos
- Opções "Sim/Não"

Para alterar as configurações, selecione a opção "⚙️ Configurações" no menu principal.

## Resolução de Problemas

**Erro ao carregar arquivo de configuração**:
- Verifique se o arquivo de configuração existe no caminho padrão
- Verifique se o formato do arquivo está correto (CSV com codificação UTF-16)

**Erro ao carregar arquivo Excel**:
- Certifique-se de que o formato do arquivo é compatível
- Verifique se o arquivo não está aberto em outro programa

**Falha na instalação de dependências**:
- Windows: Certifique-se de que o Python está no PATH do sistema
- macOS: Use `pip3` em vez de `pip` se estiver usando Python 3.x
- Se o arquivo requirements.txt estiver ausente, instale as dependências manualmente:
  ```
  pip install pyfiglet tabulate InquirerPy pandas
  ```

Se os problemas persistirem, verifique os requisitos de sistema e tente reinstalar as dependências.