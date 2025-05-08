# 🖨️ Planejamento de Produção para Gráfica

Este projeto tem como objetivo auxiliar no planejamento de produção para gráfica, utilizando uma planilha base personalizada e Python para processar dados do fluxo produtivo.

## 📁 Estrutura do Projeto

- `main.py`: Script principal que executa o planejamento.
- `requirements.txt`: Lista de dependências Python necessárias.
- `data/`: Diretório destinado a arquivos de entrada, configurado pelo usuário.
- `automation/`: Módulos e funções auxiliares de automação utilizadas no projeto.
- `exp/`: Pasta que armazena os arquivos exportados pelo script
- `.gitignore`: Arquivos e pastas ignorados pelo controle de versão.

## ✅ Pré-requisitos

- **Python 3.8 ou superior**: Certifique-se de que o Python está instalado. Você pode verificar a versão com:

  ```bash
  python --version
  ```

- **pip**: Gerenciador de pacotes do Python. Normalmente já vem com o Python.

- **Git**: Para clonar o repositório.

## 💻 Instalação

### 1. Clone o repositório

Abra o terminal (Mac) ou Prompt de Comando (Windows) e execute:

```bash
git clone https://github.com/danielcscruz/planejamento_producao_grf.git
cd planejamento_producao_grf
```

### 2. Crie um ambiente virtual (opcional, mas recomendado)

#### Mac/Linux:

```bash
python3 -m venv venv
source venv/bin/activate
```

#### Windows:

```bash
python -m venv venv
venv\Scripts\activate
```

### 3. Instale as dependências

```bash
pip install -r requirements.txt
```

## 🚀 Execução

Após instalar as dependências, execute o script principal:

```bash
python main.py
```

⚠️ Certifique-se de que a planilha de planejamento .xlsx  esteja no mesmo diretório que 'main.py', conforme esperado pelo script.

## 🧪 Testes

Atualmente, o projeto não possui testes automatizados. Recomenda-se a criação de testes para validar as funcionalidades à medida que o projeto evolui.

## 📄 Licença

Este projeto está sob a licença MIT. Consulte o arquivo `LICENSE` para mais informações.
