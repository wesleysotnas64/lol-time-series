# lol-time-series

## Configuração do ambiente virtual

O ambiente virtual permite instalar as dependências do projeto de forma isolada, evitando conflitos com outros projetos ou com os pacotes instalados no sistema operacional.

### 1. Windows

Os comandos abaixo devem ser executados no terminal do VS Code, utilizando o PowerShell.

**1.1. Acessar a raiz do projeto**

```powershell
cd caminho\para\lol-time-series
```

**1.2. Criar o ambiente virtual**

```powershell
python -m venv .venv
```

**1.3. Ativar o ambiente virtual**

```powershell
.\.venv\Scripts\Activate.ps1
```

Caso o PowerShell bloqueie a execução de scripts, execute o comando abaixo para permitir a ativação apenas na sessão atual do terminal:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Em seguida, tente ativar o ambiente novamente:

```powershell
.\.venv\Scripts\Activate.ps1
```

**1.4. Instalar as dependências**

Com o ambiente virtual ativado, execute:

```powershell
python -m pip install --upgrade pip
python -m pip install ipykernel numpy pandas matplotlib
```

Esses pacotes permitem executar notebooks Jupyter e realizar operações de análise e visualização de séries temporais.

**1.5. Registrar o kernel do Jupyter**

```powershell
python -m ipykernel install --user --name lol-time-series --display-name "Python (lol-time-series)"
```

No VS Code, abra o notebook desejado, clique em **Select Kernel** e selecione **Python (lol-time-series)**.

### 2. Ubuntu

Os comandos abaixo devem ser executados no terminal do sistema ou no terminal integrado do VS Code.

**2.1. Acessar a raiz do projeto**

```bash
cd /caminho/para/lol-time-series
```

**2.2. Instalar o suporte a ambientes virtuais**

Caso ainda não esteja instalado, execute:

```bash
sudo apt update
sudo apt install python3-venv python3-pip
```

**2.3. Criar o ambiente virtual**

```bash
python3 -m venv .venv
```

**2.4. Ativar o ambiente virtual**

```bash
source .venv/bin/activate
```

Após a ativação, o terminal deverá exibir `(.venv)` no início da linha.

**2.5. Instalar as dependências**

```bash
python -m pip install --upgrade pip
python -m pip install ipykernel numpy pandas matplotlib
```

**2.6. Registrar o kernel do Jupyter**

```bash
python -m ipykernel install --user --name lol-time-series --display-name "Python (lol-time-series)"
```

No VS Code, abra o notebook, clique em **Select Kernel** e selecione **Python (lol-time-series)**.

### 3. Boas práticas

- Crie um ambiente virtual separado para cada sistema operacional. O ambiente `.venv` do Windows não deve ser compartilhado com o Ubuntu.
- Não versione a pasta `.venv` no Git. Cada máquina deve criar seu próprio ambiente virtual.
- Para sair do ambiente virtual, execute:

```bash
deactivate
```

No Windows e no Ubuntu, o comando `deactivate` funciona após a ativação do ambiente.

- Para verificar qual interpretador Python está sendo utilizado:

No Windows:

```powershell
python -c "import sys; print(sys.executable)"
```

No Ubuntu:

```bash
python -c "import sys; print(sys.executable)"
```

O caminho apresentado deve apontar para o interpretador dentro da pasta `.venv`.

### 4. Arquivo `.gitignore`

Adicione as seguintes linhas ao arquivo `.gitignore`, localizado na raiz do projeto, para evitar o envio do ambiente virtual e de arquivos temporários ao GitHub:

```gitignore
# Ambientes virtuais
.venv/
venv/

# Cache do Python
__pycache__/
*.py[cod]

# Checkpoints do Jupyter
.ipynb_checkpoints/
```

Com essas configurações, o projeto poderá ser desenvolvido no Windows e no Ubuntu, utilizando ambientes virtuais independentes e mantendo o código sincronizado pelo GitHub.