# Separador de Casos por Área

Automação em Python para separar e consolidar arquivos Excel por área, desenvolvida por **Giovanni Moreira**, Assistente Financeiro na Allcare Gestora de Saúde.

---

## 📋 Visão Geral

Este projeto fornece duas partes principais:

* **Main.py** (versão 1.6.2): script principal que:

  * Solicita ao usuário a pasta onde os arquivos devem ser processados.
  * Carrega um arquivo Excel de referência.
  * Ordena e filtra seus dados.
  * Separa o conteúdo em múltiplos arquivos por área (coluna especificada).
  * Consolida todos os arquivos gerados em um único `consolidado.xlsx`.
  * Integra a biblioteca local **Separador.py** para mover arquivos em pastas nomeadas por área.

* **Separador.py** (versão 1.0.2): módulo auxiliar que:

  * Identifica sufixos no nome dos arquivos para mapear cada item a uma pasta específica.
  * Cria as pastas numeradas (ex.: `1 - FABRICA`, `2 - LOJA`, `3 - JOALHERIA`).
  * Move os arquivos para suas respectivas pastas.

---

## 🚀 Tecnologias Utilizadas

* **Python 3.x**
* **openpyxl**: leitura e escrita de arquivos `.xlsx`
* **tkinter**: diálogo de seleção de pastas
* **os** e **shutil**: operações de sistema de arquivos
* **Workbook** (do `openpyxl`)

---

## 📂 Estrutura do Repositório

```
├── Main.py          # Script principal
├── Separador.py     # Biblioteca local para mover arquivos por área
└── README.md        # Documentação do projeto
```

---

## 🛠 Instalação

1. Clone este repositório:

   ```bash
   git clone https://github.com/seu-usuario/seu-repo.git
   cd seu-repo
   ```
2. Instale as dependências:

   ```bash
   pip install openpyxl
   ```

   (o `tkinter` já vem incluso na distribuição padrão do Python)

---

## 🎯 Como Utilizar

1. Abra o arquivo `Main.py` e ajuste, se necessário:

   * O caminho do arquivo de referência Excel: `wb = openpyxl.load_workbook(r'Caminho/do/arquivo/para/separar.xlsx')`
   * A coluna base para separar os casos: parâmetro `col_base` em `preparar()` e `separar()` (ex.: `'AH'`).
2. Execute o script:

   ```bash
   python Main.py
   ```
3. Selecione a pasta alvo quando solicitado.
4. Aguarde enquanto o script:

   * Separa o arquivo Excel em várias planilhas por área.
   * Consolida todos os arquivos gerados em `consolidado.xlsx`.
   * Move os arquivos separados em pastas específicas via `Separador.py`.

---

## 🤝 Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para:

* Abrir *issues* para bugs e sugestões.
* Enviar *pull requests* aprimorando funcionalidades.

---

## 📞 Contato

**Giovanni Moreira**
Assistente Financeiro na Allcare Gestora de Saúde

💼 LinkedIn: [Giovanni Moreira](https://www.linkedin.com/in/giovanni-moreira-64654a254/)

✉️ Email: [diovanni.2566@gmail.com](mailto:diovanni.2566@gmail.com)

---

*Versão do projeto: **1.6.2***
*Data: 09/05/2025*
