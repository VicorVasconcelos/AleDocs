# AleDocs - Sistema de Geracao de Documentos Juridicos e de Agrimensura

Sistema web local desenvolvido em Python + Flask para automatizar a geracao de documentos juridicos e tecnicos de georreferenciamento, com banco de dados SQLite.

---

## Visao Geral

O AleDocs permite:

- Cadastro e edicao de clientes
- Cadastro e edicao de emissores dos documentos
- Selecao de cliente + emissor no momento da geracao
- Geracao de 7 documentos em `.docx` e download em `.zip`
- Personalizacao de layout e conteudo dos modelos Word
- Execucao local e publicacao externa via ngrok

---

## Pre-requisitos

- Python 3.9+
- Ambiente Windows, Linux ou macOS
- Acesso a internet para bibliotecas via CDN (Bootstrap/Select2)

---

## Instalacao

Abra um terminal na pasta do projeto e execute:

### 1) Criar ambiente virtual

```bash
python -m venv .venv
```

### 2) Ativar ambiente virtual

Windows (PowerShell):

```powershell
.\.venv\Scripts\Activate.ps1
```

Linux/macOS:

```bash
source .venv/bin/activate
```

### 3) Instalar dependencias

```bash
pip install -r requirements.txt
```

### 4) Gerar modelos `.docx`

```bash
python create_templates.py
```

Saida esperada (resumo):

- termo_entrega.docx
- checklist_averbacao.docx
- contrato_georreferenciamento.docx
- memorial_descritivo.docx
- recibo_servico.docx
- contrato_dacao.docx
- procuracao.docx

### 5) Iniciar sistema

```bash
python app.py
```

Acesso local:

- http://localhost:5000

---

## Publicacao Externa (ngrok)

### Opcao A - Manual

1. Inicie o app:

```bash
python app.py
```

2. Em outro terminal:

```bash
ngrok http 5000
```

3. Use o link HTTPS exibido em `Forwarding`.

### Opcao B - Automatica (sem abrir VS Code)

Use o script:

- `iniciar_aledocs_publico.bat`

Ele abre duas janelas:

- AleDocs App (Flask)
- AleDocs Ngrok (tunel publico)

Validacao sem iniciar servicos:

```bat
iniciar_aledocs_publico.bat --check
```

---

## Estrutura do Projeto

```text
AleDocs/
├── app.py
├── create_templates.py
├── requirements.txt
├── iniciar_aledocs_publico.bat
├── README.md
│
├── instance/
│   └── aledocs.db
│
├── docx_templates/
│   ├── termo_entrega.docx
│   ├── checklist_averbacao.docx
│   ├── contrato_georreferenciamento.docx
│   ├── memorial_descritivo.docx
│   ├── recibo_servico.docx
│   ├── contrato_dacao.docx
│   └── procuracao.docx
│
├── templates/
│   ├── base.html
│   ├── index.html
│   ├── cliente_form.html
│   ├── listar_clientes.html
│   ├── gerar_docs.html
│   ├── emissor_form.html
│   └── listar_emissores.html
│
└── static/
    ├── css/style.css
    ├── js/main.js
    └── img/logo.png
```

---

## Documentos Gerados

| ID do modelo | Documento |
|---|---|
| termo_entrega | Termo de Entrega de Documentos |
| checklist_averbacao | Checklist de Averbacao de Georreferenciamento |
| contrato_georreferenciamento | Contrato de Georreferenciamento |
| memorial_descritivo | Memorial Descritivo |
| recibo_servico | Recibo de Prestacao de Servico |
| contrato_dacao | Contrato de Servicos Advocaticios e Tecnicos (Dacao em Pagamento) |
| procuracao | Procuracao Particular |

---

## Fluxo de Uso

1. Cadastre ou selecione um cliente.
2. Cadastre ou selecione um emissor.
3. Escolha os documentos na tela de geracao.
4. Clique em gerar para baixar o `.zip` com os `.docx` preenchidos.
5. Abra os arquivos no Word e finalize (assinatura/revisao).

---

## Campos e Preenchimento

- Emissor obrigatorio para gerar documentos: nome e CPF/CNPJ.
- Se algum campo opcional nao estiver preenchido, o sistema deixa espaco preenchivel para completar manualmente nos documentos.
- O campo de registro foi padronizado como `Registro` (sem fixar CREA/OAB).

---

## Interface e Responsividade

- Interface baseada em Bootstrap 5.
- Uso em desktop/notebook e mobile.
- Navbar, cards, formularios e tabelas com comportamento responsivo.

---

## Problemas Comuns

**Modelos `.docx` nao encontrados**

- Execute `python create_templates.py`.

**Porta 5000 ocupada**

- Ajuste a porta no final do `app.py`.

**ngrok nao abre tunel**

- Verifique versao com `ngrok version`.
- Atualize se necessario: `ngrok update`.

---

Desenvolvido com muito amor, carinho e cafe ☕❤️.
