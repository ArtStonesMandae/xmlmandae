# Atualizador de Chave NF-e para Planilhas Mandae

Este aplicativo desenvolvido em Streamlit permite atualizar automaticamente a coluna **"CHAVE NF"** da planilha Mandae (.xlsx), utilizando os arquivos XML de Notas Fiscais Eletrônicas (NF-e) recebidos em um arquivo `.zip`.

---

## ✅ Funcionalidades

- Upload da planilha Mandae no formato `.xlsx`
- Upload do arquivo `.zip` contendo os XMLs de NF-e
- Validação de documentos CPF e CNPJ
- Preenchimento automático da coluna "CHAVE NF"
- Detecção de pedidos duplicados por documento (tratamento manual)
- Geração de arquivo de saída com nome padronizado
- Interface amigável e intuitiva via Streamlit

---

## ▶️ Como rodar localmente

1. Clone o repositório:
```bash
git clone https://github.com/seuusuario/seurepositorio.git
cd seurepositorio
```

2. Crie um ambiente virtual (opcional, mas recomendado):
```bash
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
.venv\Scripts\activate   # Windows
```

3. Instale as dependências:
```bash
pip install -r requirements.txt
```

4. Execute o app:
```bash
streamlit run app_mandae.py
```

---

## 📦 Estrutura do projeto

```
📁 projeto/
├── app_mandae.py
├── requirements.txt
└── README.md
```

---

## ℹ️ Observações

- A contagem de pedidos considera **todas as linhas com CPF ou CNPJ válido**, mesmo que o documento se repita.
- Em caso de duplicidade de documentos, as chaves de acesso **não são preenchidas**. Essas linhas devem ser revisadas manualmente com base nas informações do pedido.
- O nome do arquivo de saída segue o padrão:  
  **`{total_pedidos}Pedidos - {DD.MM} - L2.xlsx`**

---

Desenvolvido por [ArtStones]
